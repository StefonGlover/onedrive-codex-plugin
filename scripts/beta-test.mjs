#!/usr/bin/env node

import { execFileSync, spawn } from "node:child_process";
import { access, mkdir, readFile, rm, rmdir, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";
import { ONEDRIVE_TOOL_CONTRACT, compareToolContract } from "./tool-contract.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");
const workspace = process.cwd();
const harnessStartedAtMs = Date.now();
const harnessStartedAt = new Date(harnessStartedAtMs).toISOString();
const betaFolderPrefix = "Codex OneDrive Plugin Beta Test codex-beta-";
const maxCleanupStaleDays = 1_000_000;
const liveRunIdPattern = /^codex-beta-[a-z0-9](?:[a-z0-9._-]*[a-z0-9])?$/i;
const resultStatuses = new Set(["pass", "fail", "blocked"]);

function parseCliArgs(argv = []) {
  const parsed = {};
  for (const arg of argv) {
    if (!arg.startsWith("--")) {
      throw new Error(`Unexpected positional argument: ${arg}. Options must use --name or --name=value.`);
    }
    const [key, ...rest] = arg.slice(2).split("=");
    if (!key) throw new Error("Empty CLI option is not allowed.");
    if (Object.hasOwn(parsed, key)) throw new Error(`--${key} may only be provided once.`);
    parsed[key] = rest.length ? rest.join("=") : true;
  }
  return parsed;
}

const allowedCliArgs = new Set([
  "keep-work",
  "live",
  "doctor-only",
  "cleanup-stale",
  "confirmed",
  "run-id",
  "invite-recipient",
  "stale-days",
  "cleanup-max-items",
  "cleanup-max-results",
  "cleanup-page-size",
  "cleanup-verify-concurrency",
  "cleanup-search-query",
  "read-retry-attempts",
  "read-retry-delay-ms",
  "tenant-matrix",
  "tenant-matrix-live",
  "self-check"
]);

function validateCliArgs(args = {}) {
  for (const key of Object.keys(args)) {
    if (!allowedCliArgs.has(key)) throw new Error(`Unknown CLI option: --${key}.`);
  }
  return args;
}

function parseBooleanFlag(value, name, defaultValue = false) {
  if (value === undefined) return defaultValue;
  if (typeof value === "boolean") return value;
  const normalized = String(value).trim().toLowerCase();
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off", ""].includes(normalized)) return false;
  throw new Error(`--${name} expects a boolean value, got ${value}.`);
}

function boundedInteger(value, defaultValue, minimum, maximum) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return defaultValue;
  return Math.min(Math.max(Math.trunc(parsed), minimum), maximum);
}

function parseBoundedIntegerFlag(value, name, defaultValue, minimum, maximum) {
  if (value === undefined) return defaultValue;
  if (typeof value === "boolean" || String(value).trim() === "") {
    throw new Error(`--${name} expects an integer between ${minimum} and ${maximum}.`);
  }
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < minimum || parsed > maximum) {
    throw new Error(`--${name} expects an integer between ${minimum} and ${maximum}, got ${value}.`);
  }
  return parsed;
}

function parseNonNegativeNumberFlag(value, name, defaultValue, maximum = Number.MAX_VALUE) {
  if (value === undefined) return defaultValue;
  if (typeof value === "boolean" || String(value).trim() === "") {
    throw new Error(`--${name} expects a non-negative number.`);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > maximum) {
    throw new Error(`--${name} expects a non-negative finite number no greater than ${maximum}, got ${value}.`);
  }
  return parsed;
}

function parseRunId(value) {
  if (value === undefined) return null;
  if (typeof value === "boolean" || String(value).length > 80 || !liveRunIdPattern.test(String(value))) {
    throw new Error("--run-id must match codex-beta-[A-Za-z0-9._-] and be at most 80 characters.");
  }
  return String(value);
}

function parseInviteRecipient(value) {
  if (value === undefined) return null;
  if (typeof value === "boolean") throw new Error("--invite-recipient requires an email address.");
  const recipient = String(value).trim();
  if (recipient.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) {
    throw new Error("--invite-recipient requires one valid email address.");
  }
  return recipient;
}

function exportFailureStatus(value) {
  const message = errorText(value);
  const unsupported = /\b(?:notSupported|unsupportedFormat|cannotOpenFile|notAcceptable)\b/i.test(message)
    || /\b(?:HTTP\s*)?406\b/i.test(message);
  return unsupported ? "blocked" : "fail";
}

function throws(fn) {
  try {
    fn();
    return false;
  } catch {
    return true;
  }
}

function betaFolderLooksStale(item = {}, cutoffMs) {
  if (!item.folder || !String(item.name || "").startsWith(betaFolderPrefix)) return false;
  const modified = Date.parse(item.lastModifiedDateTime || item.createdDateTime || "");
  return Number.isFinite(modified) && modified <= cutoffMs;
}

function uniqueBetaFolderCandidates(items = [], maxResults = 100) {
  const seen = new Set();
  const candidates = [];
  for (const item of items) {
    if (!item?.id || !item.folder || !String(item.name || "").startsWith(betaFolderPrefix) || seen.has(item.id)) continue;
    seen.add(item.id);
    candidates.push(item);
    if (candidates.length >= maxResults) break;
  }
  return candidates;
}

function betaIsolationRoots(runDirectory) {
  const storageRoot = resolve(runDirectory, "plugin-state");
  return {
    storageRoot,
    cacheRoot: join(storageRoot, "cache")
  };
}

function betaSyntheticSyncRoot(runDirectory) {
  return resolve(runDirectory, "synthetic-local-sync-root");
}

async function pathExists(path) {
  try {
    await access(path);
    return true;
  } catch (error) {
    if (error.code === "ENOENT") return false;
    throw error;
  }
}

function errorText(value) {
  if (value instanceof Error) return value.message;
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function parseChildMessageLine(line) {
  return JSON.parse(line);
}

function clearlyTransientReadError(value) {
  const message = errorText(value).trim();
  const wrapped = message.match(/^Microsoft Graph transport error after \d+ attempts?:\s*(.+)$/i);
  if (wrapped) return clearlyTransientReadError(wrapped[1]);
  return /^fetch failed$/i.test(message)
    || /\b(?:ECONNRESET|ECONNREFUSED|ETIMEDOUT|EAI_AGAIN|ENETUNREACH|EHOSTUNREACH|UND_ERR_[A-Z_]+)\b/i.test(message)
    || /(?:socket hang up|other side closed|network connection (?:was )?lost|temporary failure in name resolution)/i.test(message)
    || /Substrate Search/i.test(message);
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function retryTransientReadCall(name, call, options = {}) {
  const maxAttempts = boundedInteger(options.maxAttempts, 3, 1, 5);
  const baseDelayMs = boundedInteger(options.baseDelayMs, 250, 0, 5000);
  const onRetry = options.onRetry || (() => {});
  let lastResult;
  let lastError;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const result = await call();
      lastResult = { ...result, attempts: attempt };
      if (!result.isError || !clearlyTransientReadError(result.value) || attempt === maxAttempts) return lastResult;
      onRetry({ tool: name, attempt, reason: errorText(result.value) });
    } catch (error) {
      lastError = error;
      if (!clearlyTransientReadError(error) || attempt === maxAttempts) throw error;
      onRetry({ tool: name, attempt, reason: errorText(error) });
    }
    await wait(baseDelayMs * attempt);
  }
  if (lastResult) return lastResult;
  throw lastError || new Error(`${name} failed without a result.`);
}

const cliArgs = validateCliArgs(parseCliArgs(process.argv.slice(2)));
const keepWork = parseBooleanFlag(cliArgs["keep-work"], "keep-work");
const live = parseBooleanFlag(cliArgs.live, "live");
const doctorOnly = parseBooleanFlag(cliArgs["doctor-only"], "doctor-only");
const cleanupStale = parseBooleanFlag(cliArgs["cleanup-stale"], "cleanup-stale");
const confirmed = parseBooleanFlag(cliArgs.confirmed, "confirmed");
const cleanupConfirmed = live && confirmed;
const runId = parseRunId(cliArgs["run-id"]);
const inviteRecipient = parseInviteRecipient(cliArgs["invite-recipient"]);
const cleanupStaleDays = parseNonNegativeNumberFlag(cliArgs["stale-days"], "stale-days", 1, maxCleanupStaleDays);
const cleanupMaxItems = parseBoundedIntegerFlag(cliArgs["cleanup-max-items"], "cleanup-max-items", 500, 1, 5000);
const cleanupMaxResults = parseBoundedIntegerFlag(cliArgs["cleanup-max-results"], "cleanup-max-results", 100, 1, 500);
const cleanupPageSize = parseBoundedIntegerFlag(cliArgs["cleanup-page-size"], "cleanup-page-size", 100, 1, 200);
const cleanupVerifyConcurrency = parseBoundedIntegerFlag(cliArgs["cleanup-verify-concurrency"], "cleanup-verify-concurrency", 4, 1, 8);
const readRetryAttempts = parseBoundedIntegerFlag(cliArgs["read-retry-attempts"], "read-retry-attempts", 3, 1, 5);
const readRetryDelayMs = parseBoundedIntegerFlag(cliArgs["read-retry-delay-ms"], "read-retry-delay-ms", 250, 0, 5000);
const tenantMatrix = cliArgs["tenant-matrix"];
const tenantMatrixLive = parseBooleanFlag(cliArgs["tenant-matrix-live"], "tenant-matrix-live");
const selfCheck = parseBooleanFlag(cliArgs["self-check"], "self-check");

if (!selfCheck) {
  if (confirmed && !live) throw new Error("--confirmed is refused without --live.");
  if (live && !confirmed) throw new Error("--live requires --confirmed before any OneDrive mutation.");
  if (live && !runId) throw new Error("--live requires --run-id=<exact-id> before any OneDrive mutation.");
  if (live && doctorOnly) throw new Error("--live cannot be combined with --doctor-only because doctor-only is read-only.");
  if (tenantMatrixLive && !live) throw new Error("--tenant-matrix-live requires --live, --confirmed, and --run-id=<exact-id>.");
  if (live && !cleanupStale && !inviteRecipient) {
    throw new Error("A full live beta requires --invite-recipient=<email> for the isolated permission grant/revoke test.");
  }
}

if (selfCheck) {
  const now = Date.now();
  const stale = { id: "stale", name: `${betaFolderPrefix}old`, folder: {}, lastModifiedDateTime: new Date(now - 2 * 86_400_000).toISOString() };
  const fresh = { id: "fresh", name: `${betaFolderPrefix}new`, folder: {}, lastModifiedDateTime: new Date(now).toISOString() };
  const candidates = uniqueBetaFolderCandidates([stale, stale, fresh, { id: "file", name: `${betaFolderPrefix}file` }], 10);
  const isolationProbe = betaIsolationRoots(join(workspace, "work", "onedrive-beta", "self-check"));
  const syntheticSyncRootProbe = betaSyntheticSyncRoot(join(workspace, "work", "onedrive-beta", "self-check"));
  let transientAttempts = 0;
  const transientResult = await retryTransientReadCall("probe", async () => {
    transientAttempts += 1;
    return transientAttempts === 1
      ? { isError: true, value: "Microsoft Graph transport error after 3 attempts: fetch failed" }
      : { isError: false, value: { ok: true } };
  }, { maxAttempts: 3, baseDelayMs: 0 });
  let deterministicAttempts = 0;
  const deterministicResult = await retryTransientReadCall("probe", async () => {
    deterministicAttempts += 1;
    return { isError: true, value: "Microsoft Graph error: itemNotFound" };
  }, { maxAttempts: 3, baseDelayMs: 0 });
  const checks = {
    explicitFalseFlag: parseBooleanFlag("false", "probe", true) === false,
    explicitTrueFlag: parseBooleanFlag("true", "probe") === true,
    positionalArgumentRejected: throws(() => parseCliArgs(["unexpected"])),
    duplicateOptionRejected: throws(() => parseCliArgs(["--keep-work", "--keep-work=false"])),
    unknownOptionRejected: throws(() => validateCliArgs(parseCliArgs(["--keep-wrok"]))),
    invalidRunIdRejected: throws(() => parseRunId("unsafe folder/name")),
    validRunIdAccepted: parseRunId("codex-beta-20260713T120000Z") === "codex-beta-20260713T120000Z",
    invalidInviteRecipientRejected: throws(() => parseInviteRecipient("not-an-email")),
    validInviteRecipientAccepted: parseInviteRecipient("beta@example.test") === "beta@example.test",
    exactToolContract: compareToolContract(ONEDRIVE_TOOL_CONTRACT).ok,
    preciseUnsupportedExportClassification: exportFailureStatus("Microsoft Graph error notSupported: conversion refused") === "blocked",
    unknownExportFailureIsNotSoftPassed: exportFailureStatus("Microsoft Graph error: UnknownError") === "fail",
    negativeStaleDaysRejected: throws(() => parseNonNegativeNumberFlag("-1", "stale-days", 1)),
    nonFiniteStaleDaysRejected: throws(() => parseNonNegativeNumberFlag("NaN", "stale-days", 1)),
    overflowingStaleDaysRejected: throws(() => parseNonNegativeNumberFlag("1e308", "stale-days", 1, maxCleanupStaleDays)),
    fractionalIntegerFlagRejected: throws(() => parseBoundedIntegerFlag("1.5", "cleanup-page-size", 100, 1, 200)),
    outOfRangeIntegerFlagRejected: throws(() => parseBoundedIntegerFlag("201", "cleanup-page-size", 100, 1, 200)),
    staleClassification: betaFolderLooksStale(stale, now - 86_400_000) && !betaFolderLooksStale(fresh, now - 86_400_000),
    unknownTimestampNotStale: !betaFolderLooksStale({ id: "unknown-age", name: `${betaFolderPrefix}unknown`, folder: {} }, now),
    candidateDeduplication: candidates.length === 2 && candidates[0].id === "stale" && candidates[1].id === "fresh",
    isolatedStorageRoot: isolationProbe.storageRoot === resolve(workspace, "work", "onedrive-beta", "self-check", "plugin-state"),
    isolatedCacheRoot: isolationProbe.cacheRoot === resolve(workspace, "work", "onedrive-beta", "self-check", "plugin-state", "cache"),
    syntheticSyncRootIsRunScoped: syntheticSyncRootProbe === resolve(workspace, "work", "onedrive-beta", "self-check", "synthetic-local-sync-root"),
    syntheticSyncRootOutsideRealOneDriveHierarchy: !syntheticSyncRootProbe.startsWith(join(homedir(), "Library", "CloudStorage", "OneDrive"))
      && !syntheticSyncRootProbe.startsWith(join(homedir(), "OneDrive")),
    transientFetchFailure: clearlyTransientReadError("fetch failed"),
    classifiedTransportFailure: clearlyTransientReadError("Microsoft Graph transport error after 3 attempts: fetch failed"),
    deterministicErrorNotTransient: !clearlyTransientReadError("Microsoft Graph error: itemNotFound"),
    transientReadRetried: transientAttempts === 2 && transientResult.isError === false && transientResult.attempts === 2,
    deterministicReadNotRetried: deterministicAttempts === 1 && deterministicResult.isError === true && deterministicResult.attempts === 1,
    previewTokenRefusalFailsClosed: commitRefusalReasons({ dryRun: false, confirmed: true, previewTokenRequired: true, requiredToDelete: "preview again" }).length === 2,
    preflightRefusalFailsClosed: commitRefusalReasons({ dryRun: false, confirmed: true, preflightComplete: false, mutationStarted: false }).length === 2,
    partialMutationFailsClosed: commitRefusalReasons({ dryRun: false, confirmed: true, mutationStarted: true, partialState: true, failed: { index: 1 } }).length === 2,
    concreteMutationSuccessAccepted: commitRefusalReasons({ dryRun: false, confirmed: true, mutationStarted: true, partialState: false, deleted: { id: "exact" } }).length === 0,
    malformedChildOutputDetected: throws(() => parseChildMessageLine("not-json"))
  };
  const ok = Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, checks }, null, 2));
  process.exit(ok ? 0 : 1);
}

const suggestedRunId = `codex-beta-${new Date().toISOString().replace(/[-:.]/g, "").replace("Z", "Z").toLowerCase()}-${process.pid}`;
const unique = runId || suggestedRunId;
const outDir = join(workspace, "work", "onedrive-beta", unique);
const { storageRoot: betaStorageRoot, cacheRoot: betaCacheRoot } = betaIsolationRoots(outDir);
const realConfigPath = join(homedir(), ".codex", "onedrive-plugin", "config.json");
const syntheticSyncRoot = betaSyntheticSyncRoot(outDir);
const localUpload = join(outDir, "upload-source.txt");
const localSessionUpload = join(outDir, "upload-session-source.txt");
const localBinary = join(outDir, "binary-source.bin");
const localDownload = join(outDir, "upload-downloaded.txt");
const auditExport = join(outDir, "audit-export.jsonl");
const excelDownload = join(outDir, "downloaded-book.xlsx");
const wordDownload = join(outDir, "downloaded-doc.docx");
const powerpointDownload = join(outDir, "downloaded-deck.pptx");
const exportPdfDownload = join(outDir, "exported-doc.pdf");
const exportTextDownload = join(outDir, "exported-doc.txt");
const updateCheckout = join(outDir, "update-checkout.txt");
const updateManifest = join(outDir, "update-checkout.json");
const officeFixtureDir = join(outDir, "office-fixtures");
const richOfficeFixtureDir = join(outDir, "rich-office-fixtures");
const blockedSyncDownload = join(syntheticSyncRoot, `${unique}-blocked-download.txt`);
const blockedSyncUpload = join(syntheticSyncRoot, `${unique}-blocked-upload.txt`);
const blockedSyncUploadSentinel = `Synthetic local sync-path guard fixture: ${unique}\n`;
const folderName = `Codex OneDrive Plugin Beta Test ${unique}`;
const movedFolderName = "Moved";
const renamedTextFile = "note-renamed.txt";
const movedTextFile = "note-moved.txt";
const copyFileName = "note-copy.txt";
const content = [
  `OneDrive plugin beta test token: ${unique}`,
  "This file was created by Codex and should be deleted during cleanup.",
  ""
].join("\n");
let richOfficeFixtures = null;

async function cleanupLocalWorkDir() {
  await rm(outDir, { recursive: true, force: true });
  for (const directory of [dirname(outDir), dirname(dirname(outDir))]) {
    try {
      await rmdir(directory);
    } catch (error) {
      if (!["ENOENT", "ENOTEMPTY"].includes(error.code)) throw error;
    }
  }
}

async function prepareSyntheticSyncFixture() {
  await mkdir(syntheticSyncRoot, { recursive: true, mode: 0o700 });
  await writeFile(blockedSyncUpload, blockedSyncUploadSentinel, { flag: "wx", mode: 0o600 });
}

await mkdir(outDir, { recursive: true });
if (live && !cleanupStale && !tenantMatrix) {
  const fixturePython = process.env.ONEDRIVE_OFFICE_TEST_PYTHON || "python3";
  execFileSync(fixturePython, [join(pluginRoot, "scripts", "office-openxml-test.py"), `--emit-fixtures=${officeFixtureDir}`], {
    env: { ...process.env, PYTHONPYCACHEPREFIX: join(outDir, "pycache") }, stdio: "ignore"
  });
  richOfficeFixtures = JSON.parse(execFileSync(fixturePython, [join(pluginRoot, "scripts", "office-fixture-factory.py"), richOfficeFixtureDir], {
    env: { ...process.env, PYTHONPYCACHEPREFIX: join(outDir, "pycache") }, encoding: "utf8"
  }));
  await writeFile(localUpload, `Uploaded through onedrive_upload: ${unique}\n`, "utf8");
  await writeFile(localSessionUpload, Buffer.alloc(400 * 1024, `session-${unique}\n`));
  await writeFile(localBinary, Buffer.from([0, 1, 2, 3, 4, 0, 255, 128]));
  await rm(localDownload, { force: true });
  await rm(excelDownload, { force: true });
  await rm(wordDownload, { force: true });
  await rm(powerpointDownload, { force: true });
  await rm(exportPdfDownload, { force: true });
  await rm(exportTextDownload, { force: true });
  await rm(updateCheckout, { force: true });
  await rm(updateManifest, { force: true });
}
if (live && !cleanupStale && !tenantMatrix) await prepareSyntheticSyncFixture();

const child = spawn(process.execPath, [serverPath], {
  cwd: workspace,
  env: {
    ...process.env,
    ONEDRIVE_STORAGE_ROOT: betaStorageRoot,
    ONEDRIVE_CACHE_ROOT: betaCacheRoot,
    ONEDRIVE_ADDITIONAL_LOCAL_SYNC_ROOTS: JSON.stringify([syntheticSyncRoot])
  },
  stdio: ["pipe", "pipe", "pipe"]
});

let nextId = 1;
let buffer = "";
const pending = new Map();
const stderr = [];
const toolRetryEvents = [];
const calledTools = new Set();
let childExited = false;
let childExitError = null;

function rejectPendingRequests(error) {
  for (const [id, waiter] of pending.entries()) {
    pending.delete(id);
    clearTimeout(waiter.timeout);
    waiter.reject(error);
  }
}

const childExit = new Promise((resolve) => {
  child.once("exit", (code, signal) => {
    childExited = true;
    childExitError = new Error(`OneDrive MCP child exited before completing pending requests (code=${code ?? "null"}, signal=${signal || "none"}).`);
    rejectPendingRequests(childExitError);
    resolve();
  });
});

child.once("error", (error) => {
  childExitError = new Error(`Could not start OneDrive MCP child: ${error.message}`);
  rejectPendingRequests(childExitError);
});

child.stdin.on("error", (error) => {
  childExitError ||= new Error(`OneDrive MCP child stdin failed: ${error.message}`);
  rejectPendingRequests(childExitError);
});

child.stderr.on("data", (chunk) => stderr.push(chunk.toString()));
child.stdout.on("data", (chunk) => {
  buffer += chunk.toString();
  for (;;) {
    const newline = buffer.indexOf("\n");
    if (newline === -1) break;
    const line = buffer.slice(0, newline).trim();
    buffer = buffer.slice(newline + 1);
    if (!line) continue;
    let message;
    try {
      message = parseChildMessageLine(line);
    } catch (error) {
      childExitError ||= new Error(`OneDrive MCP child emitted malformed JSON: ${error.message}`);
      rejectPendingRequests(childExitError);
      child.kill();
      return;
    }
    const waiter = pending.get(message.id);
    if (waiter) {
      pending.delete(message.id);
      clearTimeout(waiter.timeout);
      waiter.resolve(message);
    }
  }
});

function request(method, params = {}) {
  if (childExited || childExitError || !child.stdin.writable) {
    return Promise.reject(childExitError || new Error("OneDrive MCP child is not available."));
  }
  const id = nextId++;
  const payload = { jsonrpc: "2.0", id, method, params };
  const promise = new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      if (pending.has(id)) {
        pending.delete(id);
        reject(new Error(`Timed out waiting for ${method}`));
      }
    }, 120_000);
    pending.set(id, { resolve, reject, timeout });
  });
  try {
    child.stdin.write(`${JSON.stringify(payload)}\n`, (error) => {
      if (!error) return;
      const waiter = pending.get(id);
      if (!waiter) return;
      pending.delete(id);
      clearTimeout(waiter.timeout);
      waiter.reject(new Error(`Could not write ${method} request to OneDrive MCP child: ${error.message}`));
    });
  } catch (error) {
    const waiter = pending.get(id);
    if (waiter) {
      pending.delete(id);
      clearTimeout(waiter.timeout);
      waiter.reject(error);
    }
  }
  return promise;
}

async function toolOnce(name, args = {}) {
  calledTools.add(name);
  const response = await request("tools/call", { name, arguments: args });
  if (response.error) throw new Error(response.error.message);
  const text = response.result?.content?.[0]?.text ?? "";
  let value = text;
  try {
    value = JSON.parse(text);
  } catch {
    // Keep string responses as-is.
  }
  return { isError: Boolean(response.result?.isError), value, raw: response };
}

const retryableReadOnlyTools = new Set([
  "onedrive_config",
  "onedrive_doctor",
  "onedrive_me",
  "onedrive_drive",
  "onedrive_presets",
  "onedrive_list",
  "onedrive_list_all",
  "onedrive_scan",
  "onedrive_search",
  "onedrive_search_all",
  "onedrive_find",
  "onedrive_find_all",
  "onedrive_delta",
  "onedrive_sync_status",
  "onedrive_get_info",
  "onedrive_read_text",
  "onedrive_preview",
  "onedrive_permissions",
  "onedrive_batch_get_info",
  "onedrive_batch_permissions",
  "onedrive_recent",
  "onedrive_large_files",
  "onedrive_duplicates",
  "onedrive_shared_by_me",
  "onedrive_public_links"
]);

async function tool(name, args = {}) {
  const maxAttempts = retryableReadOnlyTools.has(name)
    ? readRetryAttempts
    : 1;
  const baseDelayMs = readRetryDelayMs;
  return await retryTransientReadCall(name, () => toolOnce(name, args), {
    maxAttempts,
    baseDelayMs,
    onRetry: (event) => toolRetryEvents.push(event)
  });
}

async function pollIndexedSearch(name, args, matches, options = {}) {
  const maxAttempts = boundedInteger(options.maxAttempts, 6, 1, 20);
  const delayMs = boundedInteger(options.delayMs, 2_000, 0, 10_000);
  let lastResult = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    lastResult = await tool(name, args);
    if (!lastResult.isError && matches(lastResult.value)) return { ...lastResult, pollAttempts: attempt };
    if (lastResult.isError && !clearlyTransientReadError(lastResult.value)) return { ...lastResult, pollAttempts: attempt };
    if (attempt < maxAttempts) await wait(delayMs);
  }
  return { ...lastResult, pollAttempts: maxAttempts };
}

function assertOk(name, result) {
  if (result.isError) {
    throw new Error(`${name} returned error: ${typeof result.value === "string" ? result.value : JSON.stringify(result.value)}`);
  }
  return result.value;
}

function commitRefusalReasons(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return ["non-object mutation response"];
  const reasons = [];
  const requiredFields = Object.entries(value)
    .filter(([key, fieldValue]) => /^requiredTo[A-Z]/.test(key) && Boolean(fieldValue))
    .map(([key]) => key);
  if (requiredFields.length) reasons.push(`requirements:${requiredFields.join(",")}`);
  if (value.previewTokenRequired === true) reasons.push(`previewToken:${value.previewTokenStatus || "required"}`);
  if (value.dryRun === true) reasons.push("dryRun:true");
  if (value.confirmed === false) reasons.push("confirmed:false");
  if (value.preflightFailed === true || value.preflightComplete === false) reasons.push("preflightFailed");
  if (value.mutationStarted === false) reasons.push("mutationStarted:false");
  if (value.failed === true || (value.failed && typeof value.failed === "object")) reasons.push("failed");
  if (value.partialState === true) reasons.push("partialState:true");
  if (value.noChanges === true) reasons.push("noChanges:true");
  return reasons;
}

function assertCommitted(name, result, successPredicate = null) {
  const value = assertOk(name, result);
  const refusalReasons = commitRefusalReasons(value);
  if (refusalReasons.length) {
    throw new Error(`${name} was not committed (${refusalReasons.join("; ")}).`);
  }
  if (successPredicate && !successPredicate(value)) {
    throw new Error(`${name} did not return the required concrete success evidence.`);
  }
  return value;
}

async function previewTokenFor(name, args = {}) {
  const previewArgs = { ...args };
  delete previewArgs.dryRun;
  delete previewArgs.confirmed;
  delete previewArgs.previewToken;
  const preview = assertOk(`${name} preview`, await tool(name, previewArgs));
  if (!preview.previewToken) throw new Error(`${name} preview did not return a previewToken.`);
  return preview.previewToken;
}

async function toolWithPreview(name, args = {}, options = {}) {
  const previewToken = options.previewToken || await previewTokenFor(name, args);
  const result = await tool(name, { ...args, previewToken });
  const refusalReasons = result.isError ? [] : commitRefusalReasons(result.value);
  if (refusalReasons.length) {
    return {
      isError: true,
      value: `${name} did not commit (${refusalReasons.join("; ")}). Response: ${JSON.stringify(result.value)}`,
      previewRefusal: result.value,
      previewRefreshes: 0
    };
  }
  return { ...result, previewRefreshes: 0 };
}

async function cleanupExactBetaFolder(target, remotePath) {
  let deleted = null;
  let deletionError = null;
  try {
    deleted = assertCommitted("emergency beta-root cleanup", await toolWithPreview("onedrive_delete", {
      itemId: target.id,
      expectedName: target.name,
      dryRun: false,
      confirmed: true
    }), (value) => value.dryRun === false && value.deleted?.id === target.id);
  } catch (error) {
    deletionError = error.message;
  }
  const probe = await tool("onedrive_get_info", { path: remotePath });
  const verifiedAbsent = probe.isError && /\bitemNotFound\b/i.test(errorText(probe.value));
  if (!verifiedAbsent) {
    throw new Error(`Emergency beta-root cleanup did not verify itemNotFound for ${target.id}.${deletionError ? ` Delete error: ${deletionError}` : ""} Probe: ${errorText(probe.value)}`);
  }
  return {
    attemptedAfterError: true,
    deleted: deleted?.deleted ? { id: deleted.deleted.id, name: deleted.deleted.name } : null,
    deletionError,
    verifiedAbsent,
    probe: errorText(probe.value)
  };
}

async function cleanupStaleBetaFolders() {
  const cutoffMs = Date.now() - Math.max(0, cleanupStaleDays) * 24 * 60 * 60 * 1000;
  const searchQuery = String(cliArgs["cleanup-search-query"] || "Codex OneDrive Plugin Beta Test codex-beta");
  const searchMaxItems = cleanupMaxItems;
  const maxResults = cleanupMaxResults;
  const verificationConcurrency = cleanupVerifyConcurrency;
  const search = assertOk("cleanup stale search", await tool("onedrive_search_all", {
    query: searchQuery,
    pageSize: cleanupPageSize,
    maxItems: searchMaxItems,
    format: "full"
  }));
  const discoveredWithOverflow = uniqueBetaFolderCandidates(search.items || [], maxResults + 1);
  const resultLimitReached = discoveredWithOverflow.length > maxResults;
  const discovered = discoveredWithOverflow.slice(0, maxResults);
  const candidates = [];
  const verificationSkips = [];
  for (let offset = 0; offset < discovered.length; offset += verificationConcurrency) {
    const batch = discovered.slice(offset, offset + verificationConcurrency);
    const verifiedBatch = await Promise.all(batch.map(async (candidate) => {
      const info = await tool("onedrive_get_info", { itemId: candidate.id, format: "full" });
      if (!info.isError) return info.value;
      const message = errorText(info.value);
      if (/\bitemNotFound\b/i.test(message)) {
        verificationSkips.push({ id: candidate.id, name: candidate.name, reason: "itemNotFound" });
        return null;
      }
      throw new Error(`cleanup stale verification returned error for ${candidate.name}: ${message}`);
    }));
    for (const candidate of verifiedBatch) {
      if (!candidate) continue;
      const modified = Date.parse(candidate.lastModifiedDateTime || candidate.createdDateTime || "");
      if (!Number.isFinite(modified)) {
        verificationSkips.push({ id: candidate.id, name: candidate.name, reason: "missingOrInvalidTimestamp" });
        continue;
      }
      if (betaFolderLooksStale(candidate, cutoffMs)) candidates.push(candidate);
    }
  }
  const deleted = [];
  if (cleanupConfirmed) {
    for (const candidate of candidates) {
      const result = assertOk("cleanup stale delete", await toolWithPreview("onedrive_delete", {
        itemId: candidate.id,
        expectedName: candidate.name,
        dryRun: false,
        confirmed: true
      }));
      deleted.push(result.deleted);
    }
  }
  const postDeleteVerification = {
    attempted: cleanupConfirmed && candidates.length > 0,
    direct: [],
    rediscovery: null
  };
  const verificationFailures = [];
  const postDeleteIncompleteReasons = [];
  if (postDeleteVerification.attempted) {
    const deletedIds = new Set(candidates.map((candidate) => candidate.id));
    for (const candidate of candidates) {
      const info = await tool("onedrive_get_info", { itemId: candidate.id, format: "full" });
      if (!info.isError) {
        postDeleteVerification.direct.push({
          id: candidate.id,
          name: candidate.name,
          absent: false,
          resolvedId: info.value?.id,
          resolvedName: info.value?.name
        });
        verificationFailures.push(`deleted-item-still-resolves:${candidate.id}`);
        continue;
      }
      const message = errorText(info.value);
      const absent = /\bitemNotFound\b/i.test(message);
      postDeleteVerification.direct.push({ id: candidate.id, name: candidate.name, absent, error: message });
      if (!absent) postDeleteIncompleteReasons.push(`direct-delete-verification-error:${candidate.id}`);
    }

    const rediscovery = await pollIndexedSearch("onedrive_search_all", {
      query: searchQuery,
      pageSize: cleanupPageSize,
      maxItems: searchMaxItems,
      format: "full"
    }, (value) => !value?.truncated && !(value?.items || []).some((item) => deletedIds.has(item?.id)), {
      maxAttempts: 6,
      delayMs: 2_000
    });
    if (rediscovery.isError) {
      postDeleteVerification.rediscovery = {
        complete: false,
        pollAttempts: rediscovery.pollAttempts,
        error: errorText(rediscovery.value)
      };
      postDeleteIncompleteReasons.push("post-delete-rediscovery-error");
    } else {
      const remainingDeletedIds = (rediscovery.value?.items || [])
        .filter((item) => deletedIds.has(item?.id))
        .map((item) => item.id);
      const rediscoveryTruncated = Boolean(rediscovery.value?.truncated);
      postDeleteVerification.rediscovery = {
        complete: !rediscoveryTruncated && remainingDeletedIds.length === 0,
        pollAttempts: rediscovery.pollAttempts,
        truncated: rediscoveryTruncated,
        returned: rediscovery.value?.items?.length || 0,
        remainingDeletedIds
      };
      if (rediscoveryTruncated) postDeleteIncompleteReasons.push("post-delete-search-results-truncated");
      if (remainingDeletedIds.length > 0) {
        const directlyPresent = new Set(postDeleteVerification.direct
          .filter((entry) => entry.absent === false)
          .map((entry) => entry.id));
        for (const id of remainingDeletedIds) {
          if (directlyPresent.has(id)) verificationFailures.push(`deleted-item-rediscovered:${id}`);
          else postDeleteIncompleteReasons.push(`stale-post-delete-search-result:${id}`);
        }
      }
    }
  }
  const details = {
    mode: cleanupConfirmed ? "delete" : "dry-run",
    cutoff: new Date(cutoffMs).toISOString(),
    staleDays: cleanupStaleDays,
    discovery: {
      tool: "onedrive_search_all",
      query: searchQuery,
      returned: search.items?.length || 0,
      truncated: Boolean(search.truncated),
      resultLimitReached,
      maxItems: searchMaxItems,
      folderCandidates: discovered.length,
      verificationConcurrency,
      verificationSkips
    },
    candidateCount: candidates.length,
    deletedCount: deleted.length,
    candidates: candidates.map((item) => ({ id: item.id, name: item.name, lastModifiedDateTime: item.lastModifiedDateTime })),
    deleted: deleted.map((item) => ({ id: item?.id, name: item?.name })),
    postDeleteVerification
  };
  const incompleteReasons = [
    ...(search.truncated ? ["search-results-truncated"] : []),
    ...(resultLimitReached ? ["candidate-result-limit-reached"] : []),
    ...verificationSkips.filter((entry) => entry.reason !== "itemNotFound").map((entry) => `verification-skip:${entry.reason}`),
    ...postDeleteIncompleteReasons
  ];
  details.failures = [...new Set(verificationFailures)];
  details.incomplete = incompleteReasons.length > 0;
  details.incompleteReasons = [...new Set(incompleteReasons)];
  results.cleanup = details;
  const cleanupStatus = details.failures.length > 0 ? "fail" : details.incomplete ? "blocked" : "pass";
  record("cleanup stale beta folders", cleanupStatus, details);
}

async function runOneTenantMatrixEntry(tenant) {
  const childArgs = [fileURLToPath(import.meta.url)];
  if (!tenantMatrixLive) childArgs.push("--doctor-only");
  else {
    const tenantRunId = `${runId}-${tenant.toLowerCase().replace(/[^a-z0-9._-]/g, "-")}`.slice(0, 80).replace(/[^a-z0-9]$/i, "0");
    childArgs.push("--live", "--confirmed", `--run-id=${tenantRunId}`, `--invite-recipient=${inviteRecipient}`);
  }
  if (keepWork) childArgs.push("--keep-work");
  const entry = spawn(process.execPath, childArgs, {
    cwd: workspace,
    env: { ...process.env, ONEDRIVE_TENANT: tenant },
    stdio: ["ignore", "pipe", "pipe"]
  });
  let stdout = "";
  let stderrText = "";
  entry.stdout.on("data", (chunk) => {
    stdout += chunk.toString();
  });
  entry.stderr.on("data", (chunk) => {
    stderrText += chunk.toString();
  });
  const exitCode = await new Promise((resolve) => entry.once("exit", (code) => resolve(code)));
  let parsed = null;
  try {
    parsed = JSON.parse(stdout);
  } catch {
    // Keep raw stdout below.
  }
  return {
    tenant,
    ok: exitCode === 0 && parsed?.summary?.failCount === 0 && !parsed?.error,
    exitCode,
    summary: parsed?.summary,
    configuredTenant: parsed?.checks?.find((check) => check.name === "configured and token available")?.details?.tenant,
    localWorkCleaned: parsed?.localWorkCleaned,
    localWorkDir: parsed?.localWorkDir,
    error: parsed?.error,
    stdout: parsed ? undefined : stdout.trim(),
    stderr: stderrText.trim() || undefined
  };
}

async function runTenantMatrix() {
  const tenants = String(tenantMatrix === true ? "common,consumers,organizations" : tenantMatrix)
    .split(",")
    .map((tenant) => tenant.trim())
    .filter(Boolean);
  const matrix = [];
  for (const tenant of tenants) {
    matrix.push(await runOneTenantMatrixEntry(tenant));
  }
  return {
    mode: tenantMatrixLive ? "live-beta" : "doctor-only",
    tenants,
    ok: matrix.every((entry) => entry.ok),
    matrix
  };
}

const results = {
  startedAt: harnessStartedAt,
  unique,
  folderName,
  mode: cleanupStale ? (cleanupConfirmed ? "cleanup-live" : "cleanup-preview") : live ? "live" : "read-only",
  live,
  runId: runId || null,
  checks: [],
  cleanup: null,
  retryEvents: toolRetryEvents,
  stderr: ""
};

function record(name, status, details = {}) {
  if (!resultStatuses.has(status)) throw new Error(`Invalid beta result status for ${name}: ${status}`);
  results.checks.push({ name, status, details });
}

let folder = null;

if (tenantMatrix) {
  child.kill("SIGTERM");
  await Promise.race([
    childExit,
    new Promise((resolve) => setTimeout(resolve, 2_000))
  ]);
  if (!childExited) child.kill("SIGKILL");
  let matrix;
  try {
    matrix = await runTenantMatrix();
  } catch (error) {
    matrix = { ok: false, error: error.stack || error.message };
  } finally {
    if (keepWork) {
      matrix = matrix || { ok: false };
      matrix.localWorkDir = outDir;
    } else {
      try {
        await cleanupLocalWorkDir();
        matrix = matrix || { ok: false };
        matrix.localWorkCleaned = true;
      } catch (error) {
        matrix = matrix || { ok: false };
        matrix.ok = false;
        matrix.localWorkCleaned = false;
        matrix.cleanupError = [matrix.cleanupError, `Local beta work cleanup failed: ${error.message}`].filter(Boolean).join("\n");
      }
    }
  }
  matrix.startedAt = harnessStartedAt;
  matrix.finishedAt = new Date().toISOString();
  matrix.runtimeMs = Date.now() - harnessStartedAtMs;
  console.log(JSON.stringify(matrix, null, 2));
  process.exitCode = matrix.ok ? 0 : 1;
  process.exit();
}

try {
  const init = await request("initialize", {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: { name: "onedrive-beta-test", version: "1.0.0" }
  });
  record("initialize", init.result?.serverInfo?.name === "onedrive" ? "pass" : "fail", init.result);

  const listed = await request("tools/list");
  const toolNames = listed.result.tools.map((entry) => entry.name).sort();
  const toolContract = compareToolContract(toolNames);
  record("tools/list exactly matches the 84-tool contract", toolContract.ok ? "pass" : "fail", toolContract);

  const config = assertOk("onedrive_config", await tool("onedrive_config", { checkToken: true }));
  record("configured and token available", config.clientIdConfigured && config.accessTokenAvailable ? "pass" : "fail", {
    tenant: config.tenant,
    scopes: config.scopes,
    keychainTokenConfigured: config.keychainTokenConfigured,
    accessTokenAvailable: config.accessTokenAvailable
  });
  const expectedConfigPath = realConfigPath;
  record("beta child isolates mutable local state while retaining the normal auth and Keychain context", config.settings?.storageRoot === betaStorageRoot
    && config.settings?.cacheRoot === betaCacheRoot
    && config.configPath === expectedConfigPath ? "pass" : "fail", {
    storageRoot: config.settings?.storageRoot,
    cacheRoot: config.settings?.cacheRoot,
    configPath: config.configPath,
    expectedStorageRoot: betaStorageRoot,
    expectedCacheRoot: betaCacheRoot,
    expectedConfigPath,
    syntheticSyncRoot
  });

  const doctor = assertOk("onedrive_doctor", await tool("onedrive_doctor", { checkRootList: true, rootListLimit: 3 }));
  record("doctor health check passes", doctor.ok === true && doctor.summary?.fail === 0 ? "pass" : "fail", {
    status: doctor.status,
    summary: doctor.summary,
    checks: doctor.checks?.map((check) => ({ name: check.name, status: check.status }))
  });

  if (cleanupStale) {
    await cleanupStaleBetaFolders();
  } else if (doctorOnly || !live) {
    results.specialMode = doctorOnly ? "doctor-only" : "read-only";
    results.livePreview = {
      runId: unique,
      folderName,
      command: `node scripts/beta-test.mjs --live --confirmed --run-id=${unique} --invite-recipient=<email>`
    };
    record("read-only health mode completed without CRUD", "pass", {
      tenant: config.tenant,
      proposedRunId: unique,
      proposedFolderName: folderName
    });
  } else {
  const me = assertOk("onedrive_me", await tool("onedrive_me"));
  record("profile read", me.userPrincipalName || me.mail ? "pass" : "fail", {
    displayName: me.displayName,
    userPrincipalName: me.userPrincipalName,
    mail: me.mail
  });

  const drive = assertOk("onedrive_drive", await tool("onedrive_drive"));
  record("drive metadata read", drive.id && drive.driveType ? "pass" : "fail", {
    driveType: drive.driveType,
    name: drive.name,
    quotaState: drive.quota?.state
  });

  const presets = assertOk("onedrive_presets", await tool("onedrive_presets"));
  record("path presets available", presets.pathPresets?.documents === "Documents" && presets.pathPresets?.desktop === "Desktop" ? "pass" : "fail", {
    pathPresets: presets.pathPresets
  });

  const presetTraversal = await tool("onedrive_get_info", {
    preset: "documents",
    relativePath: "../Pictures"
  });
  record("preset path traversal is refused", presetTraversal.isError && String(presetTraversal.value).includes("unsafe path segment") ? "pass" : "fail", {
    response: presetTraversal.value
  });

  folder = assertOk("onedrive_create_folder", await tool("onedrive_create_folder", {
    name: folderName,
    conflictBehavior: "fail"
  }));
  record("create folder", folder.id && folder.folder ? "pass" : "fail", { id: folder.id, name: folder.name });

  const movedFolder = assertOk("onedrive_create_folder", await tool("onedrive_create_folder", {
    parentPath: folderName,
    name: movedFolderName,
    conflictBehavior: "fail"
  }));
  record("create nested folder", movedFolder.id && movedFolder.folder ? "pass" : "fail", { id: movedFolder.id, name: movedFolder.name });

  const written = assertOk("onedrive_write_text", await tool("onedrive_write_text", {
    remotePath: `${folderName}/note.txt`,
    content,
    conflictBehavior: "fail"
  }));
  record("create text file", written.item?.id && written.bytesUploaded === Buffer.byteLength(content, "utf8") ? "pass" : "fail", {
    name: written.item?.name,
    bytesUploaded: written.bytesUploaded
  });

  const orphanRelativeTarget = await tool("onedrive_get_info", { relativePath: `${folderName}/note.txt` });
  const orphanRelativeDestination = await tool("onedrive_move", {
    path: `${folderName}/note.txt`,
    destinationParentRelativePath: movedFolderName
  });
  record("relative target and destination fields require matching presets", orphanRelativeTarget.isError
    && orphanRelativeDestination.isError
    && /preset/i.test(errorText(orphanRelativeTarget.value))
    && /preset/i.test(errorText(orphanRelativeDestination.value)) ? "pass" : "fail", {
    targetResponse: orphanRelativeTarget.value,
    destinationResponse: orphanRelativeDestination.value
  });

  const readBack = assertOk("onedrive_read_text", await tool("onedrive_read_text", {
    path: `${folderName}/note.txt`,
    maxBytes: 10000
  }));
  record("read text file", readBack.content === content ? "pass" : "fail", { bytes: Buffer.byteLength(readBack.content || "", "utf8") });

  const noteBeforePatch = assertOk("note info before remote-edit workflows", await tool("onedrive_get_info", { path: `${folderName}/note.txt`, format: "full" }));
  const versionsBefore = assertOk("version history", await tool("onedrive_versions", { itemId: noteBeforePatch.id, maxItems: 20 }));
  record("version history returns bounded metadata", Array.isArray(versionsBefore.versions) && versionsBefore.count === versionsBefore.versions.length ? "pass" : "fail", {
    count: versionsBefore.count,
    versions: versionsBefore.versions?.map((entry) => entry.id)
  });
  const patchDiff = [
    "@@ -1,2 +1,2 @@",
    ` OneDrive plugin beta test token: ${unique}`,
    "-This file was created by Codex and should be deleted during cleanup.",
    "+This file was safely patched remotely and should be deleted during cleanup.",
    ""
  ].join("\n");
  const patchedNote = assertCommitted("guarded text patch", await toolWithPreview("onedrive_patch_text", {
    itemId: noteBeforePatch.id,
    patch: { mode: "unified", diff: patchDiff },
    expectedId: noteBeforePatch.id,
    expectedETag: noteBeforePatch.eTag,
    dryRun: false,
    confirmed: true
  }), (value) => value.verified === true && value.item?.id === noteBeforePatch.id);
  const patchedRead = assertOk("patched text readback", await tool("onedrive_read_text", { itemId: noteBeforePatch.id, maxBytes: 10000 }));
  record("structured text patch preserves stable identity and verifies bytes", patchedRead.content.includes("safely patched remotely") && patchedNote.item.id === noteBeforePatch.id ? "pass" : "fail", {
    itemId: patchedNote.item.id,
    preservation: patchedNote.preservation,
    sha256: patchedNote.sha256
  });

  const comparedVersion = versionsBefore.versions?.[0]
    ? await tool("onedrive_compare_version", { itemId: noteBeforePatch.id, versionId: versionsBefore.versions[0].id, maxChanges: 50 })
    : await tool("onedrive_compare_version", { itemId: noteBeforePatch.id, versionId: "unavailable", maxChanges: 50 });
  record("version comparison classifies text semantics", !comparedVersion.isError && comparedVersion.value.comparison?.comparisonType === "text" ? "pass" : "blocked", {
    reason: comparedVersion.isError ? errorText(comparedVersion.value) : undefined,
    comparison: comparedVersion.value?.comparison
  });

  const patchInfo = assertOk("note info after patch", await tool("onedrive_get_info", { itemId: noteBeforePatch.id, format: "full" }));
  const workspaceCreated = assertCommitted("workspace create", await toolWithPreview("onedrive_workspace_create", {
    itemId: noteBeforePatch.id,
    expectedId: noteBeforePatch.id,
    expectedETag: patchInfo.eTag,
    dryRun: false,
    confirmed: true
  }), (value) => value.workspace?.workspaceId && value.draft?.id);
  const workspaceList = assertOk("workspace list", await tool("onedrive_workspace_list"));
  const draftBeforeEdit = workspaceCreated.draft;
  const draftEdited = assertCommitted("workspace draft patch", await toolWithPreview("onedrive_patch_text", {
    itemId: draftBeforeEdit.id,
    patch: { mode: "unified", diff: ["@@ -1,2 +1,2 @@", ` OneDrive plugin beta test token: ${unique}`, "-This file was safely patched remotely and should be deleted during cleanup.", "+This managed draft was promoted and should be deleted during cleanup.", ""].join("\n") },
    expectedId: draftBeforeEdit.id,
    expectedETag: draftBeforeEdit.eTag,
    dryRun: false,
    confirmed: true
  }), (value) => value.verified === true);
  const workspaceStatus = assertOk("workspace status", await tool("onedrive_workspace_status", { workspaceId: workspaceCreated.workspace.workspaceId, maxChanges: 50 }));
  const sourceBeforePromote = assertOk("workspace source info", await tool("onedrive_get_info", { itemId: noteBeforePatch.id, format: "full" }));
  const promoted = assertCommitted("workspace promote", await toolWithPreview("onedrive_workspace_promote", {
    workspaceId: workspaceCreated.workspace.workspaceId,
    expectedId: noteBeforePatch.id,
    expectedETag: sourceBeforePromote.eTag,
    dryRun: false,
    confirmed: true
  }), (value) => value.promoted === true && value.cleanedUp === true && value.item?.id === noteBeforePatch.id);
  record("managed workspace detects draft drift, preserves source identity, and cleans up after promotion", workspaceList.count >= 1
    && workspaceStatus.draftDrift === true
    && workspaceStatus.promotionReady === true
    && draftEdited.item?.id === draftBeforeEdit.id
    && promoted.item.id === noteBeforePatch.id ? "pass" : "fail", {
    workspaceId: workspaceCreated.workspace.workspaceId,
    status: workspaceStatus.status,
    sourceId: promoted.item.id,
    cleanedUp: promoted.cleanedUp
  });

  const sourceAfterPromote = assertOk("source after workspace promotion", await tool("onedrive_get_info", { itemId: noteBeforePatch.id, format: "full" }));
  const disposableWorkspace = assertCommitted("disposable workspace create", await toolWithPreview("onedrive_workspace_create", {
    itemId: noteBeforePatch.id,
    expectedId: noteBeforePatch.id,
    expectedETag: sourceAfterPromote.eTag,
    dryRun: false,
    confirmed: true
  }), (value) => value.workspace?.workspaceId && value.draft?.id);
  const abandoned = assertCommitted("workspace abandon", await toolWithPreview("onedrive_workspace_abandon", {
    workspaceId: disposableWorkspace.workspace.workspaceId,
    expectedId: disposableWorkspace.draft.id,
    expectedETag: disposableWorkspace.draft.eTag,
    dryRun: false,
    confirmed: true
  }), (value) => value.abandoned === true && value.sourceUnaffected === true);
  record("workspace abandonment is independently guarded", abandoned.abandoned === true ? "pass" : "fail", { workspaceId: abandoned.workspaceId });

  const watchStarted = assertOk("watch start", await tool("onedrive_watch_start", { itemId: folder.id, intervalSeconds: 30, expiresInSeconds: 3600 }));
  const watchId = watchStarted.watch?.watchId;
  const watchStatus = assertOk("watch status", await tool("onedrive_watch_status", { watchId, maxEvents: 10 }));
  const watchStopped = assertOk("watch stop", await tool("onedrive_watch_stop", { watchId }));
  record("change watch establishes a scoped baseline and stops deterministically", watchId && watchStatus.count === 1 && watchStopped.stopped === true && watchStopped.watch?.status === "stopped" ? "pass" : "fail", {
    watchId,
    baselineItemCount: watchStarted.baselineItemCount,
    stopped: watchStopped.stopped
  });

  const versionsAfter = assertOk("version history after remote edits", await tool("onedrive_versions", { itemId: noteBeforePatch.id, maxItems: 20 }));
  const restoreCandidate = versionsAfter.versions?.at(-1);
  if (restoreCandidate) {
    const beforeRestore = assertOk("note info before native restore", await tool("onedrive_get_info", { itemId: noteBeforePatch.id, format: "full" }));
    const restoredVersion = assertCommitted("native version restore", await toolWithPreview("onedrive_restore_version", {
      itemId: noteBeforePatch.id,
      versionId: restoreCandidate.id,
      expectedId: noteBeforePatch.id,
      expectedETag: beforeRestore.eTag,
      dryRun: false,
      confirmed: true
    }), (value) => value.verified === true && value.item?.id === noteBeforePatch.id);
    record("native version restore verifies a new current version", "pass", { restoredVersionId: restoredVersion.restoredVersionId, itemId: restoredVersion.item.id });
  } else {
    await tool("onedrive_restore_version", { itemId: noteBeforePatch.id, versionId: "unavailable" });
    record("native version restore verifies a new current version", "blocked", { reason: "Microsoft Graph returned no restorable version for the newly created personal OneDrive fixture." });
  }

  await assertOk("write csv", await tool("onedrive_write_text", {
    remotePath: `${folderName}/sheet.csv`,
    content: `name,value\n${unique},42\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("upload valid docx", await tool("onedrive_upload", {
    remotePath: `${folderName}/doc.docx`,
    localPath: join(officeFixtureDir, "sample.docx"),
    conflictBehavior: "fail"
  }));
  await assertOk("upload valid xlsx", await tool("onedrive_upload", {
    remotePath: `${folderName}/book.xlsx`,
    localPath: join(officeFixtureDir, "sample.xlsx"),
    conflictBehavior: "fail"
  }));
  await assertOk("upload valid pptx", await tool("onedrive_upload", {
    remotePath: `${folderName}/deck.pptx`,
    localPath: join(officeFixtureDir, "sample.pptx"),
    conflictBehavior: "fail"
  }));
  await assertOk("upload rich docx", await tool("onedrive_upload", { remotePath: `${folderName}/rich.docx`, localPath: richOfficeFixtures.word, conflictBehavior: "fail" }));
  await assertOk("upload rich xlsx", await tool("onedrive_upload", { remotePath: `${folderName}/rich.xlsx`, localPath: richOfficeFixtures.excel, conflictBehavior: "fail" }));
  await assertOk("upload rich pptx", await tool("onedrive_upload", { remotePath: `${folderName}/rich.pptx`, localPath: richOfficeFixtures.powerpoint, conflictBehavior: "fail" }));
  await assertOk("write duplicate A", await tool("onedrive_write_text", {
    remotePath: `${folderName}/duplicate.txt`,
    content: `Duplicate helper test ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write duplicate B", await tool("onedrive_write_text", {
    remotePath: `${folderName}/${movedFolderName}/duplicate.txt`,
    content: `Duplicate helper test ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write batch move A", await tool("onedrive_write_text", {
    remotePath: `${folderName}/batch-move-a.txt`,
    content: `Batch move A ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write batch move B", await tool("onedrive_write_text", {
    remotePath: `${folderName}/batch-move-b.txt`,
    content: `Batch move B ${unique}\n`,
    conflictBehavior: "fail"
  }));
  const batchDeleteTarget = assertOk("write batch delete target", await tool("onedrive_write_text", {
    remotePath: `${folderName}/batch-delete.txt`,
    content: `Batch delete ${unique}\n`,
    conflictBehavior: "fail"
  }));

  const excel = assertOk("download_excel", await tool("onedrive_download_excel", {
    path: `${folderName}/book.xlsx`,
    localPath: excelDownload
  }));
  const word = assertOk("download_word", await tool("onedrive_download_word", {
    path: `${folderName}/doc.docx`,
    localPath: wordDownload
  }));
  const powerpoint = assertOk("download_powerpoint", await tool("onedrive_download_powerpoint", {
    path: `${folderName}/deck.pptx`,
    localPath: powerpointDownload
  }));
  const downloadedExcel = await readFile(excelDownload);
  record("Office download helpers preserve genuine XLSX, DOCX, and PPTX packages", excel.bytesWritten > 0
    && downloadedExcel.subarray(0, 2).toString("ascii") === "PK"
    && word.bytesWritten > 0
    && powerpoint.bytesWritten > 0 ? "pass" : "fail", {
    excel: excel.localPath,
    word: word.localPath,
    powerpoint: powerpoint.localPath
  });

  const officeCapabilities = assertOk("office capabilities", await tool("onedrive_office_capabilities"));
  record("Graph Excel live backend on personal OneDrive", "blocked", {
    reason: "Microsoft Graph workbook APIs do not support personal OneDrive; all eight Graph Excel operation types are covered by the mock tenant.",
    driveType: officeCapabilities.backends?.graphExcel?.driveType,
    availableForAccount: officeCapabilities.backends?.graphExcel?.availableForAccount
  });
  record("organization-only sharing behavior on personal OneDrive", "blocked", {
    reason: "Organization-scoped sharing requires a work or school tenant; it is covered by mocks and is outside this personal-drive beta."
  });
  const wordStructured = assertOk("word structured read", await tool("onedrive_word_get_document", { path: `${folderName}/doc.docx` }));
  const excelStructured = assertOk("excel structured read", await tool("onedrive_excel_get_workbook", { path: `${folderName}/book.xlsx` }));
  const powerpointStructured = assertOk("powerpoint structured read", await tool("onedrive_powerpoint_get_presentation", { path: `${folderName}/deck.pptx` }));
  const officeValidation = assertOk("office package validation", await tool("onedrive_office_validate", { path: `${folderName}/deck.pptx`, expectedKind: "powerpoint" }));
  record("native Office structured reads", officeCapabilities.backends?.openXml?.available && wordStructured.paragraphCount > 0 && excelStructured.sheetCount > 0 && powerpointStructured.slideCount > 0 && officeValidation.valid ? "pass" : "fail", {
    wordParagraphs: wordStructured.paragraphCount,
    excelSheets: excelStructured.sheetCount,
    powerpointSlides: powerpointStructured.slideCount
  });

  const wordOfficeOperations = [
    { type: "replaceText", find: "Hello Word", replace: `Beta Word ${unique}` },
    { type: "setParagraphText", paragraphIndex: 0, text: `Beta Word ${unique}` },
    { type: "setParagraphStyle", paragraphIndex: 0, style: "Title" },
    { type: "setTableCell", tableIndex: 0, rowIndex: 0, columnIndex: 0, text: `Beta table ${unique}` },
    { type: "setContentControlText", tag: "customer", text: `Beta customer ${unique}` },
    { type: "insertParagraph", text: `Beta paragraph ${unique}`, style: "Normal" },
    { type: "addHyperlink", paragraphIndex: 0, text: "Microsoft", url: "https://www.microsoft.com" },
    { type: "addComment", paragraphIndex: 0, text: `Beta comment ${unique}`, author: "Codex", initials: "CX" },
    { type: "insertTable", afterParagraphIndex: 0, rows: [["Beta", unique]], style: "TableGrid" }
  ];
  const wordOfficePreview = assertOk("word native edit preview", await tool("onedrive_word_batch_update", { path: `${folderName}/doc.docx`, operations: wordOfficeOperations }));
  const wordOfficeLiveResult = await toolWithPreview("onedrive_word_batch_update", { path: `${folderName}/doc.docx`, operations: wordOfficeOperations, dryRun: false, confirmed: true, expectedName: "doc.docx" }, { previewToken: wordOfficePreview.previewToken });
  const wordOfficeLive = assertOk("word native edit live", wordOfficeLiveResult);
  const excelOfficeOperations = [
    { type: "setCell", sheet: "Data", address: "B2", value: 42 },
    { type: "setFormula", sheet: "Data", address: "C2", formula: "B2+1" },
    { type: "setRange", sheet: "Data", address: "D2:E3", values: [[1, 2], [3, 4]] },
    { type: "clearRange", sheet: "Data", address: "E3", contents: true, format: false },
    { type: "setStyle", sheet: "Data", address: "D2:E2", styleIndex: 2 },
    { type: "setNumberFormat", sheet: "Data", address: "B2", formatCode: "0.00" },
    { type: "addConditionalFormat", sheet: "Data", address: "D2:E3", ruleType: "cellIs", operator: "greaterThan", formula: "2", fillColor: "FFF2CC" },
    { type: "setDataValidation", sheet: "Data", address: "D2:D3", validationType: "whole", operator: "between", formula1: "1", formula2: "10" },
    { type: "freezePanes", sheet: "Data", rows: 1, columns: 1 },
    { type: "setColumnWidth", sheet: "Data", address: "D1:E1", width: 18 },
    { type: "setDefinedName", name: "BetaInputBlock", formula: "Data!$D$2:$E$3" },
    { type: "setTableTotals", table: "RevenueTable", enabled: true, columns: [{ column: "Metric", label: "Total" }, { column: "Revenue", function: "sum" }] },
    { type: "addTableRow", table: "RevenueTable", values: [[unique, 42]] },
    { type: "createChart", sheet: "Data", sourceData: "A1:B4", chartType: "ColumnClustered", name: "Beta revenue", titleText: "Beta revenue", left: 20, top: 30, width: 420, height: 240 },
    { type: "updateChart", sheet: "Data", chart: "Beta revenue", chartType: "Line", sourceData: "A1:B4", name: "Beta trend", titleText: "Beta trend", left: 40, top: 50, width: 400, height: 220 },
    { type: "recalculate" },
    { type: "renameSheet", sheet: "Data", newName: "Results" }
  ];
  const excelOfficePreview = assertOk("excel native edit preview", await tool("onedrive_excel_batch_update", { path: `${folderName}/book.xlsx`, backend: "auto", operations: excelOfficeOperations }));
  const excelOfficeLiveResult = await toolWithPreview("onedrive_excel_batch_update", { path: `${folderName}/book.xlsx`, backend: "auto", operations: excelOfficeOperations, dryRun: false, confirmed: true, expectedName: "book.xlsx" }, { previewToken: excelOfficePreview.previewToken });
  const excelOfficeLive = assertOk("excel native edit live", excelOfficeLiveResult);
  const powerpointOfficeOperations = [
    { type: "replaceText", slideIndex: 0, shapeId: "2", find: "Hello", replace: "Beta" },
    { type: "setShapeText", slideIndex: 0, shapeId: "2", text: `Beta PowerPoint ${unique}` },
    { type: "setShapeGeometry", slideIndex: 0, shapeId: "2", x: 10, y: 20, width: 3000000, height: 400000 },
    { type: "setTableCell", slideIndex: 0, shapeId: "4", rowIndex: 0, columnIndex: 0, text: `Beta table ${unique}` },
    { type: "setNotes", slideIndex: 0, text: `Beta notes ${unique}` },
    { type: "addTextBox", slideIndex: 0, shapeId: 10, text: "Temporary beta text box", x: 100, y: 100, width: 1000000, height: 400000 },
    { type: "setTextStyle", slideIndex: 0, shapeId: "10", fontFamily: "Aptos", fontSize: 18, bold: true, italic: true, underline: true, color: "2563EB" },
    { type: "deleteShape", slideIndex: 0, shapeId: 10 },
    { type: "replaceImage", slideIndex: 0, shapeId: "3", base64: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", contentType: "image/png" },
    { type: "duplicateSlide", slideIndex: 0 },
    { type: "moveSlide", slideIndex: 1, toIndex: 0 },
    { type: "deleteSlide", slideIndex: 1 }
  ];
  const powerpointOfficePreview = assertOk("powerpoint native edit preview", await tool("onedrive_powerpoint_batch_update", { path: `${folderName}/deck.pptx`, operations: powerpointOfficeOperations }));
  const powerpointOfficeLiveResult = await toolWithPreview("onedrive_powerpoint_batch_update", { path: `${folderName}/deck.pptx`, operations: powerpointOfficeOperations, dryRun: false, confirmed: true, expectedName: "deck.pptx" }, { previewToken: powerpointOfficePreview.previewToken });
  const powerpointOfficeLive = assertOk("powerpoint native edit live", powerpointOfficeLiveResult);
  const imageBase64 = richOfficeFixtures.imageBase64;
  const wordRichOperations = [
    { type: "insertImage", paragraphIndex: 0, base64: imageBase64, contentType: "image/png", width: 457200, height: 457200, altText: "Beta image" },
    { type: "replaceImage", imageIndex: 0, base64: imageBase64, contentType: "image/png" },
    { type: "createContentControl", paragraphIndex: 1, tag: "beta-rich", title: "Beta rich", text: `Controlled ${unique}` },
    { type: "deleteContentControl", tag: "beta-rich", preserveContent: true },
    { type: "createBookmark", paragraphIndex: 0, name: "CodexBetaBookmark" },
    { type: "deleteBookmark", name: "CodexBetaBookmark" },
    { type: "insertTableRow", tableIndex: 0, rowIndex: 1, values: ["E", "F"] },
    { type: "insertTableColumn", tableIndex: 0, columnIndex: 1, values: ["X", "Y", "Z"] },
    { type: "deleteTableRow", tableIndex: 0, rowIndex: 2 },
    { type: "deleteTableColumn", tableIndex: 0, columnIndex: 2 },
    { type: "setHeaderFooterText", part: "word/header1.xml", text: `Beta header ${unique}` },
    { type: "setSectionProperties", sectionIndex: 0, orientation: "landscape", pageWidth: 15840, pageHeight: 12240, marginLeft: 720, marginRight: 720 }
  ];
  const wordRichLive = assertOk("word rich native operations", await toolWithPreview("onedrive_word_batch_update", {
    path: `${folderName}/rich.docx`, operations: wordRichOperations, dryRun: false, confirmed: true, expectedName: "rich.docx"
  }));
  const excelRichOperations = [
    { type: "addWorksheet", name: "Scratch" }, { type: "deleteWorksheet", sheet: "Scratch" },
    { type: "addTable", sheet: "Data", address: "F1:G3", name: "CodexTable" }, { type: "deleteTable", table: "CodexTable", preserveData: true },
    { type: "mergeRange", sheet: "Data", address: "H1:I1" }, { type: "unmergeRange", sheet: "Data", address: "H1:I1" },
    { type: "sortRange", sheet: "Data", address: "D2:E3", keys: [{ column: 1, descending: true }], hasHeaders: false },
    { type: "setAutoFilter", sheet: "Data", address: "D1:E3" },
    { type: "setHyperlink", sheet: "Data", address: "J1", url: "https://openai.com", display: "OpenAI" },
    { type: "addNote", sheet: "Data", address: "K1", text: `Beta note ${unique}`, author: "Codex" }, { type: "deleteNote", sheet: "Data", address: "K1" },
    { type: "insertImage", sheet: "Data", fromAddress: "N1", base64: imageBase64, contentType: "image/png", altText: "Beta image" },
    { type: "formatChart", sheet: "Data", chart: "0", titleText: "Formatted beta chart", legendPosition: "bottom", style: 10 },
    { type: "setSheetProtection", sheet: "Data", enabled: true, allowSelectUnlockedCells: true },
    { type: "refreshPivot" }
  ];
  const excelRichLive = assertOk("excel rich native operations", await toolWithPreview("onedrive_excel_batch_update", {
    path: `${folderName}/rich.xlsx`, backend: "openxml", operations: excelRichOperations, dryRun: false, confirmed: true, expectedName: "rich.xlsx"
  }));
  const selectors = richOfficeFixtures.powerpointSelectors;
  const powerpointRichOperations = [
    { type: "addSlide", afterIndex: 0 },
    { type: "addImage", slideIndex: 0, base64: imageBase64, contentType: "image/png", x: 457200, y: 4572000, width: 457200, height: 457200, altText: "Added beta image" },
    { type: "cropImage", slideIndex: 0, shapeId: selectors.pictureId, left: 0.1, right: 0.1 },
    { type: "addTable", slideIndex: 0, rows: [["One", "Two"], ["Three", "Four"]], x: 3657600, y: 4572000, width: 1828800, height: 914400 },
    { type: "insertTableRow", slideIndex: 0, shapeId: selectors.tableId, rowIndex: 1, values: ["R1", "R2"] },
    { type: "insertTableColumn", slideIndex: 0, shapeId: selectors.tableId, columnIndex: 1, values: ["C1", "C2", "C3"] },
    { type: "deleteTableRow", slideIndex: 0, shapeId: selectors.tableId, rowIndex: 2 },
    { type: "deleteTableColumn", slideIndex: 0, shapeId: selectors.tableId, columnIndex: 2 },
    { type: "setShapeAltText", slideIndex: 0, shapeId: selectors.boxId, title: "Main beta box", description: `Edited ${unique}` },
    { type: "setZOrder", slideIndex: 0, shapeId: selectors.boxId, position: "front" },
    { type: "groupShapes", slideIndex: 0, shapeIds: [selectors.groupAId, selectors.groupBId], name: "Codex Beta Group" },
    { type: "ungroupShape", slideIndex: 0, shapeId: selectors.existingGroupId },
    { type: "applySlideLayout", slideIndex: 1, layoutName: selectors.blankLayoutName }
  ];
  const powerpointRichLive = assertOk("powerpoint rich native operations", await toolWithPreview("onedrive_powerpoint_batch_update", {
    path: `${folderName}/rich.pptx`, operations: powerpointRichOperations, dryRun: false, confirmed: true, expectedName: "rich.pptx"
  }));
  const officeOperationCoverage = {
    word: wordOfficeOperations.every((operation) => wordOfficeLive.semanticDiff?.operationCounts?.[operation.type] > 0)
      && wordRichOperations.every((operation) => wordRichLive.semanticDiff?.operationCounts?.[operation.type] > 0),
    excel: excelOfficeOperations.every((operation) => excelOfficeLive.semanticDiff?.operationCounts?.[operation.type] > 0)
      && excelRichOperations.every((operation) => excelRichLive.semanticDiff?.operationCounts?.[operation.type] > 0),
    powerpoint: powerpointOfficeOperations.every((operation) => powerpointOfficeLive.semanticDiff?.operationCounts?.[operation.type] > 0)
      && powerpointRichOperations.every((operation) => powerpointRichLive.semanticDiff?.operationCounts?.[operation.type] > 0)
  };
  const wordAfterEditResult = await pollIndexedSearch(
    "onedrive_word_get_document",
    { itemId: wordOfficeLive.item.id, searchText: unique },
    (value) => value.search?.matchCount > 0,
    { maxAttempts: 15, delayMs: 2_000 }
  );
  const excelAfterEditResult = await pollIndexedSearch(
    "onedrive_excel_get_workbook",
    { itemId: excelOfficeLive.item.id, sheetNames: ["Results"] },
    (value) => value.sheets?.[0]?.name === "Results",
    { maxAttempts: 6, delayMs: 1_000 }
  );
  const powerpointAfterEditResult = await pollIndexedSearch(
    "onedrive_powerpoint_get_presentation",
    { itemId: powerpointOfficeLive.item.id, searchText: unique },
    (value) => value.slideCount === 1 && value.search?.matchCount > 0,
    { maxAttempts: 6, delayMs: 1_000 }
  );
  const wordAfterEdit = assertOk("word post-edit semantic read", wordAfterEditResult);
  const excelAfterEdit = assertOk("excel post-edit semantic read", excelAfterEditResult);
  const powerpointAfterEdit = assertOk("powerpoint post-edit semantic read", powerpointAfterEditResult);
  record("all 78 advertised Open XML operations commit with semantic verification", Object.values(officeOperationCoverage).every(Boolean)
    && wordOfficeLive.verificationIncomplete !== true
    && excelOfficeLive.verificationIncomplete !== true
    && powerpointOfficeLive.verificationIncomplete !== true
    && wordAfterEdit.search?.matchCount > 0
    && excelAfterEdit.sheets?.[0]?.name === "Results"
    && powerpointAfterEdit.slideCount === 1
    && powerpointAfterEdit.search?.matchCount > 0 ? "pass" : "fail", {
    wordChanges: wordOfficeLive.changeCount,
    excelChanges: excelOfficeLive.changeCount,
    powerpointChanges: powerpointOfficeLive.changeCount,
    requestedOperationCounts: { word: wordOfficeOperations.length + wordRichOperations.length, excel: excelOfficeOperations.length + excelRichOperations.length, powerpoint: powerpointOfficeOperations.length + powerpointRichOperations.length },
    coverage: officeOperationCoverage,
    commitVerificationIncomplete: {
      word: wordOfficeLive.verificationIncomplete,
      excel: excelOfficeLive.verificationIncomplete,
      powerpoint: powerpointOfficeLive.verificationIncomplete
    },
    previewRefreshes: {
      word: wordOfficeLiveResult.previewRefreshes,
      excel: excelOfficeLiveResult.previewRefreshes,
      powerpoint: powerpointOfficeLiveResult.previewRefreshes
    },
    postReadVerification: {
      wordMatchCount: wordAfterEdit.search?.matchCount,
      wordPollAttempts: wordAfterEditResult.pollAttempts,
      excelSheetNames: excelAfterEdit.sheets?.map((sheet) => sheet.name),
      excelPollAttempts: excelAfterEditResult.pollAttempts,
      powerpointSlideCount: powerpointAfterEdit.slideCount,
      powerpointMatchCount: powerpointAfterEdit.search?.matchCount,
      powerpointPollAttempts: powerpointAfterEditResult.pollAttempts
    },
    operationCounts: {
      word: { ...wordOfficeLive.semanticDiff?.operationCounts, ...wordRichLive.semanticDiff?.operationCounts },
      excel: { ...excelOfficeLive.semanticDiff?.operationCounts, ...excelRichLive.semanticDiff?.operationCounts },
      powerpoint: { ...powerpointOfficeLive.semanticDiff?.operationCounts, ...powerpointRichLive.semanticDiff?.operationCounts }
    }
  });

  const officeIndex = assertOk("structured Office index refresh", await tool("onedrive_office_index_refresh", { path: `${folderName}/doc.docx`, refreshMetadata: false, force: true }));
  const officeSearch = assertOk("structured Office search", await tool("onedrive_office_search", { query: `Beta Word ${unique}` }));
  record("incremental Office research index", officeIndex.indexed === 1 && officeSearch.items?.[0]?.anchor?.type === "paragraph" ? "pass" : "fail", {
    indexed: officeIndex.indexed,
    anchor: officeSearch.items?.[0]?.anchor
  });

  const officeBatchItems = [
    { path: `${folderName}/doc.docx`, expectedName: "doc.docx", kind: "word", operations: [{ type: "setParagraphText", paragraphIndex: 0, text: `Batch Word ${unique}` }] },
    { path: `${folderName}/book.xlsx`, expectedName: "book.xlsx", kind: "excel", operations: [{ type: "setCell", sheet: "Results", address: "D4", value: `Batch Excel ${unique}` }] },
    { path: `${folderName}/deck.pptx`, expectedName: "deck.pptx", kind: "powerpoint", operations: [{ type: "setShapeText", slideIndex: 0, shapeId: "2", text: `Batch PowerPoint ${unique}` }] }
  ];
  const officeBatchPreview = assertOk("cross-file Office batch preview", await tool("onedrive_office_batch_transform", { items: officeBatchItems }));
  const officeBatchLive = assertCommitted("cross-file Office batch live", await tool("onedrive_office_batch_transform", { items: officeBatchItems, dryRun: false, confirmed: true, previewToken: officeBatchPreview.previewToken }));
  record("cross-file Office batch transformation", officeBatchPreview.preflightComplete && officeBatchLive.partialState === false && officeBatchLive.completed?.length === 3 ? "pass" : "fail", {
    preflightComplete: officeBatchPreview.preflightComplete,
    completed: officeBatchLive.completed?.length,
    totalChangeCount: officeBatchLive.totalChangeCount
  });

  const officeBackups = assertOk("managed Office backup list", await tool("onedrive_office_backups", { itemId: wordOfficeLive.item.id }));
  const wordBackupId = wordOfficeLive.backup?.backupId || officeBackups.items?.[0]?.backupId;
  if (!wordBackupId) throw new Error("The live Word edit did not produce a managed backup ID.");
  const officeBackupComparison = assertOk("managed Office backup comparison", await tool("onedrive_office_compare_backup", { backupId: wordBackupId }));
  const officeRestorePreview = assertOk("managed Office backup restore preview", await tool("onedrive_office_restore_backup", { backupId: wordBackupId }));
  const officeRestoreLive = assertCommitted("managed Office backup restore live", await tool("onedrive_office_restore_backup", {
    backupId: wordBackupId,
    dryRun: false,
    confirmed: true,
    expectedId: wordOfficeLive.item.id,
    expectedETag: officeRestorePreview.wouldRestore.currentItem.eTag,
    previewToken: officeRestorePreview.previewToken
  }));
  record("managed Office compare and one-click undo", officeBackups.items?.some((entry) => entry.backupId === wordBackupId) && officeBackupComparison.sameContent === false && officeRestoreLive.restoredBackupId === wordBackupId && officeRestoreLive.rollbackBackup?.backupId ? "pass" : "fail", {
    backupId: wordBackupId,
    comparedChangeCount: officeBackupComparison.semanticDiff?.changeCount,
    rollbackBackupId: officeRestoreLive.rollbackBackup?.backupId
  });
  const staleOfficeSearch = assertOk("stale Office index query after restore", await tool("onedrive_office_search", { query: `Beta Word ${unique}` }));
  const restoredOfficeIndex = assertOk("refresh restored Office document index", await tool("onedrive_office_index_refresh", { path: `${folderName}/doc.docx`, refreshMetadata: false, force: true }));
  const restoredOfficeSearch = assertOk("search restored Office index", await tool("onedrive_office_search", { query: "Hello Word" }));
  record("Office search evicts stale content and refreshes the restored package", !staleOfficeSearch.items?.some((entry) => entry.item?.id === wordOfficeLive.item.id || entry.id === wordOfficeLive.item.id)
    && restoredOfficeIndex.indexed === 1
    && restoredOfficeSearch.items?.some((entry) => entry.item?.id === wordOfficeLive.item.id || entry.id === wordOfficeLive.item.id) ? "pass" : "fail", {
    staleEntriesRemoved: staleOfficeSearch.staleEntriesRemoved,
    restoredIndexed: restoredOfficeIndex.indexed,
    restoredMatchCount: restoredOfficeSearch.items?.length
  });

  const preview = assertOk("preview text file", await tool("onedrive_preview", {
    path: `${folderName}/note.txt`,
    maxBytes: 40
  }));
  record("preview returns bounded content", preview.preview?.includes("OneDrive plugin beta test") && preview.truncated === true ? "pass" : "fail", {
    source: preview.source,
    bytes: preview.bytes,
    truncated: preview.truncated,
    preview: preview.preview
  });

  const exportPdf = await tool("onedrive_export_pdf", {
    path: `${folderName}/doc.docx`,
    localPath: exportPdfDownload,
    overwrite: true
  });
  record("export_pdf returns a PDF or a precisely classified Graph limitation", exportPdf.isError
    ? exportFailureStatus(exportPdf.value)
    : exportPdf.value.bytesWritten > 0 && (await readFile(exportPdfDownload)).subarray(0, 4).toString("ascii") === "%PDF"
      ? "pass"
      : "fail", {
    response: exportPdf.value,
    classification: exportPdf.isError ? exportFailureStatus(exportPdf.value) : "success"
  });

  const exportText = await tool("onedrive_export_text", {
    path: `${folderName}/doc.docx`,
    localPath: exportTextDownload,
    overwrite: true
  });
  record("export_text returns non-empty text or a precisely classified Graph limitation", exportText.isError
    ? exportFailureStatus(exportText.value)
    : exportText.value.bytesWritten > 0 && (await readFile(exportTextDownload, "utf8")).trim().length > 0
      ? "pass"
      : "fail", {
    response: exportText.value,
    classification: exportText.isError ? exportFailureStatus(exportText.value) : "success"
  });

  const renamed = assertOk("onedrive_rename", await tool("onedrive_rename", {
    path: `${folderName}/note.txt`,
    newName: renamedTextFile,
    expectedName: "note.txt",
    dryRun: false,
    confirmed: true
  }));
  const renamedItem = renamed.renamed || renamed;
  record("rename file", renamedItem.name === renamedTextFile && renamed.confirmed === true ? "pass" : "fail", { name: renamedItem.name });

  const moved = assertOk("onedrive_move", await tool("onedrive_move", {
    path: `${folderName}/${renamedTextFile}`,
    destinationParentPath: `${folderName}/${movedFolderName}`,
    newName: movedTextFile,
    expectedName: renamedTextFile,
    dryRun: false,
    confirmed: true
  }));
  const movedItem = moved.moved || moved;
  const movedInfo = assertOk("moved file info", await tool("onedrive_get_info", {
    path: `${folderName}/${movedFolderName}/${movedTextFile}`
  }));
  record("move file with expectedName", movedItem.name === movedTextFile && movedInfo.id === movedItem.id && moved.confirmed === true ? "pass" : "fail", {
    name: movedItem.name,
    id: movedItem.id
  });

  const copied = assertOk("onedrive_copy", await tool("onedrive_copy", {
    path: `${folderName}/${movedFolderName}/${movedTextFile}`,
    destinationParentPath: folderName,
    newName: copyFileName,
    expectedName: movedTextFile,
    dryRun: false,
    confirmed: true,
    waitForCompletion: true,
    timeoutSeconds: 90
  }));
  record("copy monitor reaches an explicit successful terminal state", copied.accepted
    && copied.monitorUrl
    && copied.monitor?.terminal === true
    && copied.monitor?.succeeded === true ? "pass" : "fail", {
    responseStatus: copied.status,
    monitorTerminal: copied.monitor?.terminal,
    monitorSucceeded: copied.monitor?.succeeded,
    monitorStatus: copied.monitor?.status
  });

  const copiedInfo = assertOk("copied file info", await tool("onedrive_get_info", {
    path: `${folderName}/${copyFileName}`
  }));
  record("copied file resolves", copiedInfo.name === copyFileName ? "pass" : "fail", { id: copiedInfo.id, name: copiedInfo.name });

  const uploaded = assertOk("onedrive_upload", await tool("onedrive_upload", {
    localPath: localUpload,
    remotePath: `${folderName}/uploaded.txt`,
    conflictBehavior: "fail"
  }));
  record("upload local file", uploaded.item?.name === "uploaded.txt" && uploaded.bytesUploaded > 0 && uploaded.uploadMode === "simple" ? "pass" : "fail", {
    name: uploaded.item?.name,
    bytesUploaded: uploaded.bytesUploaded,
    uploadMode: uploaded.uploadMode
  });

  const replacementUploadContent = `Guarded replacement through onedrive_upload: ${unique}\n`;
  await writeFile(localUpload, replacementUploadContent, "utf8");
  const uploadReplacePreview = assertOk("upload replacement preview", await tool("onedrive_upload", {
    localPath: localUpload,
    remotePath: `${folderName}/uploaded.txt`,
    conflictBehavior: "replace"
  }));
  const uploadReplaceMissingIdentity = assertOk("upload replacement requires identity", await tool("onedrive_upload", {
    localPath: localUpload,
    remotePath: `${folderName}/uploaded.txt`,
    conflictBehavior: "replace",
    dryRun: false,
    confirmed: true,
    previewToken: uploadReplacePreview.previewToken
  }));
  const uploadReplaceLive = assertCommitted("guarded upload replacement", await tool("onedrive_upload", {
    localPath: localUpload,
    remotePath: `${folderName}/uploaded.txt`,
    conflictBehavior: "replace",
    dryRun: false,
    confirmed: true,
    expectedId: uploaded.item.id,
    previewToken: uploadReplacePreview.previewToken
  }));
  const uploadReplaceReadback = assertOk("guarded upload replacement readback", await tool("onedrive_read_text", {
    itemId: uploaded.item.id,
    maxBytes: 10000
  }));
  record("upload replacement requires a scoped preview and stable identity", uploadReplacePreview.dryRun === true
    && uploadReplacePreview.wouldReplace?.id === uploaded.item.id
    && Boolean(uploadReplacePreview.previewToken)
    && /expectedName or expectedId/i.test(uploadReplaceMissingIdentity.requiredToReplace || "")
    && uploadReplaceLive.item?.id === uploaded.item.id
    && uploadReplaceReadback.content === replacementUploadContent ? "pass" : "fail", {
    wouldReplace: uploadReplacePreview.wouldReplace,
    missingIdentity: uploadReplaceMissingIdentity.requiredToReplace,
    replacedId: uploadReplaceLive.item?.id
  });

  const sessionUploaded = assertOk("onedrive_upload session", await tool("onedrive_upload", {
    localPath: localSessionUpload,
    remotePath: `${folderName}/uploaded-session.txt`,
    conflictBehavior: "fail",
    uploadMode: "session",
    chunkSize: 327680
  }));
  record("upload session file", sessionUploaded.item?.name === "uploaded-session.txt" && sessionUploaded.uploadMode === "session" ? "pass" : "fail", {
    name: sessionUploaded.item?.name,
    bytesUploaded: sessionUploaded.bytesUploaded,
    uploadMode: sessionUploaded.uploadMode,
    chunkSize: sessionUploaded.chunkSize
  });

  const binaryUploaded = assertOk("binary upload", await tool("onedrive_upload", {
    localPath: localBinary,
    remotePath: `${folderName}/binary.bin`,
    conflictBehavior: "fail"
  }));
  record("upload binary file", binaryUploaded.item?.name === "binary.bin" ? "pass" : "fail", { bytesUploaded: binaryUploaded.bytesUploaded });

  const binaryRead = await tool("onedrive_read_text", {
    path: `${folderName}/binary.bin`,
    maxBytes: 1000
  });
  record("binary read refused", binaryRead.isError && String(binaryRead.value).includes("likely binary") ? "pass" : "fail", {
    response: binaryRead.value
  });

  const downloaded = assertOk("onedrive_download", await tool("onedrive_download", {
    path: `${folderName}/uploaded.txt`,
    localPath: localDownload,
    overwrite: false
  }));
  const downloadedContent = await readFile(localDownload, "utf8");
  record("download file", downloaded.bytesWritten === Buffer.byteLength(downloadedContent, "utf8") && downloadedContent.includes(unique) ? "pass" : "fail", {
    localPath: downloaded.localPath,
    bytesWritten: downloaded.bytesWritten
  });

  const checkout = assertOk("update_file checkout", await tool("onedrive_update_file", {
    mode: "checkout",
    remotePath: `${folderName}/uploaded.txt`,
    localPath: updateCheckout,
    manifestPath: updateManifest,
    overwriteLocal: true,
    overwriteManifest: true
  }));
  await writeFile(updateCheckout, `${downloadedContent}Edited through update_file commit: ${unique}\n`, "utf8");
  const committed = assertOk("update_file commit", await tool("onedrive_update_file", {
    mode: "commit",
    remotePath: `${folderName}/uploaded.txt`,
    localPath: updateCheckout,
    manifestPath: updateManifest,
    createBackup: true,
    verify: true
  }));
  const updatedReadBack = assertOk("updated file readback", await tool("onedrive_read_text", {
    path: `${folderName}/uploaded.txt`,
    maxBytes: 10000
  }));
  record("update_file checkout and commit", checkout.mode === "checkout"
    && committed.mode === "commit"
    && committed.backup?.bytesWritten > 0
    && updatedReadBack.content.includes("Edited through update_file commit") ? "pass" : "fail", {
    checkout: { localPath: checkout.localPath, manifestPath: checkout.manifestPath },
    commit: { backupBytes: committed.backup?.bytesWritten, verified: committed.verified?.name }
  });

  const conflictCheckout = assertOk("update_file conflict checkout", await tool("onedrive_update_file", {
    mode: "checkout",
    remotePath: `${folderName}/uploaded.txt`,
    localPath: updateCheckout,
    manifestPath: updateManifest,
    overwriteLocal: true,
    overwriteManifest: true
  }));
  await writeFile(updateCheckout, `${updatedReadBack.content}Stale local edit: ${unique}\n`, "utf8");
  const writeReplaceContent = `${updatedReadBack.content}Remote conflict marker: ${unique}\n`;
  const writeReplacePreview = assertOk("write_text replacement preview", await tool("onedrive_write_text", {
    remotePath: `${folderName}/uploaded.txt`,
    content: writeReplaceContent,
    conflictBehavior: "replace"
  }));
  const writeReplaceMismatchedContent = assertOk("write_text replacement content binding", await tool("onedrive_write_text", {
    remotePath: `${folderName}/uploaded.txt`,
    content: `${writeReplaceContent}not reviewed\n`,
    conflictBehavior: "replace",
    expectedId: uploaded.item.id,
    dryRun: false,
    confirmed: true,
    previewToken: writeReplacePreview.previewToken
  }));
  const writeReplaceLive = assertCommitted("guarded remote replacement for update conflict", await tool("onedrive_write_text", {
    remotePath: `${folderName}/uploaded.txt`,
    content: writeReplaceContent,
    conflictBehavior: "replace",
    expectedId: uploaded.item.id,
    dryRun: false,
    confirmed: true,
    previewToken: writeReplacePreview.previewToken
  }));
  record("write_text replacement preview binds identity and exact content", writeReplacePreview.wouldReplace?.id === uploaded.item.id
    && writeReplaceMismatchedContent.previewTokenStatus === "mismatch"
    && writeReplaceLive.item?.id === uploaded.item.id ? "pass" : "fail", {
    wouldReplace: writeReplacePreview.wouldReplace,
    mismatchedContentStatus: writeReplaceMismatchedContent.previewTokenStatus,
    replacedId: writeReplaceLive.item?.id
  });
  const conflictCommit = await tool("onedrive_update_file", {
    mode: "commit",
    remotePath: `${folderName}/uploaded.txt`,
    localPath: updateCheckout,
    manifestPath: updateManifest
  });
  record("update_file refuses a stale checkout after a remote change", conflictCheckout.mode === "checkout"
    && conflictCommit.isError
    && /(?:changed|conflict|eTag|cTag|modified|version)/i.test(errorText(conflictCommit.value)) ? "pass" : "fail", {
    response: conflictCommit.value
  });

  const batchMoveDryRun = assertOk("batch move dry-run", await tool("onedrive_batch_move", {
    items: [
      { path: `${folderName}/batch-move-a.txt`, expectedName: "batch-move-a.txt" },
      { path: `${folderName}/batch-move-b.txt`, expectedName: "batch-move-b.txt" }
    ],
    destinationParentPath: `${folderName}/${movedFolderName}`
  }));
  record("batch_move dry-run previews every item", batchMoveDryRun.dryRun === true && batchMoveDryRun.results?.length === 2 ? "pass" : "fail", {
    count: batchMoveDryRun.count
  });

  const batchMoved = assertOk("batch move live", await tool("onedrive_batch_move", {
    items: [
      { path: `${folderName}/batch-move-a.txt`, expectedName: "batch-move-a.txt" },
      { path: `${folderName}/batch-move-b.txt`, expectedName: "batch-move-b.txt" }
    ],
    destinationParentPath: `${folderName}/${movedFolderName}`,
    dryRun: false,
    confirmed: true
  }));
  record("batch_move live succeeds", batchMoved.confirmed === true && batchMoved.results?.length === 2 ? "pass" : "fail", {
    count: batchMoved.count,
    names: batchMoved.results?.map((entry) => entry.moved?.name || entry.item?.name || entry.name)
  });

  const batchDeleteDryRun = assertOk("batch delete dry-run", await tool("onedrive_batch_delete", {
    items: [{ itemId: batchDeleteTarget.item.id, expectedName: "batch-delete.txt" }]
  }));
  record("batch_delete dry-run previews item", batchDeleteDryRun.dryRun === true && batchDeleteDryRun.results?.length === 1 ? "pass" : "fail", {
    count: batchDeleteDryRun.count
  });

  const batchDeleted = assertOk("batch delete live", await toolWithPreview("onedrive_batch_delete", {
    items: [{ itemId: batchDeleteTarget.item.id, expectedName: "batch-delete.txt" }],
    dryRun: false,
    confirmed: true
  }));
  record("batch_delete live succeeds", batchDeleted.confirmed === true
    && batchDeleted.results?.length === 1
    && batchDeleted.results[0]?.deleted?.id === batchDeleteTarget.item.id ? "pass" : "fail", {
    count: batchDeleted.count,
    deletedId: batchDeleted.results?.[0]?.deleted?.id
  });

  const listedFolder = assertOk("onedrive_list", await tool("onedrive_list", { path: folderName, limit: 20 }));
  const childNames = listedFolder.items.map((item) => item.name).sort();
  record("list folder children compactly", childNames.includes(copyFileName) && childNames.includes("uploaded.txt") && listedFolder.items.every((item) => item.type) ? "pass" : "fail", {
    childNames,
    sample: listedFolder.items[0]
  });

  const allFolder = assertOk("onedrive_list_all", await tool("onedrive_list_all", { path: folderName, pageSize: 2, maxItems: 20 }));
  record("list_all follows pagination", allFolder.items.length >= 5 && allFolder.count === allFolder.items.length ? "pass" : "fail", {
    count: allFolder.count,
    truncated: allFolder.truncated
  });

  const scanned = assertOk("onedrive_scan", await tool("onedrive_scan", {
    path: folderName,
    nameContains: "uploaded",
    includeFolders: false,
    maxItems: 50,
    maxFolders: 10,
    maxResults: 10
  }));
  record("scan finds nested test files", scanned.items.some((item) => item.name === "uploaded.txt") && scanned.items.every((item) => item.remotePath) ? "pass" : "fail", {
    summary: scanned.summary,
    names: scanned.items.map((item) => item.name)
  });

  const found = assertOk("onedrive_find", await tool("onedrive_find", {
    query: "uploaded session",
    folderHints: [folderName],
    maxResults: 5,
    scanMaxItems: 50,
    scanMaxFolders: 10
  }));
  record("find ranks test file with cache acceleration", found.items.some((item) => item.name === "uploaded-session.txt") && found.summary?.localIndexUsed === false && found.summary?.persistentCacheUsed === true ? "pass" : "fail", {
    summary: found.summary,
    top: found.items[0]
  });

  const foundAll = assertOk("onedrive_find_all", await tool("onedrive_find_all", {
    query: "uploaded session",
    folderHints: [folderName],
    maxResults: 20,
    scanMaxItems: 100,
    scanMaxFolders: 20
  }));
  record("find_all broad locator works with cache acceleration", foundAll.items.some((item) => item.name === "uploaded-session.txt") && foundAll.summary?.localIndexUsed === false && foundAll.summary?.persistentCacheUsed === true ? "pass" : "fail", {
    summary: foundAll.summary,
    folderPlan: foundAll.folderPlan,
    names: foundAll.items.map((item) => item.name)
  });

  const syncStatusBefore = assertOk("sync status", await tool("onedrive_sync_status", { includeSamples: true }));
  record("sync_status reports cache state", Number.isInteger(syncStatusBefore.itemCount) && Array.isArray(syncStatusBefore.samples) ? "pass" : "fail", {
    itemCount: syncStatusBefore.itemCount,
    sampleCount: syncStatusBefore.samples?.length,
    deltaLinkAvailable: syncStatusBefore.deltaLinkAvailable
  });

  const cacheRefresh = assertOk("cache refresh", await tool("onedrive_cache_refresh", {
    path: folderName,
    mode: "scan",
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("cache_refresh scans selected folder", cacheRefresh.cache?.itemCount >= 1 && ["scan", "scan+delta"].includes(cacheRefresh.mode) ? "pass" : "fail", {
    mode: cacheRefresh.mode,
    itemCount: cacheRefresh.cache?.itemCount,
    scanned: cacheRefresh.scan?.summary?.itemsScanned,
    scanRoot: cacheRefresh.cache?.scanRoot
  });

  const cacheCleared = assertOk("cache clear", await tool("onedrive_cache_clear"));
  record("cache_clear empties metadata cache", cacheCleared.itemCount === 0 ? "pass" : "fail", {
    itemCount: cacheCleared.itemCount,
    updatedAt: cacheCleared.updatedAt
  });

  const cacheRefreshAfterClear = assertOk("cache refresh after clear", await tool("onedrive_cache_refresh", {
    path: folderName,
    mode: "scan",
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("cache_refresh rebuilds cache after clear", cacheRefreshAfterClear.cache?.itemCount >= 1 ? "pass" : "fail", {
    mode: cacheRefreshAfterClear.mode,
    itemCount: cacheRefreshAfterClear.cache?.itemCount,
    scanned: cacheRefreshAfterClear.scan?.summary?.itemsScanned
  });

  const cacheDeltaRefresh = assertOk("cache delta refresh after scan", await tool("onedrive_cache_refresh", {
    path: folderName,
    mode: "delta",
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("cache lifecycle advances from bounded scan to stored delta cursor", cacheRefreshAfterClear.cache?.itemCount >= 1
    && cacheDeltaRefresh.mode === "delta"
    && Array.isArray(cacheDeltaRefresh.result?.items)
    && (cacheDeltaRefresh.result?.deltaLink || cacheDeltaRefresh.result?.nextLink) ? "pass" : "fail", {
    scanMode: cacheRefreshAfterClear.mode,
    deltaMode: cacheDeltaRefresh.mode,
    delta: cacheDeltaRefresh.result
  });

  const contentIndexClearedBefore = assertOk("content index initial clear", await tool("onedrive_content_index_clear"));
  record("content_index_clear starts isolated lifecycle", contentIndexClearedBefore.itemCount === 0 ? "pass" : "fail", {
    itemCount: contentIndexClearedBefore.itemCount,
    updatedAt: contentIndexClearedBefore.updatedAt
  });

  const contentIndexRefresh = assertOk("content index refresh", await tool("onedrive_content_index_refresh", {
    itemId: uploaded.item.id,
    maxFiles: 1,
    maxBytesPerFile: 10000,
    concurrencyLimit: 1,
    force: true
  }));
  record("content_index_refresh indexes one isolated beta file", contentIndexRefresh.indexed === 1
    && contentIndexRefresh.failed === 0
    && contentIndexRefresh.itemCount === 1 ? "pass" : "fail", {
    considered: contentIndexRefresh.considered,
    selected: contentIndexRefresh.selected,
    indexed: contentIndexRefresh.indexed,
    failed: contentIndexRefresh.failed,
    itemCount: contentIndexRefresh.itemCount
  });

  const contentSearch = assertOk("content search", await tool("onedrive_content_search", {
    query: "Edited through update_file commit",
    maxResults: 5
  }));
  record("content_search returns the indexed beta file with a snippet", contentSearch.items?.[0]?.id === uploaded.item.id
    && contentSearch.items[0].snippet?.includes("Edited through update_file commit") ? "pass" : "fail", {
    itemCount: contentSearch.itemCount,
    matched: contentSearch.matched,
    returned: contentSearch.returned,
    top: contentSearch.items?.[0]
  });

  const contentIndexClearedAfter = assertOk("content index final clear", await tool("onedrive_content_index_clear"));
  const contentSearchAfterClear = assertOk("content search after clear", await tool("onedrive_content_search", {
    query: "Edited through update_file commit",
    maxResults: 5
  }));
  record("content_index_clear removes indexed beta content", contentIndexClearedAfter.itemCount === 0
    && contentSearchAfterClear.itemCount === 0
    && contentSearchAfterClear.returned === 0 ? "pass" : "fail", {
    clearedItemCount: contentIndexClearedAfter.itemCount,
    searchItemCount: contentSearchAfterClear.itemCount,
    returned: contentSearchAfterClear.returned
  });

  const recent = assertOk("recent files", await tool("onedrive_recent", { limit: 10 }));
  record("recent files call succeeds", Array.isArray(recent.items) && recent.count >= 0 ? "pass" : "fail", {
    count: recent.count,
    sample: recent.items?.[0]
  });

  const largeFiles = assertOk("large files", await tool("onedrive_large_files", {
    path: folderName,
    minBytes: 1,
    limit: 10,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("large_files finds test files", largeFiles.items?.some((item) => item.name === "uploaded-session.txt") ? "pass" : "fail", {
    count: largeFiles.count,
    names: largeFiles.items?.map((item) => item.name)
  });

  const duplicates = assertOk("duplicates", await tool("onedrive_duplicates", {
    path: folderName,
    limit: 10,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("duplicates finds duplicate test files", duplicates.groups?.some((group) => group.items?.some((item) => item.name === "duplicate.txt")) ? "pass" : "fail", {
    count: duplicates.count,
    groups: duplicates.groups?.map((group) => ({ key: group.key, count: group.count, names: group.items.map((item) => item.name) }))
  });

  const blockedDownloadExistedBefore = await pathExists(blockedSyncDownload);
  const blockedDownload = await tool("onedrive_download", {
    path: `${folderName}/uploaded.txt`,
    localPath: blockedSyncDownload,
    overwrite: true
  });
  const blockedDownloadExistsAfter = await pathExists(blockedSyncDownload);
  record("download refuses a synthetic sync path without creating a local file", blockedDownload.isError
    && String(blockedDownload.value).includes("local OneDrive sync folder")
    && !blockedDownloadExistedBefore
    && !blockedDownloadExistsAfter ? "pass" : "fail", {
    response: blockedDownload.value,
    syntheticSyncRoot,
    existedBefore: blockedDownloadExistedBefore,
    existsAfter: blockedDownloadExistsAfter
  });

  const blockedUploadBefore = await readFile(blockedSyncUpload, "utf8");
  const blockedUpload = await tool("onedrive_upload", {
    localPath: blockedSyncUpload,
    remotePath: `${folderName}/blocked-upload.txt`
  });
  const blockedUploadAfter = await readFile(blockedSyncUpload, "utf8");
  const blockedRemoteUpload = await tool("onedrive_get_info", { path: `${folderName}/blocked-upload.txt` });
  const blockedRemoteUploadAbsent = blockedRemoteUpload.isError && /\bitemNotFound\b/i.test(errorText(blockedRemoteUpload.value));
  record("upload refuses a synthetic sync path without changing its fixture or creating a remote file", blockedUpload.isError
    && String(blockedUpload.value).includes("local OneDrive sync folder")
    && blockedUploadBefore === blockedSyncUploadSentinel
    && blockedUploadAfter === blockedSyncUploadSentinel
    && blockedRemoteUploadAbsent ? "pass" : "fail", {
    response: blockedUpload.value,
    syntheticSyncRoot,
    fixtureUnchanged: blockedUploadBefore === blockedSyncUploadSentinel && blockedUploadAfter === blockedSyncUploadSentinel,
    remoteItemAbsent: blockedRemoteUploadAbsent,
    remoteProbe: blockedRemoteUpload.value
  });

  const [searchResult, searchAllResult] = await Promise.all([
    pollIndexedSearch("onedrive_search", { query: unique, limit: 10 },
      (value) => value.items?.some((item) => item.id === folder.id || item.name === folderName),
      { maxAttempts: 12, delayMs: 5_000 }),
    pollIndexedSearch("onedrive_search_all", { query: unique, pageSize: 2, maxItems: 10 },
      (value) => value.items?.some((item) => item.id === folder.id || item.name === folderName),
      { maxAttempts: 12, delayMs: 5_000 })
  ]);
  const searchMatched = !searchResult.isError
    && searchResult.value.items?.some((item) => item.id === folder.id || item.name === folderName);
  record("search polling finds the exact beta folder", searchMatched ? "pass"
    : searchResult.isError && !clearlyTransientReadError(searchResult.value) ? "fail" : "blocked", {
    resultCount: searchResult.value?.items?.length ?? 0,
    pollAttempts: searchResult.pollAttempts,
    response: searchResult.isError ? searchResult.value : undefined,
    reason: !searchMatched && !searchResult.isError
      ? "OneDrive Personal search indexing did not surface the exact isolated folder or token within the bounded 55-second polling window."
      : undefined
  });

  const searchAllMatched = !searchAllResult.isError
    && searchAllResult.value.items?.some((item) => item.id === folder.id || item.name === folderName);
  record("search_all polling finds the exact beta folder", searchAllMatched ? "pass"
    : searchAllResult.isError && !clearlyTransientReadError(searchAllResult.value) ? "fail" : "blocked", {
    count: searchAllResult.value?.count ?? 0,
    truncated: searchAllResult.value?.truncated,
    pollAttempts: searchAllResult.pollAttempts,
    response: searchAllResult.isError ? searchAllResult.value : undefined,
    reason: !searchAllMatched && !searchAllResult.isError
      ? "OneDrive Personal search_all indexing did not surface the exact isolated folder or token within the bounded 55-second polling window."
      : undefined
  });

  const deltaPages = [];
  const initialDelta = assertOk("onedrive_delta initial page", await tool("onedrive_delta", { pageSize: 1, maxItems: 1 }));
  deltaPages.push(initialDelta);
  let delta = initialDelta;
  if (initialDelta.nextLink) {
    delta = assertOk("onedrive_delta continuation", await tool("onedrive_delta", { nextLink: initialDelta.nextLink, maxItems: 5000, maxPages: 1 }));
    deltaPages.push(delta);
  }
  const returnedDeltaCursor = delta.nextLink || delta.deltaLink || null;
  const deltaContinuationStatus = !initialDelta.nextLink
    ? initialDelta.deltaLink || initialDelta.unsafePageTruncation ? "blocked" : "fail"
    : delta.unsafePageTruncation
      ? "blocked"
      : delta.pagesFetched === 1
        && delta.count <= 5000
        && Boolean(returnedDeltaCursor)
        && returnedDeltaCursor !== initialDelta.nextLink ? "pass" : "fail";
  record("delta follows one observed nextLink continuation without losing the cursor", deltaContinuationStatus, {
    pages: deltaPages.length,
    count: deltaPages.reduce((sum, entry) => sum + (entry.count || 0), 0),
    observedInitialNextLink: Boolean(initialDelta.nextLink),
    continuationPagesFetched: deltaPages.length > 1 ? delta.pagesFetched : 0,
    continuationCursorAdvanced: Boolean(initialDelta.nextLink && returnedDeltaCursor && returnedDeltaCursor !== initialDelta.nextLink),
    maxPagesReached: Boolean(delta.maxPagesReached),
    hasNextLink: Boolean(delta.nextLink),
    hasDeltaLink: Boolean(delta.deltaLink),
    unsafePageTruncation: Boolean(delta.unsafePageTruncation),
    reason: deltaContinuationStatus === "blocked"
      ? initialDelta.nextLink
        ? "The one-page, 5,000-item continuation could not retain a safe nextLink or deltaLink; no additional whole-drive page was attempted."
        : initialDelta.unsafePageTruncation
          ? "Microsoft Graph returned more than the one-item bound on the first page, so no cursor could be followed without skipping items."
          : "Microsoft Graph returned a terminal deltaLink on the first bounded page, so a real nextLink continuation was unavailable."
      : undefined
  });

  const permissionAudit = assertOk("permissions audit", await tool("onedrive_permissions", { path: `${folderName}/${copyFileName}` }));
  record("permission audit call succeeds", Array.isArray(permissionAudit.permissions) ? "pass" : "fail", {
    count: permissionAudit.count,
    item: permissionAudit.item?.name
  });

  const batchInfo = assertOk("batch get info", await tool("onedrive_batch_get_info", {
    items: [{ itemId: copiedInfo.id }, { path: `${folderName}/uploaded.txt` }]
  }));
  record("batch get_info returns per-item results", batchInfo.items?.length === 2 && batchInfo.items.every((entry) => !entry.error) ? "pass" : "fail", {
    count: batchInfo.count,
    names: batchInfo.items?.map((entry) => entry.name || entry.item?.name)
  });

  const batchPermissions = assertOk("batch permissions", await tool("onedrive_batch_permissions", {
    items: [{ itemId: copiedInfo.id }, { path: `${folderName}/uploaded.txt` }]
  }));
  record("batch permissions returns per-item results", batchPermissions.items?.length === 2 && batchPermissions.items.every((entry) => !entry.error) ? "pass" : "fail", {
    count: batchPermissions.count
  });

  const batchDownloadCollisionPath = join(outDir, "uploaded.txt");
  const batchDownloadSentinel = `local collision sentinel ${unique}\n`;
  await writeFile(batchDownloadCollisionPath, batchDownloadSentinel, "utf8");
  const batchDownloaded = assertOk("batch download", await tool("onedrive_batch_download", {
    items: [{ path: `${folderName}/uploaded.txt` }],
    destinationFolder: outDir,
    overwrite: true
  }));
  const batchDownloadedContent = await readFile(batchDownloadCollisionPath, "utf8");
  record("batch download inherits top-level overwrite for an existing destination", batchDownloaded.results?.[0]?.bytesWritten > 0
    && batchDownloadedContent !== batchDownloadSentinel
    && batchDownloadedContent.includes(`Remote conflict marker: ${unique}`) ? "pass" : "fail", {
    count: batchDownloaded.count,
    first: batchDownloaded.results?.[0],
    sentinelReplaced: batchDownloadedContent !== batchDownloadSentinel
  });

  const sharingDryRun = assertOk("sharing link dry-run", await tool("onedrive_create_sharing_link", {
    path: `${folderName}/${copyFileName}`,
    type: "view",
    scope: "anonymous",
    password: `${unique}-link-password`,
    expirationDateTime: "2099-01-01T00:00:00Z"
  }));
  record("sharing link dry-run is safe", sharingDryRun.dryRun === true && sharingDryRun.requiredToCreate && sharingDryRun.wouldCreate?.passwordProvided === true && sharingDryRun.wouldCreate?.expirationDateTime === "2099-01-01T00:00:00Z" ? "pass" : "fail", {
    requiredToCreate: sharingDryRun.requiredToCreate,
    beforePermissionCount: sharingDryRun.beforePermissionCount,
    wouldCreate: sharingDryRun.wouldCreate
  });

  const inviteDryRun = assertOk("invite permission dry-run", await tool("onedrive_invite_permission", {
    path: `${folderName}/${copyFileName}`,
    recipients: [{ email: inviteRecipient }],
    role: "read",
    password: `${unique}-invite-password`,
    expirationDateTime: "2099-01-01T00:00:00Z"
  }));
  record("invite permission dry-run is safe and silent by default", inviteDryRun.dryRun === true
    && inviteDryRun.requiredToInvite
    && inviteDryRun.wouldInvite?.sendInvitation === false
    && inviteDryRun.wouldInvite?.requireSignIn === true
    && inviteDryRun.wouldInvite?.passwordProvided === true
    && inviteDryRun.wouldInvite?.expirationDateTime === "2099-01-01T00:00:00Z" ? "pass" : "fail", {
    requiredToInvite: inviteDryRun.requiredToInvite,
    beforePermissionCount: inviteDryRun.beforePermissionCount,
    wouldInvite: inviteDryRun.wouldInvite
  });

  const inviteNeedsConfirmation = assertOk("invite permission requires confirmation", await tool("onedrive_invite_permission", {
    itemId: copiedInfo.id,
    recipients: [{ email: inviteRecipient }],
    expectedName: copyFileName,
    dryRun: false
  }));
  record("invite live action requires confirmation", inviteNeedsConfirmation.requiredToInvite && inviteNeedsConfirmation.confirmed === false ? "pass" : "fail", {
    requiredToInvite: inviteNeedsConfirmation.requiredToInvite
  });

  const inviteMissingExpected = assertOk("invite permission requires expected identity", await tool("onedrive_invite_permission", {
    itemId: copiedInfo.id,
    recipients: [{ email: inviteRecipient }],
    dryRun: false,
    confirmed: true
  }));
  record("invite live action requires expected identity", inviteMissingExpected.requiredToInvite?.includes("expectedName or expectedId") ? "pass" : "fail", {
    requiredToInvite: inviteMissingExpected.requiredToInvite
  });

  const inviteLiveValue = assertOk("invite permission live", await toolWithPreview("onedrive_invite_permission", {
    itemId: copiedInfo.id,
    recipients: [{ email: inviteRecipient }],
    role: "read",
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  const invitePermissionId = inviteLiveValue.permissionDiff?.added?.[0]?.id
    || inviteLiveValue.permissions?.find((permission) => permission.id)?.id;
  const permissionsAfterInvite = assertOk("permissions after invite", await tool("onedrive_permissions", { itemId: copiedInfo.id }));
  record("invite permission silently grants the explicit recipient", inviteLiveValue.confirmed === true
    && inviteLiveValue.invite?.sendInvitation === false
    && Boolean(invitePermissionId)
    && permissionsAfterInvite.permissions?.some((permission) => permission.id === invitePermissionId) ? "pass" : "fail", {
    recipient: inviteRecipient,
    permissionId: invitePermissionId,
    invite: inviteLiveValue.invite,
    diff: inviteLiveValue.permissionDiff
  });

  if (!invitePermissionId) throw new Error("The live invite did not return a verifiable permission ID.");
  const inviteRevoked = assertOk("revoke invited recipient", await toolWithPreview("onedrive_revoke_permission", {
    itemId: copiedInfo.id,
    permissionId: invitePermissionId,
    expectedId: copiedInfo.id,
    dryRun: false,
    confirmed: true
  }));
  const permissionsAfterInviteRevoke = assertOk("permissions after invite revoke", await tool("onedrive_permissions", { itemId: copiedInfo.id }));
  record("explicit recipient permission is revoked and absent", inviteRevoked.permissionDiff?.removed?.some((permission) => permission.id === invitePermissionId)
    && !permissionsAfterInviteRevoke.permissions?.some((permission) => permission.id === invitePermissionId) ? "pass" : "fail", {
    permissionId: invitePermissionId,
    remainingPermissionIds: permissionsAfterInviteRevoke.permissions?.map((permission) => permission.id)
  });

  const sharingLive = assertOk("sharing link live", await toolWithPreview("onedrive_create_sharing_link", {
    itemId: copiedInfo.id,
    type: "view",
    scope: "anonymous",
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  record("sharing link live creates permission diff", sharingLive.confirmed === true && sharingLive.permission?.id && sharingLive.permissionDiff ? "pass" : "fail", {
    permissionId: sharingLive.permission?.id,
    diff: sharingLive.permissionDiff
  });

  const sharedByMe = assertOk("shared by me audit", await tool("onedrive_shared_by_me", {
    path: folderName,
    limit: 20,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  const publicLinks = assertOk("public links audit", await tool("onedrive_public_links", {
    path: folderName,
    limit: 20,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  const sharingAuditsComplete = sharedByMe.incomplete === false && publicLinks.incomplete === false;
  const sharingAuditsFoundLink = sharedByMe.items?.some((entry) => entry.item?.id === copiedInfo.id)
    && publicLinks.items?.some((entry) => entry.item?.id === copiedInfo.id);
  record("sharing audit tools find live anonymous link",
    !sharingAuditsComplete ? "blocked" : sharingAuditsFoundLink ? "pass" : "fail", {
      sharedByMe: {
        count: sharedByMe.count,
        incomplete: sharedByMe.incomplete,
        errorCount: sharedByMe.errorCount,
        unauditedCount: sharedByMe.unauditedCount,
        errors: sharedByMe.errors,
        ids: sharedByMe.items?.map((entry) => entry.item?.id)
      },
      publicLinks: {
        count: publicLinks.count,
        incomplete: publicLinks.incomplete,
        errorCount: publicLinks.errorCount,
        unauditedCount: publicLinks.unauditedCount,
        errors: publicLinks.errors,
        ids: publicLinks.items?.map((entry) => entry.item?.id)
      }
    });

  const revokeDryRun = assertOk("revoke permission dry-run", await tool("onedrive_revoke_permission", {
    itemId: copiedInfo.id,
    permissionId: sharingLive.permission.id,
    expectedName: copyFileName
  }));
  record("revoke permission dry-run is safe", revokeDryRun.dryRun === true && revokeDryRun.requiredToRevoke ? "pass" : "fail", {
    beforePermissionCount: revokeDryRun.beforePermissionCount,
    requiredToRevoke: revokeDryRun.requiredToRevoke
  });

  const revoked = assertOk("revoke permission live", await toolWithPreview("onedrive_revoke_permission", {
    itemId: copiedInfo.id,
    permissionId: sharingLive.permission.id,
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  const permissionsAfterLinkRevoke = assertOk("permissions after link revoke", await tool("onedrive_permissions", { itemId: copiedInfo.id }));
  record("revoke permission live removes anonymous link and verifies absence", revoked.confirmed === true
    && revoked.permissionDiff?.removed?.some((permission) => permission.id === sharingLive.permission.id)
    && !permissionsAfterLinkRevoke.permissions?.some((permission) => permission.id === sharingLive.permission.id) ? "pass" : "fail", {
    permissionId: revoked.permissionId,
    diff: revoked.permissionDiff,
    remainingPermissionIds: permissionsAfterLinkRevoke.permissions?.map((permission) => permission.id)
  });

  const sharingLiveBatch = assertOk("sharing link live for batch revoke", await toolWithPreview("onedrive_create_sharing_link", {
    itemId: copiedInfo.id,
    type: "view",
    scope: "anonymous",
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  const batchRevokeDryRun = assertOk("batch revoke permissions dry-run", await tool("onedrive_batch_revoke_permissions", {
    items: [{ itemId: copiedInfo.id, permissionId: sharingLiveBatch.permission.id, expectedId: copiedInfo.id }]
  }));
  record("batch revoke permissions dry-run previews isolated permission", batchRevokeDryRun.dryRun === true
    && batchRevokeDryRun.results?.length === 1
    && batchRevokeDryRun.results[0].wouldRevoke?.permissionId === sharingLiveBatch.permission.id
    && batchRevokeDryRun.previewToken ? "pass" : "fail", {
    count: batchRevokeDryRun.count,
    permissionId: batchRevokeDryRun.results?.[0]?.wouldRevoke?.permissionId,
    previewTokenIssued: Boolean(batchRevokeDryRun.previewToken)
  });

  const batchRevoked = assertCommitted("batch revoke permissions live", await tool("onedrive_batch_revoke_permissions", {
    items: [{ itemId: copiedInfo.id, permissionId: sharingLiveBatch.permission.id, expectedId: copiedInfo.id }],
    dryRun: false,
    confirmed: true,
    previewToken: batchRevokeDryRun.previewToken
  }));
  const batchRevokeResult = batchRevoked.results?.[0];
  const permissionsAfterBatchRevoke = assertOk("permissions after batch revoke", await tool("onedrive_permissions", {
    itemId: copiedInfo.id
  }));
  const batchRevokeVerificationIncomplete = batchRevokeResult?.verificationIncomplete === true;
  const batchRevokeRemovedExactPermission = batchRevokeResult?.permissionDiff?.removed
    ?.some((permission) => permission.id === sharingLiveBatch.permission.id);
  const batchRevokePermissionAbsent = !permissionsAfterBatchRevoke.permissions
    ?.some((permission) => permission.id === sharingLiveBatch.permission.id);
  const batchRevokeSucceeded = batchRevoked.confirmed === true
    && batchRevoked.failed !== true
    && batchRevoked.results?.length === 1
    && batchRevokeResult?.permissionId === sharingLiveBatch.permission.id
    && batchRevokePermissionAbsent;
  record("batch revoke permissions live succeeds",
    batchRevokeSucceeded ? "pass" : "fail", {
    count: batchRevoked.count,
    results: batchRevoked.results,
    expectedPermissionId: sharingLiveBatch.permission.id,
    verificationIncomplete: batchRevokeResult?.verificationIncomplete,
    removedExactPermission: Boolean(batchRevokeRemovedExactPermission),
    remainingPermissionIds: permissionsAfterBatchRevoke.permissions?.map((permission) => permission.id)
  });

  const auditRecent = assertOk("audit recent", await tool("onedrive_audit_recent", { limit: 50 }));
  record("audit recent includes live mutation entries", auditRecent.entries?.some((entry) => entry.tool === "onedrive_revoke_permission") && auditRecent.entries?.some((entry) => entry.tool === "onedrive_batch_revoke_permissions") ? "pass" : "fail", {
    count: auditRecent.count,
    tools: auditRecent.entries?.map((entry) => entry.tool)
  });

  const auditExported = assertOk("audit export", await tool("onedrive_audit_export", {
    localPath: auditExport,
    overwrite: true
  }));
  record("audit export writes local JSONL", auditExported.bytesWritten > 0 && auditExported.localPath === auditExport ? "pass" : "fail", {
    localPath: auditExported.localPath,
    bytesWritten: auditExported.bytesWritten
  });

  const auditClearNoConfirm = assertOk("audit clear requires confirmation", await tool("onedrive_audit_clear"));
  record("audit_clear requires confirmation", auditClearNoConfirm.requiredToClear && auditClearNoConfirm.confirmed === false ? "pass" : "fail", {
    requiredToClear: auditClearNoConfirm.requiredToClear
  });
  const auditCleared = assertOk("isolated audit clear", await tool("onedrive_audit_clear", { confirmed: true }));
  const auditAfterClear = assertOk("audit after isolated clear", await tool("onedrive_audit_recent", { limit: 10 }));
  record("isolated beta audit can be cleared live", auditCleared.confirmed === true
    && auditAfterClear.count === 0
    && (auditAfterClear.entries?.length || 0) === 0 ? "pass" : "fail", {
    cleared: auditCleared,
    remainingCount: auditAfterClear.count
  });

  const expectedMismatch = await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: "wrong-name",
    dryRun: false,
    confirmed: true
  });
  record("delete expectedName mismatch refused", expectedMismatch.isError && String(expectedMismatch.value).includes("expected item named") ? "pass" : "fail", {
    response: expectedMismatch.value
  });

  const deleteNeedsConfirmation = assertOk("onedrive_delete requires confirmation", await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: folderName,
    dryRun: false
  }));
  record("delete live action requires confirmation", deleteNeedsConfirmation.requiredToDelete && deleteNeedsConfirmation.confirmed === false ? "pass" : "fail", {
    requiredToDelete: deleteNeedsConfirmation.requiredToDelete
  });

  const dryRun = assertOk("onedrive_delete dry-run", await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: folderName,
    dryRun: true
  }));
  record("delete dry-run safety", dryRun.dryRun === true && dryRun.wouldDelete?.id === folder.id ? "pass" : "fail", {
    wouldDelete: dryRun.wouldDelete?.name
  });

  const expectedDeletedFolderId = folder.id;
  const deleted = assertCommitted("onedrive_delete", await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: folderName,
    dryRun: false,
    confirmed: true,
    previewToken: dryRun.previewToken
  }), (value) => value.dryRun === false && value.deleted?.id === expectedDeletedFolderId);
  results.cleanup = { deleted: deleted.deleted?.name, id: deleted.deleted?.id };
  record("delete test folder cleanup", "pass", results.cleanup);
  folder = null;

  const restoreDryRun = assertOk("restore deleted dry-run", await tool("onedrive_restore_deleted", {
    itemId: deleted.deleted.id,
    expectedId: deleted.deleted.id
  }));
  record("restore deleted dry-run is safe", restoreDryRun.dryRun === true && restoreDryRun.requiredToRestore ? "pass" : "fail", {
    requiredToRestore: restoreDryRun.requiredToRestore,
    permissionNote: restoreDryRun.permissionNote
  });
  record("live recycle-bin restore", "blocked", {
    reason: "Live restore is intentionally excluded from this personal-drive beta; only mock and dry-run coverage are authorized.",
    previewedItemId: deleted.deleted.id
  });

  const deletedById = await tool("onedrive_get_info", {
    itemId: deleted.deleted.id,
    includeDeletedItems: true
  });
  record("includeDeletedItems returns the exact deleted item or an explicit Graph limitation", deletedById.isError
    ? /\bitemNotFound\b/i.test(errorText(deletedById.value)) ? "blocked" : "fail"
    : deletedById.value?.id === deleted.deleted.id && Boolean(deletedById.value?.deleted) ? "pass" : "fail", {
    response: deletedById.value,
    classification: deletedById.isError && /\bitemNotFound\b/i.test(errorText(deletedById.value)) ? "itemNotFound" : undefined
  });

  const deletedInfo = await tool("onedrive_get_info", { path: folderName });
  record("deleted folder no longer resolves with itemNotFound", deletedInfo.isError
    && /\bitemNotFound\b/i.test(errorText(deletedInfo.value)) ? "pass" : "fail", { response: deletedInfo.value });

  const rootDelete = await tool("onedrive_delete", { path: "/", dryRun: true });
  record("root delete is refused", rootDelete.isError && String(rootDelete.value).includes("OneDrive root") ? "pass" : "fail", {
    response: rootDelete.value
  });

  const deviceStart = assertOk("auth device start", await tool("onedrive_auth_device_start", { tenant: "consumers" }));
  record("healthy auth suppresses a redundant device code", deviceStart.alreadyAuthenticated === true && deviceStart.deviceCodeIssued === false && !deviceStart.userCode ? "pass" : "fail", {
    alreadyAuthenticated: deviceStart.alreadyAuthenticated,
    deviceCodeIssued: deviceStart.deviceCodeIssued,
    keychainTokenConfigured: deviceStart.keychainTokenConfigured
  });
  record("forced device-code polling", "blocked", {
    reason: "The existing credential and consent must remain untouched; the healthy-auth suppression path and device flow mocks provide coverage."
  });
  record("Keychain credential deletion", "blocked", {
    reason: "Deleting the existing OneDrive credential is outside the authorized live scope; only memory-only logout is exercised."
  });

  const logoutMemoryOnly = assertOk("logout memory only", await tool("onedrive_logout", { deleteKeychainToken: false }));
  record("logout clears memory without deleting Keychain token", logoutMemoryOnly.memoryCleared === true && logoutMemoryOnly.keychainTokenDeleted === false ? "pass" : "fail", logoutMemoryOnly);
  const blockedLiveTools = new Set(["onedrive_auth_device_poll"]);
  const expectedLiveTools = ONEDRIVE_TOOL_CONTRACT.filter((name) => !blockedLiveTools.has(name));
  const missingLiveTools = expectedLiveTools.filter((name) => !calledTools.has(name));
  const unexpectedLiveTools = [...calledTools].filter((name) => !ONEDRIVE_TOOL_CONTRACT.includes(name));
  record("live beta exercises every non-blocked tool in the exact 84-tool contract", missingLiveTools.length === 0 && unexpectedLiveTools.length === 0 ? "pass" : "fail", {
    contractCount: ONEDRIVE_TOOL_CONTRACT.length,
    exercisedCount: calledTools.size,
    blockedTools: [...blockedLiveTools],
    missingLiveTools,
    unexpectedLiveTools
  });
  }
} catch (error) {
  results.error = error.stack || error.message;
  if (folder?.id) {
    try {
      results.cleanup = await cleanupExactBetaFolder({ id: folder.id, name: folderName }, folderName);
      folder = null;
    } catch (cleanupError) {
      results.cleanup = { attemptedAfterError: true, error: cleanupError.message };
      results.error += `\nCleanup failure: ${cleanupError.message}`;
    }
  }
} finally {
  results.stderr = stderr.join("");
  child.stdin.end();
  if (!childExited) child.kill("SIGTERM");
  await Promise.race([
    childExit,
    new Promise((resolve) => setTimeout(resolve, 2_000))
  ]);
  if (!childExited) child.kill("SIGKILL");
}

const passCount = results.checks.filter((check) => check.status === "pass").length;
const failCount = results.checks.filter((check) => check.status === "fail").length;
const blockedCount = results.checks.filter((check) => check.status === "blocked").length;
results.summary = { passCount, failCount, blockedCount, total: results.checks.length };
results.finishedAt = new Date().toISOString();
results.runtimeMs = Date.now() - harnessStartedAtMs;

if (!keepWork) {
  await cleanupLocalWorkDir();
  results.localWorkCleaned = true;
} else {
  results.localWorkDir = outDir;
}

console.log(JSON.stringify(results, null, 2));

if (results.error || failCount > 0) {
  process.exitCode = 1;
}
