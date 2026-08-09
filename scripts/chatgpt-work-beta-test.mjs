#!/usr/bin/env node

import { spawn } from "node:child_process";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptsRoot = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(scriptsRoot, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");
const runIdPattern = /^codex-beta-[a-z0-9](?:[a-z0-9._-]*[a-z0-9])?$/iu;
const expectedTools = [
  "fetch",
  "onedrive_commit_actions",
  "onedrive_create_folder",
  "onedrive_delete",
  "onedrive_export_file",
  "onedrive_invite_permission",
  "onedrive_office_batch_transform",
  "onedrive_office_capabilities",
  "onedrive_open_files",
  "onedrive_patch_text",
  "onedrive_preview_actions",
  "onedrive_read_actions",
  "onedrive_restore_deleted",
  "onedrive_upload_file",
  "onedrive_write_text"
].sort();

function parseArgs(argv = []) {
  const parsed = {};
  for (const argument of argv) {
    if (!argument.startsWith("--")) throw new Error(`Unexpected positional argument: ${argument}`);
    const [key, ...valueParts] = argument.slice(2).split("=");
    if (!key || Object.hasOwn(parsed, key)) throw new Error(`Invalid or duplicate option: --${key || "unknown"}`);
    parsed[key] = valueParts.length ? valueParts.join("=") : true;
  }
  for (const key of Object.keys(parsed)) {
    if (!["run-id", "report", "keep-work", "self-check"].includes(key)) throw new Error(`Unknown option: --${key}`);
  }
  return parsed;
}

function parseBoolean(value, name) {
  if (value === undefined) return false;
  if (typeof value === "boolean") return value;
  const normalized = String(value).trim().toLowerCase();
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off", ""].includes(normalized)) return false;
  throw new Error(`--${name} expects a boolean value.`);
}

function parseRunId(value) {
  if (typeof value !== "string" || value.length > 80 || !runIdPattern.test(value)) {
    throw new Error("--run-id is required and must match codex-beta-[A-Za-z0-9._-].");
  }
  return value;
}

function normalize(value = "") {
  return String(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[_]+/gu, " ")
    .replace(/[^\p{L}\p{N}.-]+/gu, " ")
    .replace(/\s+/gu, " ")
    .trim();
}

function exactTitle(expected) {
  const key = normalize(expected);
  return (item) => normalize(item?.title || item?.name) === key;
}

function containsMetadata(term) {
  const key = normalize(term);
  return (item) => normalize(`${item?.title || item?.name || ""} ${item?.path || ""}`).includes(key);
}

function evaluateSearchQuality(searches, cases) {
  const reciprocalRanks = [];
  let unrelatedTopFive = 0;
  const details = [];
  for (const [index, testCase] of cases.entries()) {
    const operation = searches[index];
    const items = operation?.value?.items || [];
    const rank = items.findIndex(testCase.matches) + 1;
    reciprocalRanks.push(rank ? 1 / rank : 0);
    unrelatedTopFive += items.slice(0, 5).filter((item) => !testCase.matches(item)).length;
    details.push({
      query: testCase.query,
      rank: rank || null,
      returned: items.length,
      topTitle: items[0]?.title || items[0]?.name || null,
      rankedSearch: operation?.value?.rankedSearch === true,
      searchMode: operation?.value?.searchMode || null,
      metadataQualifiedResults: operation?.value?.metadataQualifiedResults ?? null,
      verifiedContentResults: operation?.value?.verifiedContentResults ?? null,
      rawGraphContentOnlyResultsSuppressed: operation?.value?.rawGraphContentOnlyResultsSuppressed ?? null,
      durationMs: operation?.durationMs ?? null
    });
  }
  return {
    queryCount: cases.length,
    exactAtOne: details.filter((entry) => entry.rank === 1).length,
    mrrAt10: reciprocalRanks.reduce((sum, value) => sum + value, 0) / Math.max(1, cases.length),
    unrelatedTopFive,
    details
  };
}

const args = parseArgs(process.argv.slice(2));
const selfCheck = parseBoolean(args["self-check"], "self-check");
if (selfCheck) {
  const fixture = evaluateSearchQuality([
    { durationMs: 1, value: { rankedSearch: true, searchMode: "ranked", items: [{ title: "Qaldris", path: "Ventures/Qaldris" }] } }
  ], [{ query: "Qaldris", matches: containsMetadata("Qaldris") }]);
  const checks = {
    validRunIdAccepted: parseRunId("codex-beta-work-20260809") === "codex-beta-work-20260809",
    invalidRunIdRejected: (() => { try { parseRunId("unsafe/run"); return false; } catch { return true; } })(),
    exactTitleMatcher: exactTitle("Example_File.docx")({ title: "Example File.docx" }) === true,
    qualityMetrics: fixture.mrrAt10 === 1 && fixture.exactAtOne === 1 && fixture.unrelatedTopFive === 0,
    exactWorkToolCount: expectedTools.length === 15
  };
  const ok = Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, checks }, null, 2));
  process.exit(ok ? 0 : 1);
}

const runId = parseRunId(args["run-id"]);
const keepWork = parseBoolean(args["keep-work"], "keep-work");
const reportPath = args.report === undefined ? null : resolve(String(args.report));
const runRoot = join(pluginRoot, "work", "onedrive-beta", runId, "chatgpt-work");
const storageRoot = join(runRoot, "plugin-state");
const startedAtMs = Date.now();
const startedAt = new Date(startedAtMs).toISOString();
await mkdir(storageRoot, { recursive: true, mode: 0o700 });

const child = spawn(process.execPath, [serverPath], {
  cwd: pluginRoot,
  env: {
    ...process.env,
    ONEDRIVE_TOOL_PROFILE: "chatgpt",
    ONEDRIVE_STORAGE_ROOT: storageRoot,
    ONEDRIVE_CACHE_ROOT: join(storageRoot, "cache"),
    ONEDRIVE_PERFORMANCE_LOG: "1"
  },
  stdio: ["pipe", "pipe", "pipe"]
});

let nextId = 1;
let stdoutBuffer = "";
let stderr = "";
const pending = new Map();
let childError = null;

function rejectPending(error) {
  for (const [id, waiter] of pending) {
    pending.delete(id);
    clearTimeout(waiter.timeout);
    waiter.reject(error);
  }
}

child.once("error", (error) => {
  childError = error;
  rejectPending(error);
});
child.once("exit", (code, signal) => {
  if (code === 0 || child.killed) return;
  childError = new Error(`ChatGPT Work beta child exited early (code=${code}, signal=${signal || "none"}).`);
  rejectPending(childError);
});
child.stderr.on("data", (chunk) => { stderr += chunk.toString("utf8"); });
child.stdout.on("data", (chunk) => {
  stdoutBuffer += chunk.toString("utf8");
  for (;;) {
    const newline = stdoutBuffer.indexOf("\n");
    if (newline < 0) break;
    const line = stdoutBuffer.slice(0, newline).trim();
    stdoutBuffer = stdoutBuffer.slice(newline + 1);
    if (!line) continue;
    let message;
    try { message = JSON.parse(line); } catch { continue; }
    const waiter = pending.get(message.id);
    if (!waiter) continue;
    pending.delete(message.id);
    clearTimeout(waiter.timeout);
    if (message.error) waiter.reject(new Error(message.error.message || JSON.stringify(message.error)));
    else waiter.resolve(message.result);
  }
});

function request(method, params = {}, timeoutMs = 180_000) {
  if (childError || !child.stdin.writable) return Promise.reject(childError || new Error("ChatGPT Work beta child is unavailable."));
  const id = nextId++;
  const promise = new Promise((resolvePromise, reject) => {
    const timeout = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`${method} timed out after ${timeoutMs}ms.`));
    }, timeoutMs);
    pending.set(id, { resolve: resolvePromise, reject, timeout });
  });
  child.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
  return promise;
}

async function callTool(name, toolArgs) {
  const started = Date.now();
  const result = await request("tools/call", { name, arguments: toolArgs });
  const text = result.content?.[0]?.text || "{}";
  let value;
  try { value = JSON.parse(text); } catch { value = text; }
  if (result.isError) throw new Error(`${name} failed: ${typeof value === "string" ? value : JSON.stringify(value)}`);
  return { value, wallMs: Date.now() - started };
}

const cases = [
  { query: "Qaldris", matches: containsMetadata("Qaldris") },
  { query: "HACCP Study Guide.pdf", matches: exactTitle("HACCP Study Guide.pdf") },
  { query: "Digital Quality Management Insights (1).docx", matches: exactTitle("Digital Quality Management Insights (1).docx") }
];
const report = { runId, mode: "chatgpt-work-read-only", startedAt, checks: [], remoteMutations: 0 };

function check(name, passed, details = {}) {
  report.checks.push({ name, status: passed ? "pass" : "fail", details });
}

try {
  const initialized = await request("initialize", {
    protocolVersion: "2025-11-25",
    capabilities: {},
    clientInfo: { name: "onedrive-chatgpt-work-beta", version: "1" }
  });
  check("focused ChatGPT Work server initializes", initialized.serverInfo?.name === "onedrive" && String(initialized.serverInfo?.version || "").includes(".chatgpt."), initialized.serverInfo);

  const listed = await request("tools/list");
  const names = (listed.tools || []).map((tool) => tool.name).sort();
  check("tools/list matches the exact reviewed 15-tool Work surface", JSON.stringify(names) === JSON.stringify(expectedTools), { expected: expectedTools, actual: names });

  const read = await callTool("onedrive_read_actions", {
    actions: [
      { operation: "list", path: "/", limit: 10, format: "compact" },
      ...cases.map(({ query }) => ({ operation: "search", query, limit: 10, format: "compact" }))
    ]
  });
  const operations = read.value.results || [];
  const rootList = operations[0];
  const searches = operations.slice(1);
  check("combined Work read returns root listing and every independent ranked search", operations.length === cases.length + 1 && operations.every((entry) => entry.isError === false) && (rootList?.value?.items || []).length > 0, { operationCount: operations.length, rootItems: rootList?.value?.items?.length || 0, wallMs: read.wallMs });

  const quality = evaluateSearchQuality(searches, cases);
  check("ranked Work search meets beta relevance thresholds", quality.exactAtOne === cases.length && quality.mrrAt10 === 1 && quality.unrelatedTopFive === 0 && quality.details.every((entry) => entry.rankedSearch && entry.searchMode === "ranked"), quality);

  const fetchCandidate = searches[1]?.value?.items?.find(cases[1].matches);
  if (fetchCandidate?.id) {
    const fetched = await callTool("fetch", { id: fetchCandidate.id });
    check("selected Work search result is readable through fetch", fetched.value?.title === "HACCP Study Guide.pdf" && String(fetched.value?.text || "").length > 0, { title: fetched.value?.title || null, textLength: String(fetched.value?.text || "").length, wallMs: fetched.wallMs, previewSource: fetched.value?.metadata?.previewSource || null });
  } else {
    check("selected Work search result is readable through fetch", false, { reason: "HACCP Study Guide.pdf was not returned." });
  }
} catch (error) {
  check("beta harness completed without an exception", false, { error: error.stack || error.message });
} finally {
  child.kill("SIGTERM");
  report.finishedAt = new Date().toISOString();
  report.runtimeMs = Date.now() - startedAtMs;
  report.stderr = stderr.slice(-20_000);
  report.summary = {
    passCount: report.checks.filter((entry) => entry.status === "pass").length,
    failCount: report.checks.filter((entry) => entry.status === "fail").length,
    total: report.checks.length
  };
  if (reportPath) {
    await mkdir(dirname(reportPath), { recursive: true });
    await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
  }
  if (!keepWork) await rm(runRoot, { recursive: true, force: true });
}

console.log(JSON.stringify({
  ok: report.summary.failCount === 0,
  runId,
  mode: report.mode,
  summary: report.summary,
  quality: report.checks.find((entry) => entry.name.includes("relevance thresholds"))?.details || null,
  reportPath
}, null, 2));
process.exit(report.summary.failCount === 0 ? 0 : 1);
