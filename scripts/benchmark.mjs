#!/usr/bin/env node

import { spawn } from "node:child_process";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");

function parseCliArgs(argv = []) {
  return Object.fromEntries(argv.map((arg) => {
    const [key, ...rest] = arg.replace(/^--/, "").split("=");
    return [key, rest.length ? rest.join("=") : true];
  }));
}

function parseBooleanFlag(value, name, defaultValue = false) {
  if (value === undefined) return defaultValue;
  if (typeof value === "boolean") return value;
  const normalized = String(value).trim().toLowerCase();
  if (["1", "true", "yes", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "off", ""].includes(normalized)) return false;
  throw new Error(`--${name} expects a boolean value, got ${value}.`);
}

function rejectPendingRequests(pendingRequests, error) {
  for (const [id, waiter] of pendingRequests.entries()) {
    pendingRequests.delete(id);
    clearTimeout(waiter.timeout);
    waiter.reject(error);
  }
}

function benchmarkSucceeded(results = []) {
  return results.length > 0 && results.every((entry) => !entry.isError);
}

async function runSelfCheck() {
  const probePending = new Map();
  const probe = new Promise((resolve, reject) => {
    probePending.set(1, {
      resolve,
      reject,
      timeout: setTimeout(() => reject(new Error("probe timeout was not cleared")), 1000)
    });
  });
  rejectPendingRequests(probePending, new Error("probe child exit"));
  let childExitRejected = false;
  try {
    await probe;
  } catch (error) {
    childExitRejected = error.message === "probe child exit";
  }
  const checks = {
    clearBareFlag: parseBooleanFlag(true, "clear") === true,
    clearExplicitTrue: parseBooleanFlag("true", "clear") === true,
    clearExplicitFalse: parseBooleanFlag("false", "clear", true) === false,
    clearNumericFalse: parseBooleanFlag("0", "clear", true) === false,
    toolErrorsFailBenchmark: benchmarkSucceeded([{ isError: false }, { isError: true }]) === false,
    successfulToolsPassBenchmark: benchmarkSucceeded([{ isError: false }]) === true,
    pendingRejectedOnChildExit: childExitRejected && probePending.size === 0
  };
  const ok = Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, checks }, null, 2));
  return ok;
}

const args = parseCliArgs(process.argv.slice(2));
if (parseBooleanFlag(args["self-check"], "self-check")) {
  process.exit(await runSelfCheck() ? 0 : 1);
}

const query = String(args.query || "project plan");
const maxItems = Number(args.maxItems || 1500);
const maxFolders = Number(args.maxFolders || 250);
const maxFiles = Number(args.maxFiles || 50);
const maxBytesPerFile = Number(args.maxBytesPerFile || 262144);
const searchConcurrency = Number(args.searchConcurrency || 2);
const clear = parseBooleanFlag(args.clear, "clear");

let nextId = 1;
const pending = new Map();
let stdoutBuffer = "";
let stderrBuffer = "";
let childExited = false;
let childExitError = null;

const child = spawn(process.execPath, [serverPath], {
  cwd: pluginRoot,
  env: process.env,
  stdio: ["pipe", "pipe", "pipe"]
});

child.once("error", (error) => {
  childExitError = new Error(`Could not start OneDrive MCP child: ${error.message}`);
  rejectPendingRequests(pending, childExitError);
});

child.once("exit", (code, signal) => {
  childExited = true;
  childExitError = new Error(`OneDrive MCP child exited before completing pending requests (code=${code ?? "null"}, signal=${signal || "none"}).`);
  rejectPendingRequests(pending, childExitError);
});

child.stdin.on("error", (error) => {
  childExitError ||= new Error(`OneDrive MCP child stdin failed: ${error.message}`);
  rejectPendingRequests(pending, childExitError);
});

child.stdout.on("data", (chunk) => {
  stdoutBuffer += chunk.toString("utf8");
  for (;;) {
    const newline = stdoutBuffer.indexOf("\n");
    if (newline < 0) break;
    const line = stdoutBuffer.slice(0, newline);
    stdoutBuffer = stdoutBuffer.slice(newline + 1);
    if (!line.trim()) continue;
    let message;
    try {
      message = JSON.parse(line);
    } catch {
      continue;
    }
    const waiter = pending.get(message.id);
    if (!waiter) continue;
    pending.delete(message.id);
    clearTimeout(waiter.timeout);
    if (message.error) waiter.reject(new Error(message.error.message || JSON.stringify(message.error)));
    else waiter.resolve(message.result);
  }
});

child.stderr.on("data", (chunk) => {
  stderrBuffer += chunk.toString("utf8");
});

function request(method, params = {}) {
  if (childExited || childExitError || !child.stdin.writable) {
    return Promise.reject(childExitError || new Error("OneDrive MCP child is not available."));
  }
  const id = nextId;
  nextId += 1;
  const promise = new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      if (!pending.has(id)) return;
      pending.delete(id);
      reject(new Error(`${method} timed out`));
    }, 120_000).unref?.();
    pending.set(id, { resolve, reject, timeout });
  });
  try {
    child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n", (error) => {
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

async function tool(name, toolArgs = {}) {
  const startedAt = Date.now();
  const response = await request("tools/call", { name, arguments: toolArgs });
  const text = response.content?.[0]?.text || "{}";
  let value;
  try {
    value = JSON.parse(text);
  } catch {
    value = text;
  }
  return {
    name,
    isError: Boolean(response.isError),
    wallMs: Date.now() - startedAt,
    value
  };
}

function summary(result) {
  const value = result.value || {};
  const nestedSummary = value.summary || {};
  return {
    tool: result.name,
    isError: result.isError,
    wallMs: result.wallMs,
    durationMs: value.durationMs ?? nestedSummary.durationMs,
    returned: value.returned ?? nestedSummary.returned ?? value.items?.length,
    mode: value.mode,
    itemCount: value.itemCount ?? value.cache?.itemCount,
    liveSearchDurationMs: nestedSummary.liveSearchDurationMs,
    scanDurationMs: nestedSummary.scanDurationMs,
    graphSearchCalls: nestedSummary.graphSearchCalls,
    searchTermsPlanned: nestedSummary.searchTermsPlanned,
    searchTermsExecuted: nestedSummary.searchTermsExecuted,
    searchTermsSkipped: nestedSummary.searchTermsSkipped,
    searchStopReason: nestedSummary.searchStopReason,
    scanAttempts: nestedSummary.scanAttempts,
    cacheCandidates: nestedSummary.cacheCandidates,
    contentIndexCandidates: nestedSummary.contentIndexCandidates,
    considered: value.considered,
    eligible: value.eligible,
    selected: value.selected,
    indexed: value.indexed,
    reused: value.reused,
    failed: value.failed,
    graphContentReadsAttempted: value.graphContentReadsAttempted,
    details: result.isError ? value : undefined
  };
}

function progress(step, details = {}) {
  console.error(JSON.stringify({
    ts: new Date().toISOString(),
    step,
    ...details
  }));
}

try {
  await request("initialize", {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: { name: "onedrive-benchmark", version: "1" }
  });

  const results = [];
  if (clear) {
    progress("clear-local-cache");
    results.push(await tool("onedrive_cache_clear"));
    results.push(await tool("onedrive_content_index_clear"));
  }

  progress("cold-find", { query, maxItems, maxFolders });
  const coldFind = await tool("onedrive_find", {
    query,
    useCache: false,
    useContentIndex: false,
    searchConcurrency,
    scanMaxItems: maxItems,
    scanMaxFolders: maxFolders
  });
  results.push(coldFind);

  progress("cache-refresh", { maxItems, maxFolders });
  results.push(await tool("onedrive_cache_refresh", {
    mode: "auto",
    maxItems,
    maxFolders
  }));

  progress("warm-find", { query, maxItems, maxFolders });
  const warmFind = await tool("onedrive_find", {
    query,
    useCache: true,
    useContentIndex: false,
    searchConcurrency,
    scanMaxItems: maxItems,
    scanMaxFolders: maxFolders
  });
  results.push(warmFind);

  progress("content-index-refresh", { maxFiles, maxBytesPerFile });
  results.push(await tool("onedrive_content_index_refresh", {
    maxFiles,
    maxBytesPerFile
  }));

  progress("content-search", { query });
  const contentSearch = await tool("onedrive_content_search", {
    query,
    maxResults: 10
  });
  results.push(contentSearch);

  const selectedId = contentSearch.value?.items?.[0]?.id || warmFind.value?.items?.[0]?.id || coldFind.value?.items?.[0]?.id;
  if (selectedId) {
    progress("preview-selected", { selectedId });
    results.push(await tool("onedrive_preview", {
      itemId: selectedId,
      maxBytes: 4096
    }));
  }

  const ok = benchmarkSucceeded(results);
  console.log(JSON.stringify({
    ok,
    query,
    clear,
    caps: { maxItems, maxFolders, maxFiles, maxBytesPerFile, searchConcurrency },
    summary: results.map(summary),
    note: "Use --clear to include local cache clear before the cold run. All remote operations are read-only, but cache/index files are local writes."
  }, null, 2));
  if (!ok) process.exitCode = 1;
} catch (error) {
  console.error(JSON.stringify({
    ok: false,
    error: error.message,
    stderr: stderrBuffer.trim() || undefined
  }, null, 2));
  process.exitCode = 1;
} finally {
  if (!childExited) child.kill();
}
