#!/usr/bin/env node

import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");

const requests = [];
const counters = new Map();

function count(key) {
  const next = (counters.get(key) || 0) + 1;
  counters.set(key, next);
  return next;
}

function json(res, status, body, headers = {}) {
  res.writeHead(status, { "Content-Type": "application/json", ...headers });
  res.end(JSON.stringify(body));
}

function empty(res, status, headers = {}) {
  res.writeHead(status, headers);
  res.end();
}

function item(id, name, extra = {}) {
  return {
    id,
    name,
    webUrl: `https://example.test/${encodeURIComponent(id)}`,
    size: 12,
    createdDateTime: "2026-07-04T00:00:00Z",
    lastModifiedDateTime: "2026-07-04T00:00:00Z",
    parentReference: { path: "/drive/root:" },
    file: { mimeType: "text/plain" },
    ...extra
  };
}

const graph = createServer((req, res) => {
  const url = new URL(req.url, "http://127.0.0.1");
  const path = url.pathname;
  requests.push({ method: req.method, path, url: req.url });

  if (req.method === "GET" && path === "/v1.0/me/drive") {
    if (count("drive") === 1) return json(res, 429, { error: { code: "tooManyRequests", message: "retry please" } }, { "Retry-After": "0" });
    return json(res, 200, { id: "drive", driveType: "personal", name: "Mock OneDrive" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delete-target") {
    return json(res, 200, item("delete-target", "delete-me.txt"));
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/delete-target") {
    count("delete");
    return empty(res, 204);
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/copy-src") {
    return json(res, 200, item("copy-src", "copy-source.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root") {
    return json(res, 200, { id: "root", name: "root", root: {}, folder: {} });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/copy-src/copy") {
    return empty(res, 202, { Location: `http://127.0.0.1:${graph.address().port}/monitor/copy` });
  }

  if (req.method === "GET" && path === "/monitor/copy") {
    return empty(res, 303, { Location: `http://127.0.0.1:${graph.address().port}/v1.0/me/drive/items/copied` });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root:/Documents/Pictures:") {
    return json(res, 200, item("pictures", "Pictures", { folder: { childCount: 0 }, file: undefined }));
  }

  json(res, 404, { error: { code: "notFound", message: `${req.method} ${req.url}` } });
});

await new Promise((resolve) => graph.listen(0, "127.0.0.1", resolve));
const graphBaseUrl = `http://127.0.0.1:${graph.address().port}/v1.0`;

const child = spawn(process.execPath, [serverPath], {
  cwd: pluginRoot,
  env: {
    ...process.env,
    ONEDRIVE_TEST_ACCESS_TOKEN: "mock-token",
    ONEDRIVE_GRAPH_BASE_URL: graphBaseUrl
  },
  stdio: ["pipe", "pipe", "pipe"]
});

let nextId = 1;
let buffer = "";
const pending = new Map();
const stderr = [];

child.stderr.on("data", (chunk) => stderr.push(chunk.toString()));
child.stdout.on("data", (chunk) => {
  buffer += chunk.toString();
  for (;;) {
    const newline = buffer.indexOf("\n");
    if (newline === -1) break;
    const line = buffer.slice(0, newline).trim();
    buffer = buffer.slice(newline + 1);
    if (!line) continue;
    const message = JSON.parse(line);
    const waiter = pending.get(message.id);
    if (waiter) {
      pending.delete(message.id);
      clearTimeout(waiter.timeout);
      waiter.resolve(message);
    }
  }
});

function request(method, params = {}) {
  const id = nextId++;
  const payload = { jsonrpc: "2.0", id, method, params };
  const promise = new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`Timed out waiting for ${method}`));
    }, 20_000);
    pending.set(id, { resolve, reject, timeout });
  });
  child.stdin.write(`${JSON.stringify(payload)}\n`);
  return promise;
}

async function tool(name, args = {}) {
  const response = await request("tools/call", { name, arguments: args });
  if (response.error) throw new Error(response.error.message);
  const text = response.result?.content?.[0]?.text ?? "";
  let value = text;
  try {
    value = JSON.parse(text);
  } catch {
    // Keep plain text.
  }
  return { isError: Boolean(response.result?.isError), value };
}

function assert(condition, message, details = {}) {
  if (!condition) {
    const error = new Error(message);
    error.details = details;
    throw error;
  }
}

const results = [];
async function check(name, fn) {
  try {
    const details = await fn();
    results.push({ name, status: "pass", details });
  } catch (error) {
    results.push({ name, status: "fail", error: error.message, details: error.details || {} });
  }
}

try {
  await request("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "mock-graph-test", version: "1.0.0" } });

  await check("GET requests retry 429 once", async () => {
    const result = await tool("onedrive_drive");
    assert(!result.isError, "onedrive_drive returned an error", result);
    assert(result.value.name === "Mock OneDrive", "drive response did not come from mock Graph", result.value);
    assert(counters.get("drive") === 2, "expected one retry after 429", { count: counters.get("drive") });
    return { attempts: counters.get("drive") };
  });

  await check("preset traversal is blocked before Graph", async () => {
    const before = requests.length;
    const result = await tool("onedrive_get_info", { preset: "documents", relativePath: "../Pictures" });
    assert(result.isError, "traversal call should fail");
    assert(String(result.value).includes("unsafe path segment"), "unexpected traversal error", result);
    assert(requests.length === before, "traversal should not reach mock Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("normal preset path still reaches Graph", async () => {
    const result = await tool("onedrive_get_info", { preset: "documents", relativePath: "Pictures" });
    assert(!result.isError, "safe preset path should succeed", result);
    assert(result.value.id === "pictures", "unexpected item", result.value);
    return { id: result.value.id };
  });

  await check("delete requires confirmation before DELETE", async () => {
    const result = await tool("onedrive_delete", { itemId: "delete-target", expectedName: "delete-me.txt", dryRun: false });
    assert(!result.isError, "delete confirmation guard should return a structured result", result);
    assert(result.value.requiredToDelete, "delete did not require confirmation", result.value);
    assert(!counters.get("delete"), "DELETE should not be called without confirmation", { deleteCount: counters.get("delete") });
    return { requiredToDelete: result.value.requiredToDelete };
  });

  await check("confirmed delete sends one DELETE", async () => {
    const result = await tool("onedrive_delete", { itemId: "delete-target", expectedName: "delete-me.txt", dryRun: false, confirmed: true });
    assert(!result.isError, "confirmed delete should succeed", result);
    assert(counters.get("delete") === 1, "expected exactly one DELETE", { deleteCount: counters.get("delete") });
    return { deleteCount: counters.get("delete") };
  });

  await check("copy monitor exposes manual 303", async () => {
    const result = await tool("onedrive_copy", { itemId: "copy-src", waitForCompletion: true, timeoutSeconds: 5 });
    assert(!result.isError, "copy should succeed", result);
    assert(result.value.monitor?.status === 303, "copy monitor did not preserve 303", result.value.monitor);
    assert(result.value.monitor?.resourceLocation?.includes("/v1.0/me/drive/items/copied"), "missing resource location", result.value.monitor);
    return result.value.monitor;
  });
} finally {
  child.stdin.end();
  child.kill("SIGTERM");
  graph.close();
}

const failCount = results.filter((result) => result.status === "fail").length;
console.log(JSON.stringify({ graphBaseUrl, results, stderr: stderr.join(""), summary: { total: results.length, failCount } }, null, 2));
if (failCount > 0) process.exitCode = 1;
