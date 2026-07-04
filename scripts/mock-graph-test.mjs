#!/usr/bin/env node

import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { dirname, join, resolve } from "node:path";
import { homedir } from "node:os";
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

function text(res, status, body, headers = {}) {
  res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8", ...headers });
  res.end(body);
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

function folder(id, name, extra = {}) {
  return item(id, name, {
    folder: { childCount: 1 },
    file: undefined,
    ...extra
  });
}

const graph = createServer((req, res) => {
  const url = new URL(req.url, "http://127.0.0.1");
  const path = url.pathname;
  const decodedUrl = decodeURIComponent(req.url);
  requests.push({ method: req.method, path, url: req.url });

  if (req.method === "GET" && path === "/v1.0/me/drive") {
    if (count("drive") === 1) return json(res, 429, { error: { code: "tooManyRequests", message: "retry please" } }, { "Retry-After": "0" });
    return json(res, 200, { id: "drive", driveType: "personal", name: "Mock OneDrive" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delete-target") {
    return json(res, 200, item("delete-target", "delete-me.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/flaky-network") {
    if (count("flaky-network") === 1) {
      req.socket.destroy();
      return;
    }
    return json(res, 200, item("flaky-network", "flaky-network.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/text-error") {
    return text(res, 503, "temporary plain text failure");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/evil-pager/children") {
    return json(res, 200, {
      value: [item("before-evil-link", "before-evil-link.txt")],
      "@odata.nextLink": "https://evil.example.test/steal"
    });
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

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='")) {
    return json(res, 200, { value: [] });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/root/children") {
    return json(res, 200, {
      value: [
        folder("folder-a", "Folder A"),
        item("root-note", "root-note.txt")
      ],
      "@odata.nextLink": `http://127.0.0.1:${graph.address().port}/v1.0/mock/root-children-page-2`
    });
  }

  if (req.method === "GET" && path === "/v1.0/mock/root-children-page-2") {
    return json(res, 200, {
      value: [
        folder("folder-b", "Folder B")
      ]
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/folder-a/children") {
    return json(res, 200, {
      value: [
        item("deep-deck", "Deep Summary Deck.pptx", {
          parentReference: { path: "/drive/root:/Folder A" },
          file: { mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
        })
      ]
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/folder-b/children") {
    return json(res, 200, {
      value: [
        item("deep-pdf", "Nested Eval.pdf", {
          parentReference: { path: "/drive/root:/Folder B" },
          file: { mimeType: "application/pdf" }
        }),
        item("quarterly-report", "Quarterly Report.docx", {
          parentReference: { path: "/drive/root:/Folder B" },
          file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        })
      ]
    });
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

async function listTools() {
  const response = await request("tools/list", {});
  if (response.error) throw new Error(response.error.message);
  return response.result?.tools || [];
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

  await check("utility auth/config schemas are callable without item targets", async () => {
    const toolList = await listTools();
    const utilityNames = new Set([
      "onedrive_config",
      "onedrive_auth_device_start",
      "onedrive_auth_device_poll",
      "onedrive_logout"
    ]);
    const utilities = toolList.filter((entry) => utilityNames.has(entry.name));
    assert(utilities.length === utilityNames.size, "missing utility tools", { found: utilities.map((entry) => entry.name) });
    for (const entry of utilities) {
      assert(!entry.inputSchema?.anyOf, `${entry.name} should not require OneDrive item target fields`, entry.inputSchema);
    }
    return { checked: utilities.map((entry) => entry.name).sort() };
  });

  await check("GET requests retry 429 once", async () => {
    const result = await tool("onedrive_drive");
    assert(!result.isError, "onedrive_drive returned an error", result);
    assert(result.value.name === "Mock OneDrive", "drive response did not come from mock Graph", result.value);
    assert(counters.get("drive") === 2, "expected one retry after 429", { count: counters.get("drive") });
    return { attempts: counters.get("drive") };
  });

  await check("GET requests retry transient network failures", async () => {
    const result = await tool("onedrive_get_info", { itemId: "flaky-network" });
    assert(!result.isError, "transient network failure should be retried", result);
    assert(result.value.id === "flaky-network", "unexpected item after retry", result.value);
    assert(counters.get("flaky-network") === 2, "expected one retry after socket failure", { count: counters.get("flaky-network") });
    return { attempts: counters.get("flaky-network") };
  });

  await check("plain text Graph errors are surfaced", async () => {
    const result = await tool("onedrive_get_info", { itemId: "text-error" });
    assert(result.isError, "plain text Graph error should fail");
    assert(String(result.value).includes("temporary plain text failure"), "text error body was not included", result);
    return { message: String(result.value) };
  });

  await check("untrusted absolute nextLink is blocked before fetch", async () => {
    const before = requests.length;
    const result = await tool("onedrive_list_all", { itemId: "evil-pager", maxItems: 2 });
    assert(result.isError, "untrusted nextLink should fail");
    assert(String(result.value).includes("untrusted URL"), "unexpected nextLink error", result);
    const afterPaths = requests.slice(before).map((request) => request.url);
    assert(!afterPaths.some((url) => url.includes("evil.example.test")), "should not fetch untrusted nextLink", { afterPaths });
    return { graphRequestsAdded: requests.length - before };
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

  await check("remotePreset requires an explicit relative destination", async () => {
    const before = requests.length;
    const result = await tool("onedrive_write_text", { remotePreset: "documents", content: "hello" });
    assert(result.isError, "remotePreset without remoteRelativePath should fail");
    assert(String(result.value).includes("remoteRelativePath is required"), "unexpected remotePreset error", result);
    assert(requests.length === before, "remotePreset validation should not reach Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("unsafe create folder names are blocked before Graph", async () => {
    const before = requests.length;
    const result = await tool("onedrive_create_folder", { parentPath: "Folder A", name: "bad/name" });
    assert(result.isError, "path-like folder name should fail");
    assert(String(result.value).includes("single item name"), "unexpected folder name error", result);
    assert(requests.length === before, "unsafe folder name should not reach Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("unsafe rename names are blocked before Graph", async () => {
    const before = requests.length;
    const result = await tool("onedrive_rename", { itemId: "delete-target", newName: "bad/name" });
    assert(result.isError, "path-like rename should fail");
    assert(String(result.value).includes("single item name"), "unexpected rename error", result);
    assert(requests.length === before, "unsafe rename should not reach Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("unsafe restore destinations are blocked in dry-run", async () => {
    const before = requests.length;
    const result = await tool("onedrive_restore_deleted", { itemId: "deleted-item", destinationParentPath: "../Documents" });
    assert(result.isError, "unsafe restore destination should fail");
    assert(String(result.value).includes("unsafe path segment"), "unexpected restore destination error", result);
    assert(requests.length === before, "unsafe restore dry-run should not reach Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
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

  await check("download refuses local OneDrive sync destination by default", async () => {
    const result = await tool("onedrive_download", {
      itemId: "delete-target",
      localPath: join(homedir(), "Library", "CloudStorage", "OneDrive-Personal", "blocked-download.txt")
    });
    assert(result.isError, "download to local OneDrive sync path should fail");
    assert(String(result.value).includes("local OneDrive sync folder"), "unexpected sync-path guard message", result);
    return { blocked: true };
  });

  await check("upload refuses local OneDrive sync source before local or Graph work", async () => {
    const before = requests.length;
    const result = await tool("onedrive_upload", {
      localPath: join(homedir(), "Library", "CloudStorage", "OneDrive-Personal", "blocked-upload.txt"),
      remotePath: "Blocked Upload.txt"
    });
    assert(result.isError, "upload from local OneDrive sync path should fail");
    assert(String(result.value).includes("local OneDrive sync folder"), "unexpected sync-path guard message", result);
    assert(requests.length === before, "upload guard should not reach mock Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("recursive scan finds nested files beyond root", async () => {
    const result = await tool("onedrive_scan", {
      nameContains: "deck",
      extensions: ["pptx"],
      includeFolders: false,
      maxItems: 20,
      maxResults: 10
    });
    assert(!result.isError, "scan should succeed", result);
    assert(result.value.summary.itemsScanned === 6, "scan did not inspect all mock items", result.value.summary);
    assert(result.value.summary.foldersVisited === 3, "scan did not visit nested folders", result.value.summary);
    assert(result.value.summary.matched === 1, "scan should match one nested deck", result.value.summary);
    assert(result.value.items[0]?.id === "deep-deck", "scan did not return the nested deck", result.value.items);
    assert(result.value.items[0]?.remotePath === "Folder A/Deep Summary Deck.pptx", "scan did not include useful remotePath", result.value.items[0]);
    return {
      summary: result.value.summary,
      found: result.value.items[0]
    };
  });

  await check("stateless find falls back to remote scan for nested deck", async () => {
    const result = await tool("onedrive_find", {
      query: "Deep Summary Deck",
      maxResults: 3,
      scanMaxItems: 20,
      scanMaxFolders: 10
    });
    assert(!result.isError, "find should succeed", result);
    assert(result.value.summary.localIndexUsed === false, "find must not use a local index", result.value.summary);
    assert(result.value.summary.persistentCacheUsed === false, "find must not use persistent cache", result.value.summary);
    assert(result.value.summary.usedScanFallback === true, "find should use scan fallback when search misses", result.value.summary);
    assert(result.value.items[0]?.id === "deep-deck", "find did not return the nested deck first", result.value.items);
    assert(result.value.items[0]?.score >= 78, "find top score should be confident", result.value.items[0]);
    return {
      summary: result.value.summary,
      found: result.value.items[0],
      searchTerms: result.value.searchTerms
    };
  });

  await check("conversational find query still scans for a PowerPoint deck", async () => {
    const result = await tool("onedrive_find", {
      query: "where is my powerpoint deck",
      maxResults: 3,
      scanMaxItems: 20,
      scanMaxFolders: 10
    });
    assert(!result.isError, "conversational find should succeed", result);
    assert(result.value.summary.usedScanFallback === true, "conversational find should use scan fallback", result.value.summary);
    assert(result.value.items[0]?.id === "deep-deck", "conversational find did not return the nested deck first", result.value.items);
    return {
      summary: result.value.summary,
      found: result.value.items[0],
      searchTerms: result.value.searchTerms,
      inferred: result.value.inferred
    };
  });

  await check("report find does not force a PDF-only scan", async () => {
    const result = await tool("onedrive_find", {
      query: "Quarterly Report",
      maxResults: 3,
      scanMaxItems: 20,
      scanMaxFolders: 10
    });
    assert(!result.isError, "report find should succeed", result);
    assert(result.value.summary.usedScanFallback === true, "report find should use scan fallback when search misses", result.value.summary);
    assert(result.value.items[0]?.id === "quarterly-report", "report find should return the docx report", result.value.items);
    assert(!result.value.inferred?.strictExtensions?.includes(".pdf"), "report should not infer a strict PDF filter", result.value.inferred);
    return {
      summary: result.value.summary,
      found: result.value.items[0],
      inferred: result.value.inferred
    };
  });
} finally {
  child.stdin.end();
  child.kill("SIGTERM");
  graph.close();
}

const failCount = results.filter((result) => result.status === "fail").length;
console.log(JSON.stringify({ graphBaseUrl, results, stderr: stderr.join(""), summary: { total: results.length, failCount } }, null, 2));
if (failCount > 0) process.exitCode = 1;
