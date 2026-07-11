#!/usr/bin/env node

import { createServer } from "node:http";
import { execFileSync, spawn } from "node:child_process";
import { chmodSync, existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");
const mockHome = join(pluginRoot, "work", "mock-home");
const keepWork = process.argv.includes("--keep-work");
rmSync(mockHome, { recursive: true, force: true });
mkdirSync(mockHome, { recursive: true });
const officeFixtureDir = join(mockHome, "office-fixtures");
execFileSync("/usr/bin/python3", [join(pluginRoot, "scripts", "office-openxml-test.py"), `--emit-fixtures=${officeFixtureDir}`], {
  env: { ...process.env, PYTHONPYCACHEPREFIX: join(mockHome, "pycache") },
  stdio: "ignore"
});
let officeWordBuffer = readFileSync(join(officeFixtureDir, "sample.docx"));
let officeExcelBuffer = readFileSync(join(officeFixtureDir, "sample.xlsx"));
let officeBusinessBuffer = Buffer.from(officeExcelBuffer);
let officePowerPointBuffer = readFileSync(join(officeFixtureDir, "sample.pptx"));
let officeWordPostMetadataFailuresRemaining = 0;
let failNextOfficeExcelPut = false;
const officeVersions = { "office-word": 1, "office-excel": 1, "office-business": 1, "office-powerpoint": 1 };

const requests = [];
const authRequests = [];
const graphBodies = [];
const counters = new Map();
let refreshResponseDelayMs = 0;
let deviceStartResponseDelayMs = 0;
let devicePollResponseDelayMs = 0;
let devicePollShouldSucceed = false;
let rootDeltaScenario = "empty";

function count(key) {
  const next = (counters.get(key) || 0) + 1;
  counters.set(key, next);
  return next;
}

async function waitForCounter(key, minimum, timeoutMs = 2000) {
  const deadline = Date.now() + timeoutMs;
  while ((counters.get(key) || 0) < minimum) {
    if (Date.now() >= deadline) throw new Error(`Timed out waiting for counter ${key} to reach ${minimum}.`);
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

function createMcpClient(env) {
  const processChild = spawn(process.execPath, [serverPath], {
    cwd: pluginRoot,
    env,
    stdio: ["pipe", "pipe", "pipe"]
  });
  let clientNextId = 1;
  let clientBuffer = "";
  const clientPending = new Map();
  processChild.stdout.on("data", (chunk) => {
    clientBuffer += chunk.toString("utf8");
    for (;;) {
      const newline = clientBuffer.indexOf("\n");
      if (newline < 0) break;
      const line = clientBuffer.slice(0, newline).trim();
      clientBuffer = clientBuffer.slice(newline + 1);
      if (!line) continue;
      const message = JSON.parse(line);
      const waiter = clientPending.get(message.id);
      if (!waiter) continue;
      clientPending.delete(message.id);
      clearTimeout(waiter.timeout);
      waiter.resolve(message);
    }
  });
  const request = (method, params = {}) => {
    const id = clientNextId++;
    const promise = new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        clientPending.delete(id);
        reject(new Error(`Timed out waiting for isolated ${method}`));
      }, 5000);
      clientPending.set(id, { resolve, reject, timeout });
    });
    processChild.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
    return promise;
  };
  const callTool = async (name, args = {}) => {
    const response = await request("tools/call", { name, arguments: args });
    const text = response.result?.content?.[0]?.text ?? "";
    let value = text;
    try {
      value = JSON.parse(text);
    } catch {
      // Keep plain text.
    }
    return { isError: Boolean(response.result?.isError), value };
  };
  const close = async () => {
    processChild.stdin.end();
    processChild.kill("SIGTERM");
    await new Promise((resolve) => {
      const timeout = setTimeout(resolve, 1000);
      processChild.once("exit", () => {
        clearTimeout(timeout);
        resolve();
      });
    });
  };
  return { request, tool: callTool, close };
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

function binary(res, status, body, headers = {}) {
  res.writeHead(status, headers);
  res.end(body);
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => {
      const raw = Buffer.concat(chunks).toString("utf8");
      if (!raw) return resolve({});
      try {
        resolve(JSON.parse(raw));
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function readBufferBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function item(id, name, extra = {}) {
  return {
    id,
    name,
    webUrl: `https://example.test/${encodeURIComponent(id)}`,
    size: 12,
    createdDateTime: "2026-07-04T00:00:00Z",
    lastModifiedDateTime: "2026-07-04T00:00:00Z",
    eTag: `etag-${id}`,
    cTag: `ctag-${id}`,
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

function officeItem(id, name, buffer, mimeType, extra = {}) {
  return item(id, name, {
    size: buffer.length,
    eTag: `etag-${id}-${officeVersions[id]}`,
    cTag: `ctag-${id}-${officeVersions[id]}`,
    file: { mimeType },
    ...extra
  });
}

function permissionsForItem(id) {
  if (id === "root-note") {
    return [{ id: "perm-owner-root-note", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } }];
  }
  if (id === "deep-deck") {
    return [
      { id: "perm-owner-deep-deck", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } },
      { id: "perm-public-deep-deck", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/public/deep-deck" } }
    ];
  }
  if (id === "deep-pdf") {
    return [{
      id: "perm-v2-direct-deep-pdf",
      roles: ["read"],
      grantedToIdentitiesV2: [{ user: { id: "v2-user", displayName: "V2 User", email: "v2@example.test" } }]
    }];
  }
  return [];
}

const graph = createServer(async (req, res) => {
  const url = new URL(req.url, "http://127.0.0.1");
  const path = url.pathname;
  const decodedUrl = decodeURIComponent(req.url);
  requests.push({ method: req.method, path, url: req.url, headers: req.headers });

  if (req.method === "GET" && path === "/v1.0/me") {
    return json(res, 200, { displayName: "Mock User", userPrincipalName: "mock@example.test", mail: "mock@example.test" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive") {
    if (count("drive") === 1) return json(res, 429, { error: { code: "tooManyRequests", message: "retry please" } }, { "Retry-After": "0" });
    return json(res, 200, { id: "drive", driveType: "personal", name: "Mock OneDrive" });
  }
  if (req.method === "GET" && path === "/v1.0/drives/business-drive") {
    return json(res, 200, { id: "business-drive", driveType: "business", name: "Mock Business Drive" });
  }

  if (req.method === "POST" && path === "/v1.0/$batch") {
    const body = await readJsonBody(req);
    graphBodies.push(body);
    if (count("batch-outer") === 1) {
      return json(res, 503, { error: { code: "serviceUnavailable", message: "retry outer batch" } }, { "Retry-After": "0" });
    }
    const responses = (body.requests || []).map((request) => {
      const requestUrl = `/${String(request.url || "").replace(/^\/+/, "")}`;
      const idMatch = requestUrl.match(/\/items\/([^/?]+)(?:\/permissions)?(?:\?|$)/);
      const itemId = idMatch ? decodeURIComponent(idMatch[1]) : null;
      if (requestUrl.includes("/permissions")) {
        if (itemId === "permission-error") {
          count("batch-permission-error");
          return {
            id: request.id,
            status: 503,
            headers: { "Retry-After": "0" },
            body: { error: { code: "serviceUnavailable", message: "mock permission audit failure" } }
          };
        }
        return { id: request.id, status: 200, body: { value: permissionsForItem(itemId) } };
      }
      if (itemId === "batch-flaky") {
        const attempt = count("batch-flaky");
        if (attempt === 1) {
          return {
            id: request.id,
            status: 429,
            headers: { "Retry-After": "0" },
            body: { error: { code: "tooManyRequests", message: "retry this subrequest" } }
          };
        }
        return { id: request.id, status: 200, body: item("batch-flaky", "batch-flaky.txt") };
      }
      if (itemId === "root-note") count("batch-steady");
      return { id: request.id, status: 200, body: item(itemId || "batch-item", itemId === "root-note" ? "root-note.txt" : `${itemId || "batch-item"}.txt`) };
    });
    return json(res, 200, { responses });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delete-target") {
    return json(res, 200, item("delete-target", "delete-me.txt"));
  }

  if (req.method === "PATCH" && path === "/v1.0/me/drive/items/delete-target") {
    count("rename");
    return json(res, 200, item("delete-target", "renamed-cache.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/root-note") {
    return json(res, 200, item("root-note", "root-note.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root:/root-note.txt:") {
    return json(res, 200, item("root-note", "root-note.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/root-note/content") {
    return text(res, 200, "root note mock content\n");
  }

  const officeDefinitions = {
    "office-word": { name: "sample.docx", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", get buffer() { return officeWordBuffer; }, set buffer(value) { officeWordBuffer = value; } },
    "office-excel": { name: "sample.xlsx", mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", get buffer() { return officeExcelBuffer; }, set buffer(value) { officeExcelBuffer = value; } },
    "office-business": { name: "business.xlsx", mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", get buffer() { return officeBusinessBuffer; }, set buffer(value) { officeBusinessBuffer = value; }, driveId: "business-drive" },
    "office-powerpoint": { name: "sample.pptx", mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation", get buffer() { return officePowerPointBuffer; }, set buffer(value) { officePowerPointBuffer = value; } }
  };
  const officeItemMatch = path.match(/^\/v1\.0\/me\/drive\/items\/(office-(?:word|excel|business|powerpoint))(?:\/content)?$/);
  if (officeItemMatch) {
    const id = officeItemMatch[1];
    const definition = officeDefinitions[id];
    if (req.method === "GET" && path.endsWith("/content")) return binary(res, 200, definition.buffer, { "Content-Type": "application/octet-stream" });
    if (req.method === "PUT" && path.endsWith("/content")) {
      if (id === "office-excel" && failNextOfficeExcelPut) {
        failNextOfficeExcelPut = false;
        await readBufferBody(req);
        return json(res, 500, { error: { code: "internalError", message: "mock cross-file partial failure" } });
      }
      definition.buffer = await readBufferBody(req);
      officeVersions[id] += 1;
      if (id === "office-word") officeWordPostMetadataFailuresRemaining = 5;
      return json(res, 200, officeItem(id, definition.name, definition.buffer, definition.mime, definition.driveId ? { parentReference: { driveId: definition.driveId, path: `/drives/${definition.driveId}/root:` } } : {}));
    }
    if (req.method === "GET" && id === "office-word" && officeWordPostMetadataFailuresRemaining > 0) {
      officeWordPostMetadataFailuresRemaining -= 1;
      return json(res, 503, { error: { code: "serviceUnavailable", message: "mock post-commit metadata failure" } }, { "Retry-After": "0" });
    }
    if (req.method === "GET") return json(res, 200, officeItem(id, definition.name, definition.buffer, definition.mime, definition.driveId ? { parentReference: { driveId: definition.driveId, path: `/drives/${definition.driveId}/root:` } } : {}));
  }
  if (req.method === "POST" && path === "/v1.0/drives/business-drive/items/office-business/workbook/createSession") {
    const attempt = count("excel-create-session");
    if (attempt === 1) {
      return json(res, 504, { error: { code: "gatewayTimeout", message: "mock documented createSession timeout" } }, { "Retry-After": "0" });
    }
    if (attempt === 3) {
      return json(res, 202, {}, {
        Location: `http://${req.headers.host}/v1.0/drives/business-drive/items/office-business/workbook/operations/untrusted-resource`
      });
    }
    if (attempt === 4) return json(res, 202, {}, { Location: "https://attacker.invalid/workbook/operations/session" });
    return json(res, 202, {}, {
      Location: `http://${req.headers.host}/v1.0/drives/business-drive/items/office-business/workbook/operations/create-session`
    });
  }
  if (req.method === "GET" && path === "/v1.0/drives/business-drive/items/office-business/workbook/operations/untrusted-resource") {
    return json(res, 200, { id: "untrusted-resource", status: "succeeded", resourceLocation: "https://attacker.invalid/sessionInfoResource" });
  }
  if (req.method === "GET" && path === "/v1.0/drives/business-drive/items/office-business/workbook/operations/create-session") {
    if (count("excel-create-session-poll") === 1) {
      return json(res, 200, { id: "create-session", status: "running" }, { "Retry-After": "0" });
    }
    return json(res, 200, {
      id: "create-session",
      status: "succeeded",
      resourceLocation: `http://${req.headers.host}/v1.0/drives/business-drive/items/office-business/workbook/sessionInfoResource(key='create-session')`
    });
  }
  if (req.method === "GET" && path === "/v1.0/drives/business-drive/items/office-business/workbook/sessionInfoResource(key='create-session')") {
    return json(res, 200, { id: "business-session", persistChanges: true });
  }
  if (["PATCH", "POST"].includes(req.method) && path.startsWith("/v1.0/drives/business-drive/items/office-business/workbook/worksheets/")) {
    await readBufferBody(req);
    return json(res, 200, { address: "Data!A1", values: [["Graph updated"]] });
  }
  if (req.method === "POST" && path.startsWith("/v1.0/drives/business-drive/items/office-business/workbook/tables/")) {
    await readBufferBody(req);
    return json(res, 201, { index: 1, values: [["Q4", 42]] });
  }
  if (req.method === "POST" && path === "/v1.0/drives/business-drive/items/office-business/workbook/closeSession") {
    return json(res, 503, { error: { code: "serviceUnavailable", message: "mock close warning" } });
  }
  const officeUploadMatch = path.match(/^\/v1\.0\/me\/drive\/root:\/(sample\.(docx|xlsx|pptx)):\/content$/);
  if (req.method === "PUT" && officeUploadMatch) {
    const id = officeUploadMatch[2] === "docx" ? "office-word" : officeUploadMatch[2] === "xlsx" ? "office-excel" : "office-powerpoint";
    const definition = officeDefinitions[id];
    definition.buffer = await readBufferBody(req);
    officeVersions[id] += 1;
    return json(res, 200, officeItem(id, definition.name, definition.buffer, definition.mime));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/tibetan-note") {
    return json(res, 200, item("tibetan-note", "tibetan-note.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/tibetan-note/content") {
    return text(res, 200, "Tibetan language reference\n");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/beta-note") {
    return json(res, 200, item("beta-note", "beta-note.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/beta-note/content") {
    return text(res, 200, "beta launch notes\n");
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/beta-note") {
    count("delete-beta-note");
    return empty(res, 204);
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-parent") {
    return json(res, 200, folder("delta-parent", "Documents", { parentReference: { id: "root", path: "/drive/root:" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-archive") {
    return json(res, 200, folder("delta-archive", "Archive", { parentReference: { id: "root", path: "/drive/root:" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-child") {
    return json(res, 200, item("delta-child", "report.txt", { parentReference: { id: "delta-parent", path: "/drive/root:/Documents" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-index-parent") {
    return json(res, 200, folder("delta-index-parent", "Indexed", { parentReference: { id: "root", path: "/drive/root:" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-indexed") {
    return json(res, 200, item("delta-indexed", "indexed-report.txt", {
      eTag: "index-e1",
      cTag: "index-c1",
      parentReference: { id: "delta-index-parent", path: "/drive/root:/Indexed" },
      file: { mimeType: "text/plain", hashes: { quickXorHash: "OLDHASH" } }
    }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-indexed/content") {
    return text(res, 200, "persistent indexed phrase\n");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-index-parent/delta") {
    const run = count("delta-index-run");
    return json(res, 200, {
      value: [item("delta-indexed", run === 1 ? "renamed-indexed-report.txt" : "changed-indexed-report.txt", {
        eTag: run === 1 ? "index-e2" : "index-e3",
        cTag: run === 1 ? "index-c1" : undefined,
        parentReference: { id: "delta-index-parent" },
        file: { mimeType: "text/plain" }
      })],
      "@odata.deltaLink": `http://127.0.0.1:${graph.address().port}/v1.0/me/drive/items/delta-index-parent/delta?token=${run}`
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-new-parent") {
    return json(res, 200, folder("delta-new-parent", "New Parent", { parentReference: { id: "root", path: "/drive/root:" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root/delta") {
    const value = rootDeltaScenario === "pathless"
      ? [
          item("top-pathless", "top-pathless.txt", { parentReference: { id: "root", driveId: "drive" } }),
          item("moved-pathless", "moved-pathless.txt", { parentReference: { id: "delta-new-parent", driveId: "drive" } }),
          item("unresolved-pathless", "unresolved-pathless.txt", { parentReference: { id: "missing-parent", driveId: "drive" } })
        ]
      : [];
    return json(res, 200, {
      value,
      "@odata.deltaLink": `http://127.0.0.1:${graph.address().port}/v1.0/me/drive/root/delta?token=${rootDeltaScenario}`
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/concurrent-a") {
    return json(res, 200, item("concurrent-a", "concurrent-a.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/concurrent-b") {
    return json(res, 200, item("concurrent-b", "concurrent-b.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/concurrent-a/content") return text(res, 200, "concurrent alpha phrase\n");
  if (req.method === "GET" && path === "/v1.0/me/drive/items/concurrent-b/content") return text(res, 200, "concurrent beta phrase\n");

  if (req.method === "GET" && path === "/v1.0/me/drive/items/victim-folder") {
    return json(res, 200, folder("victim-folder", "Collision", { parentReference: { id: "root", path: "/drive/root:" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/victim-child") {
    return json(res, 200, item("victim-child", "victim-child.txt", { parentReference: { id: "victim-folder", path: "/drive/root:/Collision" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/victim-child/content") return text(res, 200, "victim indexed phrase\n");

  if (req.method === "GET" && path === "/v1.0/me/drive/items/replacement-collision") {
    return json(res, 200, item("replacement-collision", "Collision", { parentReference: { id: "root", path: "/drive/root:" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delta-parent/delta") {
    const run = count("delta-path-run");
    const name = run >= 2 ? "renamed-report.txt" : "report.txt";
    const parentId = run >= 3 ? "delta-archive" : "delta-parent";
    return json(res, 200, {
      value: [item("delta-child", name, { parentReference: { id: parentId } })],
      "@odata.deltaLink": `http://127.0.0.1:${graph.address().port}/v1.0/me/drive/items/delta-parent/delta?token=${run}`
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/dup-a") {
    return json(res, 200, item("dup-a", "same-name.txt", { parentReference: { path: "/drive/root:/Folder A" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/dup-b") {
    return json(res, 200, item("dup-b", "same-name.txt", { parentReference: { path: "/drive/root:/Folder B" } }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/dup-a/content") {
    return text(res, 200, "duplicate a\n");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/dup-b/content") {
    return text(res, 200, "duplicate b\n");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root:/root-note.txt:/content") {
    return text(res, 200, "root note mock content\n");
  }

  if (req.method === "PUT" && path === "/v1.0/me/drive/root:/root-note.txt:/content") {
    count("root-note-upload");
    return json(res, 200, item("root-note", "root-note.txt"));
  }

  if (req.method === "PUT" && path === "/v1.0/me/drive/root:/empty-session.txt:/content") {
    count("empty-session-simple-upload");
    return json(res, 200, item("empty-session", "empty-session.txt", { size: 0 }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/big-text") {
    return json(res, 200, item("big-text", "big-text.txt", { size: 1 }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/big-text/content") {
    return text(res, 200, "0123456789abcdefghijklmnopqrstuvwxyz");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/six-meg-text") {
    return json(res, 200, item("six-meg-text", "six-meg-text.txt", { size: 6 * 1024 * 1024 }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/six-meg-text/content") {
    return text(res, 200, "bounded six meg metadata test");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/sharing-error-root") {
    return json(res, 200, folder("sharing-error-root", "Sharing Error Root"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/sharing-error-root/children") {
    return json(res, 200, { value: [item("permission-error", "permission-error.txt")] });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/delete-fail") {
    return json(res, 200, item("delete-fail", "delete-fail.txt"));
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/delete-fail") {
    count("delete-fail");
    return json(res, 503, {
      error: {
        code: "serviceUnavailable",
        message: "mock delete failure",
        innerError: { "request-id": "mock-delete-fail-request" }
      }
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/move-fail") {
    return json(res, 200, item("move-fail", "move-fail.txt"));
  }

  if (req.method === "PATCH" && path === "/v1.0/me/drive/items/move-fail") {
    count("move-fail");
    return json(res, 503, {
      error: {
        code: "serviceUnavailable",
        message: "mock move failure",
        innerError: { "request-id": "mock-move-fail-request" }
      }
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/root-note") {
    return json(res, 200, item("root-note", "root-note.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/flaky-network") {
    if (count("flaky-network") === 1) {
      req.socket.destroy();
      return;
    }
    return json(res, 200, item("flaky-network", "flaky-network.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/stale-cache") {
    return json(res, 200, item("stale-cache", "Stale Cache Deck.pptx", {
      parentReference: { path: "/drive/root:/Missing" },
      file: { mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
    }));
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

  if (req.method === "GET" && path === "/v1.0/me/drive/items/cycle-pager/children") {
    return json(res, 200, {
      value: [item("cycle-link", "cycle-link.txt")],
      "@odata.nextLink": `http://127.0.0.1:${graph.address().port}${req.url}`
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/empty-pager") {
    return json(res, 200, item("empty-pager", "empty-pager", { folder: { childCount: 0 }, file: undefined }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/empty-pager/children") {
    const page = Number(url.searchParams.get("page") || "0");
    return json(res, 200, {
      value: [],
      "@odata.nextLink": `http://127.0.0.1:${graph.address().port}/v1.0/me/drive/items/empty-pager/children?page=${page + 1}`
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/oversized-pager/children") {
    return json(res, 200, {
      value: [item("oversized-a", "oversized-a.txt"), item("oversized-b", "oversized-b.txt")],
      "@odata.nextLink": `http://127.0.0.1:${graph.address().port}/v1.0/mock/oversized-next`
    });
  }

  if (req.method === "GET" && path === "/v1.0/mock/oversized-next") {
    return json(res, 200, { value: [item("oversized-c", "oversized-c.txt")] });
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/delete-target") {
    count("delete");
    return empty(res, 204, { "request-id": "mock-delete-request" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/copy-src") {
    return json(res, 200, item("copy-src", "copy-source.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/copy-evil") {
    return json(res, 200, item("copy-evil", "copy-evil.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/copy-http-sharepoint") {
    return json(res, 200, item("copy-http-sharepoint", "copy-http-sharepoint.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/postverify-fail") {
    return json(res, 200, item("postverify-fail", "postverify-fail.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/postverify-fail/permissions") {
    if (counters.get("postverify-create-link")) {
      count("postverify-permissions-fail");
      return json(res, 503, { error: { code: "serviceUnavailable", message: "mock post-mutation verification failure" } }, { "Retry-After": "0" });
    }
    return json(res, 200, { value: [{ id: "postverify-owner", roles: ["owner"], grantedTo: { user: { displayName: "Mock User" } } }] });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/copy-src/permissions") {
    const base = [{
      id: "perm-owner",
      roles: ["owner"],
      grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } },
      grantedToIdentitiesV2: [{ user: { id: "user-owner", displayName: "Mock User", email: "mock@example.test" } }]
    }];
    if (counters.get("create-link")) {
      base.push({ id: "perm-link", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/share/link" } });
    }
    if (counters.get("invite-permission")) {
      base.push({
        id: "perm-invite",
        roles: ["read"],
        grantedToIdentitiesV2: [{ user: { id: "user-invite", displayName: "Invited User", email: "person@example.test" } }],
        invitation: { email: "person@example.test", signInRequired: true }
      });
    }
    return json(res, 200, { value: base });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/root-note/permissions") {
    return json(res, 200, {
      value: [{ id: "perm-owner-root-note", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } }]
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/deep-deck/permissions") {
    return json(res, 200, {
      value: [
        { id: "perm-owner-deep-deck", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } },
        { id: "perm-public-deep-deck", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/public/deep-deck" } }
      ]
    });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/copy-src/createLink") {
    count("create-link");
    graphBodies.push({ key: "create-link", body: await readJsonBody(req) });
    return json(res, 200, { id: "perm-link", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/share/link" } }, { "request-id": "mock-create-link-request" });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/postverify-fail/createLink") {
    count("postverify-create-link");
    return json(res, 200, { id: "postverify-link", roles: ["read"], link: { type: "view", scope: "anonymous" } });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/copy-src/invite") {
    count("invite-permission");
    const body = await readJsonBody(req);
    graphBodies.push({ key: "invite-permission", body });
    return json(res, 200, {
      value: [{
        id: "perm-invite",
        roles: body.roles || ["read"],
        grantedToIdentitiesV2: [{ user: { id: "user-invite", displayName: "Invited User", email: body.recipients?.[0]?.email } }],
        invitation: { email: body.recipients?.[0]?.email, signInRequired: body.requireSignIn }
      }]
    }, { "request-id": "mock-invite-request" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/invite-fail") {
    return json(res, 200, item("invite-fail", "invite-fail.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/invite-fail/permissions") {
    return json(res, 200, {
      value: [{ id: "perm-owner-invite-fail", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } }]
    });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/invite-fail/invite") {
    count("invite-fail");
    graphBodies.push({ key: "invite-fail", body: await readJsonBody(req) });
    return json(res, 403, {
      error: {
        code: "accessDenied",
        message: "mock invite failure for person@example.test object 11111111-1111-1111-1111-111111111111 at https://example.test/invite?token=secret Bearer abc.def.ghi",
        innerError: { "request-id": "mock-invite-fail-request" }
      }
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-target") {
    return json(res, 200, item("revoke-target", "shared-doc.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-target/permissions") {
    const permissions = [
      { id: "perm-owner", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } }
    ];
    if (!counters.get("revoke-perm-public")) {
      permissions.push({ id: "perm-public", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/revoke/public" } });
    }
    return json(res, 200, { value: permissions });
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/revoke-target/permissions/perm-public") {
    count("revoke-perm-public");
    return empty(res, 204, { "request-id": "mock-revoke-request" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-a") {
    return json(res, 200, item("revoke-a", "revoke-a.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-a/permissions") {
    return json(res, 200, {
      value: [{ id: "perm-a", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/revoke/a" } }]
    });
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/revoke-a/permissions/perm-a") {
    count("revoke-perm-a");
    return empty(res, 204, { "request-id": "mock-revoke-a-request" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-b") {
    return json(res, 200, item("revoke-b", "revoke-b.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-b/permissions") {
    return json(res, 200, {
      value: [{ id: "perm-owner-b", roles: ["owner"], grantedTo: { user: { displayName: "Mock User", email: "mock@example.test" } } }]
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-fail") {
    return json(res, 200, item("revoke-fail", "revoke-fail.txt"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/revoke-fail/permissions") {
    return json(res, 200, {
      value: [{ id: "perm-fail", roles: ["read"], link: { type: "view", scope: "anonymous", webUrl: "https://example.test/revoke/fail" } }]
    });
  }

  if (req.method === "DELETE" && path === "/v1.0/me/drive/items/revoke-fail/permissions/perm-fail") {
    count("revoke-perm-fail");
    return json(res, 503, {
      error: {
        code: "serviceUnavailable",
        message: "mock revoke failure",
        innerError: { "request-id": "mock-revoke-fail-request" }
      }
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/quarterly-report") {
    return json(res, 200, item("quarterly-report", "Quarterly Report.docx", {
      parentReference: { path: "/drive/root:/Folder B" },
      file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
    }));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/quarterly-report/content") {
    if (url.searchParams.get("format") === "pdf") {
      return binary(res, 200, Buffer.from("%PDF-1.7 mock export\n"), { "Content-Type": "application/pdf" });
    }
    if (url.searchParams.get("format") === "text") {
      return text(res, 200, "Quarterly Report mock text export\n");
    }
    return text(res, 200, "Quarterly Report raw content\n");
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root") {
    return json(res, 200, { id: "root", name: "root", root: {}, folder: {} });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/root:/trusted-session.txt:/createUploadSession") {
    count("trusted-upload-session");
    return json(res, 200, { uploadUrl: `http://127.0.0.1:${graph.address().port}/v1.0/mock/upload-session/trusted` });
  }

  if (req.method === "PUT" && path === "/v1.0/mock/upload-session/trusted") {
    count("trusted-upload-session-put");
    return json(res, 201, item("trusted-session", "trusted-session.txt"));
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/root:/evil-session.txt:/createUploadSession") {
    count("evil-upload-session");
    return json(res, 200, { uploadUrl: "https://evil.example.test/upload/session" });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root/children") {
    return json(res, 200, {
      value: [
        folder("folder-a", "Folder A"),
        item("root-note", "root-note.txt")
      ]
    });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/root/children") {
    const body = await readJsonBody(req);
    return json(res, 201, folder("created-child-folder", body.name || "Created Child", {
      parentReference: { path: "/drive/root:" }
    }));
  }

  if (req.method === "POST" && decodedUrl.includes("/v1.0/me/drive/root:/Folder A:/children")) {
    const body = await readJsonBody(req);
    return json(res, 201, folder("created-child-folder", body.name || "Created Child", {
      parentReference: { path: "/drives/drive/root:/Folder A" }
    }));
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='Adaptive Exact Match')")) {
    return json(res, 200, {
      value: [item("adaptive-exact-hit", "Adaptive Exact Match.txt", {
        parentReference: { path: "/drive/root:/Research" }
      })]
    });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='Paged Research')")) {
    return json(res, 200, {
      value: [item("paged-research-a", "Paged Research A.txt")],
      "@odata.nextLink": `http://127.0.0.1:${graph.address().port}/v1.0/mock/paged-research-2`
    });
  }

  if (req.method === "GET" && path === "/v1.0/mock/paged-research-2") {
    return json(res, 200, { value: [item("paged-research-b", "Paged Research B.txt")] });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='Quarterly Research')")) {
    return json(res, 200, { value: [item("quarterly-research-exact", "Quarterly Research.txt")] });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='quarterly')")) {
    return json(res, 200, { value: [item("quarterly-research-broad", "Quarterly Field Research.txt")] });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='No Cache Live')")) {
    return json(res, 200, { value: [item("no-cache-live", "No Cache Live.txt")] });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='Batch Cache Research')")) {
    return json(res, 200, { value: [item("batch-cache-canonical", "Misc Notes.txt")] });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='batch cache')")) {
    return json(res, 200, { value: [item("batch-cache-expansion", "Batch Cache Notes.txt")] });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='canonical hidden phrase')")) {
    return json(res, 200, {
      value: [item("canonical-content-hit", "Minutes.docx", {
        parentReference: { path: "/drive/root:/Meetings" },
        file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
      })]
    });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='canonical hidden')")) {
    return json(res, 200, {
      value: [item("unrelated-expansion-hit", "Random Notes.txt", {
        parentReference: { path: "/drive/root:/Archive" }
      })]
    });
  }

  if (req.method === "GET" && decodedUrl.includes("/v1.0/me/drive/root/search(q='")) {
    return json(res, 200, { value: [] });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root:/Folder%20A:") {
    return json(res, 200, folder("folder-a", "Folder A", { parentReference: { id: "root", path: "/drive/root:" } }));
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

  if (req.method === "GET" && path === "/v1.0/me/drive/items/folder-a") {
    return json(res, 200, folder("folder-a", "Folder A"));
  }

  if (req.method === "PATCH" && path === "/v1.0/me/drive/items/folder-a") {
    return json(res, 200, folder("folder-a", "Folder Renamed"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/folder-a/delta") {
    return json(res, 200, {
      value: [item("deep-deck", "Deep Summary Deck.pptx", {
        parentReference: { path: "/drive/root:/Folder A" },
        file: { mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
      })],
      "@odata.deltaLink": `http://127.0.0.1:${graph.address().port}/v1.0/mock/delta/folder-a`
    });
  }

  if (req.method === "PATCH" && path === "/v1.0/me/drive/items/root-note") {
    count("move-root-note");
    return json(res, 200, item("root-note", "root-note.txt", { parentReference: { path: "/drive/root:/Folder A" } }));
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

  if (req.method === "GET" && path === "/v1.0/me/drive/items/folder-b") {
    return json(res, 200, folder("folder-b", "Folder B"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/bulk-large") {
    return json(res, 200, folder("bulk-large", "Bulk Large"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/bulk-large/children") {
    return json(res, 200, {
      value: [
        ...Array.from({ length: 5000 }, (_, index) => item(`small-${index}`, `small-${index}.txt`, { size: 1, parentReference: { path: "/drive/root:/Bulk Large" } })),
        item("huge-after-cap", "huge-after-cap.bin", { size: 999999999, parentReference: { path: "/drive/root:/Bulk Large" } })
      ]
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/bulk-dupes") {
    return json(res, 200, folder("bulk-dupes", "Bulk Dupes"));
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/bulk-dupes/children") {
    return json(res, 200, {
      value: [
        ...Array.from({ length: 5000 }, (_, index) => item(`unique-${index}`, `unique-${index}.txt`, { size: index + 1, parentReference: { path: "/drive/root:/Bulk Dupes" } })),
        item("dup-a", "same-name.txt", { size: 42, parentReference: { path: "/drive/root:/Bulk Dupes" } }),
        item("dup-b", "same-name.txt", { size: 42, parentReference: { path: "/drive/root:/Bulk Dupes" } })
      ]
    });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/items/folder-b/delta") {
    return json(res, 200, {
      value: [item("quarterly-report", "Quarterly Report.docx", {
        parentReference: { path: "/drive/root:/Folder B" },
        file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
      })],
      "@odata.deltaLink": `http://127.0.0.1:${graph.address().port}/v1.0/mock/delta/folder-b`
    });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/copy-src/copy") {
    count("copy");
    return empty(res, 202, { Location: `http://127.0.0.1:${graph.address().port}/monitor/copy?token=copy-secret`, "request-id": "mock-copy-request" });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/copy-evil/copy") {
    return empty(res, 202, { Location: "https://evil.example.test/monitor/copy" });
  }

  if (req.method === "POST" && path === "/v1.0/me/drive/items/copy-http-sharepoint/copy") {
    return empty(res, 202, { Location: "http://tenant.sharepoint.com/monitor/copy?token=copy-secret" });
  }

  if (req.method === "GET" && path === "/monitor/copy") {
    return empty(res, 303, { Location: `http://127.0.0.1:${graph.address().port}/v1.0/me/drive/items/copied?downloadToken=secret` });
  }

  if (req.method === "GET" && path === "/v1.0/me/drive/root:/Documents/Pictures:") {
    return json(res, 200, item("pictures", "Pictures", { folder: { childCount: 0 }, file: undefined }));
  }

  json(res, 404, { error: { code: "notFound", message: `${req.method} ${req.url}` } });
});

await new Promise((resolve) => graph.listen(0, "127.0.0.1", resolve));
const graphBaseUrl = `http://127.0.0.1:${graph.address().port}/v1.0`;

const identity = createServer(async (req, res) => {
  const url = new URL(req.url, "http://127.0.0.1");
  authRequests.push({ method: req.method, path: url.pathname });

  if (req.method === "POST" && url.pathname === "/common/oauth2/v2.0/devicecode") {
    return json(res, 400, {
      error: "invalid_request",
      error_description: "AADSTS9002331: Application is configured for use by Microsoft Account users only. Please use the /consumers endpoint to serve this request."
    });
  }

  if (req.method === "POST" && url.pathname === "/consumers/oauth2/v2.0/devicecode") {
    const attempt = count("identity-device-start");
    const delayMs = deviceStartResponseDelayMs;
    if (delayMs) await new Promise((resolve) => setTimeout(resolve, delayMs));
    return json(res, 200, {
      user_code: `MOCK-CODE-${attempt}`,
      device_code: `mock-device-code-${attempt}`,
      verification_uri: "https://microsoft.com/devicelogin",
      verification_uri_complete: `https://microsoft.com/devicelogin?user_code=MOCK-CODE-${attempt}`,
      expires_in: 900,
      interval: 5,
      message: `Use code MOCK-CODE-${attempt}`
    });
  }

  if (req.method === "POST" && url.pathname.endsWith("/oauth2/v2.0/token")) {
    const form = new URLSearchParams(await readRawBody(req));
    const grantType = form.get("grant_type");
    if (grantType === "refresh_token") {
      const attempt = count("identity-refresh");
      if (refreshResponseDelayMs) await new Promise((resolve) => setTimeout(resolve, refreshResponseDelayMs));
      return json(res, 200, {
        token_type: "Bearer",
        access_token: `refreshed-access-${attempt}`,
        refresh_token: `refreshed-refresh-${attempt}`,
        expires_in: 3600
      });
    }
    if (grantType === "urn:ietf:params:oauth:grant-type:device_code") {
      const attempt = count("identity-device-poll");
      if (devicePollResponseDelayMs) await new Promise((resolve) => setTimeout(resolve, devicePollResponseDelayMs));
      if (devicePollShouldSucceed) {
        return json(res, 200, {
          token_type: "Bearer",
          access_token: `device-access-${attempt}`,
          refresh_token: `device-refresh-${attempt}`,
          expires_in: 3600
        });
      }
      return json(res, 400, {
        error: "authorization_pending",
        error_description: "Authorization is still pending."
      });
    }
  }

  json(res, 404, { error: "not_found", error_description: `${req.method} ${url.pathname}` });
});

await new Promise((resolve) => identity.listen(0, "127.0.0.1", resolve));
const identityBaseUrl = `http://127.0.0.1:${identity.address().port}`;

const child = spawn(process.execPath, [serverPath], {
  cwd: pluginRoot,
  env: {
    ...process.env,
    HOME: mockHome,
    ONEDRIVE_CLIENT_ID: "mock-client-id",
    ONEDRIVE_TEST_ACCESS_TOKEN: "mock-token",
    ONEDRIVE_GRAPH_BASE_URL: graphBaseUrl,
    ONEDRIVE_IDENTITY_BASE_URL: identityBaseUrl
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

async function previewTokenFor(name, args = {}) {
  const previewArgs = { ...args };
  delete previewArgs.dryRun;
  delete previewArgs.confirmed;
  delete previewArgs.previewToken;
  const preview = await tool(name, previewArgs);
  assert(!preview.isError, `${name} preview should succeed`, preview);
  assert(preview.value?.dryRun === true, `${name} preview should be a dry-run`, preview.value);
  assert(preview.value?.previewToken, `${name} preview should return a previewToken`, preview.value);
  return preview.value.previewToken;
}

async function toolWithPreview(name, args = {}) {
  const previewToken = await previewTokenFor(name, args);
  return await tool(name, { ...args, previewToken });
}

async function listTools() {
  const response = await request("tools/list", {});
  if (response.error) throw new Error(response.error.message);
  return response.result?.tools || [];
}

function auditEntries() {
  const path = join(mockHome, ".codex", "onedrive-plugin", "audit", "mutations.jsonl");
  try {
    return readFileSync(path, "utf8").trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
  } catch {
    return [];
  }
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

  await check("new efficiency and cleanup tools are registered", async () => {
    const toolList = await listTools();
    const names = new Set(toolList.map((entry) => entry.name));
    const expected = [
      "onedrive_sync_status",
      "onedrive_cache_refresh",
      "onedrive_cache_clear",
      "onedrive_content_index_refresh",
      "onedrive_content_search",
      "onedrive_content_index_clear",
      "onedrive_preview",
      "onedrive_batch_get_info",
      "onedrive_batch_permissions",
      "onedrive_batch_download",
      "onedrive_batch_delete",
      "onedrive_batch_move",
      "onedrive_update_file",
      "onedrive_recent",
      "onedrive_large_files",
      "onedrive_duplicates",
      "onedrive_shared_by_me",
      "onedrive_public_links"
    ];
    const missing = expected.filter((name) => !names.has(name));
    assert(missing.length === 0, "missing new tools", { missing });
    return { checked: expected.length };
  });

  await check("Office Open XML inspection tools and runtime are available", async () => {
    const toolList = await listTools();
    const names = new Set(toolList.map((entry) => entry.name));
    const expected = [
      "onedrive_office_capabilities",
      "onedrive_office_validate",
      "onedrive_office_index_refresh",
      "onedrive_office_search",
      "onedrive_word_get_document",
      "onedrive_excel_get_workbook",
      "onedrive_powerpoint_get_presentation",
      "onedrive_office_backups",
      "onedrive_office_compare_backup",
      "onedrive_office_restore_backup"
    ];
    const missing = expected.filter((name) => !names.has(name));
    assert(missing.length === 0, "missing Office inspection tools", { missing });
    const capabilities = await tool("onedrive_office_capabilities");
    assert(!capabilities.isError, "Office capabilities should succeed", capabilities);
    assert(capabilities.value.runtime.pythonAvailable === true, "Office Python runtime should be available", capabilities.value);
    assert(capabilities.value.runtime.helperAvailable === true, "Office Open XML helper should be available", capabilities.value);
    assert(capabilities.value.backends.openXml.readOnlyToolsReady === true, "Open XML read tools should report ready", capabilities.value);
    assert(capabilities.value.backends.openXml.operations.excel.includes("addTableRow") && capabilities.value.backends.openXml.operations.excel.includes("setTableTotals"), "Open XML table operations should be advertised", capabilities.value);
    assert(capabilities.value.backends.graphExcel.availableForAccount === false, "mock personal drive should report Graph Excel unavailable", capabilities.value);
    return { checked: expected, runtime: capabilities.value.runtime.pythonVersion };
  });

  await check("Office Open XML read and preview-token edit workflows", async () => {
    const word = await tool("onedrive_word_get_document", { itemId: "office-word" });
    const excel = await tool("onedrive_excel_get_workbook", { itemId: "office-excel" });
    const powerpoint = await tool("onedrive_powerpoint_get_presentation", { itemId: "office-powerpoint" });
    assert(!word.isError && word.value.paragraphs?.[0]?.text === "Hello Word", "Word structured read failed", word);
    assert(!excel.isError && excel.value.sheets?.[0]?.cells?.some((cell) => cell.address === "B1" && cell.value === "Revenue"), "Excel structured read failed", excel);
    assert(!powerpoint.isError && powerpoint.value.slides?.[0]?.shapes?.[0]?.text === "Hello PowerPoint", "PowerPoint structured read failed", powerpoint);

    const wordPreview = await tool("onedrive_word_batch_update", { itemId: "office-word", operations: [{ type: "replaceText", find: "Hello Word", replace: "Updated Word" }] });
    assert(!wordPreview.isError && wordPreview.value.dryRun === true && wordPreview.value.previewToken, "Word edit preview failed", wordPreview);
    assert(wordPreview.value.semanticDiff?.operationCounts?.replaceText === 1, "Word preview should include a semantic operation summary", wordPreview);
    const wordLive = await tool("onedrive_word_batch_update", { itemId: "office-word", operations: [{ type: "replaceText", find: "Hello Word", replace: "Updated Word" }], dryRun: false, confirmed: true, expectedId: "office-word", previewToken: wordPreview.value.previewToken });
    officeWordPostMetadataFailuresRemaining = 0;
    assert(!wordLive.isError && wordLive.value.changeCount === 1 && wordLive.value.verificationIncomplete === true, "Word live edit should remain successful when post-commit metadata verification fails", wordLive);
    assert(wordLive.value.localWarnings?.some((entry) => entry.operation === "Office post-commit metadata verification"), "Word live edit should report the post-commit metadata warning", wordLive);

    const excelPreview = await tool("onedrive_excel_batch_update", { itemId: "office-excel", operations: [{ type: "setCell", sheet: "Data", address: "B2", value: "Updated" }] });
    assert(!excelPreview.isError && excelPreview.value.previewToken, "Excel edit preview failed", excelPreview);
    assert(excelPreview.value.semanticDiff?.affectedObjects?.some((entry) => entry.sheet === "Data" && entry.address === "B2"), "Excel preview should identify the affected cell", excelPreview);
    const excelLive = await tool("onedrive_excel_batch_update", { itemId: "office-excel", operations: [{ type: "setCell", sheet: "Data", address: "B2", value: "Updated" }], dryRun: false, confirmed: true, expectedId: "office-excel", previewToken: excelPreview.value.previewToken });
    assert(!excelLive.isError && excelLive.value.changeCount === 1, "Excel live edit failed", excelLive);

    const pptPreview = await tool("onedrive_powerpoint_batch_update", { itemId: "office-powerpoint", operations: [{ type: "replaceText", slideIndex: 0, shapeId: "2", find: "Hello", replace: "Updated" }] });
    assert(!pptPreview.isError && pptPreview.value.previewToken, "PowerPoint edit preview failed", pptPreview);
    assert(pptPreview.value.semanticDiff?.affectedObjects?.some((entry) => entry.slideIndex === 0 && entry.shapeId === "2"), "PowerPoint preview should identify the affected shape", pptPreview);
    const pptLive = await tool("onedrive_powerpoint_batch_update", { itemId: "office-powerpoint", operations: [{ type: "replaceText", slideIndex: 0, shapeId: "2", find: "Hello", replace: "Updated" }], dryRun: false, confirmed: true, expectedId: "office-powerpoint", previewToken: pptPreview.value.previewToken });
    assert(!pptLive.isError && pptLive.value.changeCount === 1, "PowerPoint live edit failed", pptLive);
    const nativePptOperations = [
      { type: "addTextBox", slideIndex: 0, shapeId: 4, name: "Added box", text: "Native text", x: 100, y: 200, width: 300, height: 400 },
      { type: "setTextStyle", slideIndex: 0, shapeId: "4", fontFamily: "Aptos", fontSize: 18, bold: true, color: "12AB34" },
      { type: "addTextBox", slideIndex: 0, shapeId: 5, text: "Delete me", x: 10, y: 20, width: 30, height: 40 },
      { type: "deleteShape", slideIndex: 0, shapeId: "5" },
      { type: "replaceImage", slideIndex: 0, shapeId: "3", base64: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", contentType: "image/png" }
    ];
    const nativePptPreview = await tool("onedrive_powerpoint_batch_update", { itemId: "office-powerpoint", operations: nativePptOperations });
    assert(!nativePptPreview.isError && nativePptPreview.value.changeCount === 5 && nativePptPreview.value.previewToken, "PowerPoint native edit preview failed", nativePptPreview);
    const nativePptLive = await tool("onedrive_powerpoint_batch_update", { itemId: "office-powerpoint", operations: nativePptOperations, dryRun: false, confirmed: true, expectedId: "office-powerpoint", previewToken: nativePptPreview.value.previewToken });
    assert(!nativePptLive.isError && nativePptLive.value.changeCount === 5, "PowerPoint native live edit failed", nativePptLive);

    const updatedWord = await tool("onedrive_word_get_document", { itemId: "office-word" });
    const updatedExcel = await tool("onedrive_excel_get_workbook", { itemId: "office-excel" });
    const updatedPowerpoint = await tool("onedrive_powerpoint_get_presentation", { itemId: "office-powerpoint" });
    assert(updatedWord.value.paragraphs[0].text === "Updated Word", "Word edit did not persist", updatedWord);
    assert(updatedExcel.value.sheets[0].cells.some((cell) => cell.address === "B2" && cell.value === "Updated"), "Excel edit did not persist", updatedExcel);
    assert(updatedPowerpoint.value.slides[0].shapes[0].text === "Updated PowerPoint", "PowerPoint edit did not persist", updatedPowerpoint);
    const pptShapes = new Map(updatedPowerpoint.value.slides[0].shapes.map((shape) => [shape.id, shape]));
    assert(pptShapes.get("4")?.text === "Native text" && !pptShapes.has("5"), "PowerPoint text-box/style/delete edits did not persist", updatedPowerpoint);
    assert(pptShapes.get("3")?.image?.target?.endsWith(".png"), "PowerPoint image relationship was not replaced", updatedPowerpoint);
    for (const id of ["office-word", "office-excel", "office-powerpoint"]) {
      assert(requests.some((entry) => entry.method === "PUT" && entry.path === `/v1.0/me/drive/items/${id}/content`), "Office commit did not target the stable item ID", { id });
    }
    assert(wordLive.value.backup?.backupId, "Word live edit should create a managed backup manifest", wordLive);
    const backups = await tool("onedrive_office_backups", { itemId: "office-word" });
    assert(!backups.isError && backups.value.items.some((entry) => entry.backupId === wordLive.value.backup.backupId), "Managed Word backup should be listed", backups);
    const comparison = await tool("onedrive_office_compare_backup", { backupId: wordLive.value.backup.backupId });
    assert(!comparison.isError && comparison.value.sameContent === false && comparison.value.semanticDiff.changeCount > 0, "Backup comparison should report the Word semantic change", comparison);
    const restorePreview = await tool("onedrive_office_restore_backup", { backupId: wordLive.value.backup.backupId });
    assert(!restorePreview.isError && restorePreview.value.previewToken && restorePreview.value.wouldRestore.currentItem.id === "office-word" && restorePreview.value.semanticDiff?.changeCount > 0, "Office backup restore preview failed", restorePreview);
    const restorePutCount = requests.filter((entry) => entry.method === "PUT" && entry.path === "/v1.0/me/drive/items/office-word/content").length;
    const staleRestore = await tool("onedrive_office_restore_backup", { backupId: wordLive.value.backup.backupId, dryRun: false, confirmed: true, expectedId: "office-word", expectedETag: "stale-etag", previewToken: restorePreview.value.previewToken });
    assert(!staleRestore.isError && staleRestore.value.requiredToRestore?.includes("expectedETag"), "Office backup restore should refuse a stale expected eTag", staleRestore);
    assert(requests.filter((entry) => entry.method === "PUT" && entry.path === "/v1.0/me/drive/items/office-word/content").length === restorePutCount, "Stale Office restore must not upload", { restorePutCount });
    const restored = await tool("onedrive_office_restore_backup", { backupId: wordLive.value.backup.backupId, dryRun: false, confirmed: true, expectedId: "office-word", expectedETag: restorePreview.value.wouldRestore.currentItem.eTag, previewToken: restorePreview.value.previewToken });
    assert(!restored.isError && restored.value.restoredBackupId === wordLive.value.backup.backupId && restored.value.rollbackBackup?.backupId && restored.value.remoteValidation?.package?.fingerprint === restored.value.backupValidation?.package?.fingerprint, "Office backup restore failed", restored);
    officeWordPostMetadataFailuresRemaining = 0;
    const restoredWord = await tool("onedrive_word_get_document", { itemId: "office-word" });
    assert(!restoredWord.isError && restoredWord.value.paragraphs?.[0]?.text === "Hello Word", "Office backup restore did not restore original Word content", restoredWord);
    return { wordChanges: wordLive.value.changeCount, excelChanges: excelLive.value.changeCount, powerpointChanges: pptLive.value.changeCount + nativePptLive.value.changeCount, restoredBackupId: restored.value.restoredBackupId };
  });

  await check("business Excel uses a scoped Graph workbook session", async () => {
    const operations = [{ type: "setCell", sheet: "Data", address: "B2", value: "Graph updated" }];
    const preview = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations });
    assert(!preview.isError && preview.value.backend === "graph", "business Excel preview should select Graph", preview);
    const live = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations, dryRun: false, confirmed: true, expectedId: "office-business", previewToken: preview.value.previewToken });
    assert(!live.isError && live.value.backend === "graph", "business Excel live update should use Graph", live);
    const sessionRequests = requests.filter((entry) => entry.path.includes("/office-business/workbook/"));
    const rangeRequest = sessionRequests.find((entry) => entry.method === "PATCH" && entry.path.includes("/range"));
    assert(sessionRequests.some((entry) => entry.path.endsWith("/createSession")), "Graph createSession was not called", { sessionRequests });
    assert(counters.get("excel-create-session") === 2, "Graph createSession should retry the documented safe 504 once", { sessionRequests });
    assert(counters.get("excel-create-session-poll") === 2, "Graph createSession LRO should poll until succeeded", { sessionRequests });
    assert(sessionRequests.some((entry) => entry.path.includes("/sessionInfoResource")), "Graph createSession LRO resourceLocation was not fetched", { sessionRequests });
    const createRequest = sessionRequests.find((entry) => entry.path.endsWith("/createSession"));
    assert(createRequest?.headers?.prefer === "respond-async", "Graph createSession should request asynchronous completion", { createRequest });
    assert(rangeRequest?.headers?.["workbook-session-id"] === "business-session", "Graph range write omitted workbook-session-id", { rangeRequest });
    assert(sessionRequests.some((entry) => entry.path.endsWith("/closeSession")), "Graph closeSession was not called", { sessionRequests });
    assert(live.value.uploaded?.sessionClosed === false, "failed Graph closeSession should not be reported as closed", live.value);
    assert(live.value.localWarnings?.some((entry) => entry.operation === "Graph Excel session close"), "failed Graph closeSession should be a non-fatal warning", live.value);
    return { calls: sessionRequests.length, backend: live.value.backend, sessionClosed: live.value.uploaded.sessionClosed };
  });

  await check("structured Office index refresh is incremental and returns semantic anchors", async () => {
    const first = await tool("onedrive_office_index_refresh", { itemId: "office-word", refreshMetadata: false, force: true });
    assert(!first.isError && first.value.indexed === 1, "Office structured index refresh failed", first);
    const search = await tool("onedrive_office_search", { query: "Updated Word" });
    assert(!search.isError && search.value.items[0]?.anchor?.type === "paragraph", "Office structured search did not return a paragraph anchor", search);
    const second = await tool("onedrive_office_index_refresh", { itemId: "office-word", refreshMetadata: false });
    assert(!second.isError && second.value.reused === 1 && second.value.graphContentReadsAttempted === 0, "unchanged Office item should reuse the structured index", second);
    return { indexed: first.value.indexed, reused: second.value.reused, anchor: search.value.items[0].anchor };
  });

  await check("cross-file Office batch transformation preflights every file before mutation", async () => {
    const items = [
      { itemId: "office-word", expectedId: "office-word", kind: "word", operations: [{ type: "setParagraphText", paragraphIndex: 0, text: "Batch Word" }] },
      { itemId: "office-excel", expectedId: "office-excel", kind: "excel", operations: [{ type: "setCell", sheet: "Data", address: "D4", value: "Batch Excel" }] },
      { itemId: "office-powerpoint", expectedId: "office-powerpoint", kind: "powerpoint", operations: [{ type: "setShapeText", slideIndex: 0, shapeId: "2", text: "Batch PowerPoint" }] }
    ];
    const preview = await tool("onedrive_office_batch_transform", { items });
    assert(!preview.isError && preview.value.preflightComplete === true && preview.value.itemCount === 3 && preview.value.previewToken, "cross-file preview failed", preview);
    const beforePutCount = requests.filter((entry) => entry.method === "PUT" && entry.path.includes("/office-")).length;
    const live = await tool("onedrive_office_batch_transform", { items, dryRun: false, confirmed: true, previewToken: preview.value.previewToken });
    assert(!live.isError && live.value.partialState === false && live.value.completed.length === 3, "cross-file live transformation failed", live);
    const afterPutCount = requests.filter((entry) => entry.method === "PUT" && entry.path.includes("/office-")).length;
    assert(afterPutCount - beforePutCount === 3, "cross-file transformation should commit each preflighted file exactly once", { beforePutCount, afterPutCount });
    return { itemCount: live.value.itemCount, totalChangeCount: live.value.totalChangeCount };
  });

  await check("cross-file Office batch reports recoverable partial state without replaying writes", async () => {
    const items = [
      { itemId: "office-word", expectedId: "office-word", kind: "word", operations: [{ type: "setParagraphText", paragraphIndex: 0, text: "Partial Word" }] },
      { itemId: "office-excel", expectedId: "office-excel", kind: "excel", operations: [{ type: "setCell", sheet: "Data", address: "E5", value: "Partial Excel" }] }
    ];
    const preview = await tool("onedrive_office_batch_transform", { items });
    assert(!preview.isError && preview.value.preflightComplete, "partial-state batch preview failed", preview);
    failNextOfficeExcelPut = true;
    const live = await tool("onedrive_office_batch_transform", { items, dryRun: false, confirmed: true, previewToken: preview.value.previewToken });
    assert(!live.isError && live.value.partialState === true && live.value.completed.length === 1 && live.value.failed.index === 1, "batch should surface one completed and one failed item", live);
    assert(live.value.recovery?.[0]?.backup?.backupId, "partial-state result should include a managed recovery backup", live.value);
    return { completed: live.value.completed.length, failedIndex: live.value.failed.index, recoveryBackupId: live.value.recovery[0].backup.backupId };
  });

  await check("Graph Excel async session rejects untrusted operation URLs", async () => {
    const operations = [{ type: "setCell", sheet: "Data", address: "B2", value: "Blocked" }];
    const preview = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations });
    const untrustedResource = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations, dryRun: false, confirmed: true, expectedId: "office-business", previewToken: preview.value.previewToken });
    assert(untrustedResource.isError && String(untrustedResource.value).includes("untrusted Excel session resourceLocation"), "untrusted Excel resourceLocation should be rejected", untrustedResource);

    const secondPreview = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations });
    const untrustedLocation = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations, dryRun: false, confirmed: true, expectedId: "office-business", previewToken: secondPreview.value.previewToken });
    assert(untrustedLocation.isError && String(untrustedLocation.value).includes("untrusted Excel session Location"), "untrusted Excel Location should be rejected", untrustedLocation);
    assert(!requests.some((entry) => entry.path.startsWith("/workbook/operations/session") || entry.path.startsWith("/sessionInfoResource")), "untrusted Excel async URL should never be fetched", { requests });
    return { rejected: ["Location", "resourceLocation"] };
  });

  await check("business Excel Graph supports typed table rows and chart lifecycle", async () => {
    const operations = [
      { type: "addTableRow", table: "RevenueTable", values: [["Q4", 42]] },
      { type: "createChart", sheet: "Data", chartType: "ColumnClustered", sourceData: "A1:B4", seriesBy: "Columns" },
      { type: "updateChart", sheet: "Data", chart: "Chart 1", titleText: "Revenue", width: 480, sourceData: "A1:B5", seriesBy: "Columns" }
    ];
    const preview = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations });
    assert(!preview.isError && preview.value.backend === "graph" && preview.value.changeCount === 3, "typed Graph Excel preview failed", preview);
    const before = requests.length;
    const live = await tool("onedrive_excel_batch_update", { itemId: "office-business", backend: "graph", operations, dryRun: false, confirmed: true, expectedId: "office-business", previewToken: preview.value.previewToken });
    assert(!live.isError && live.value.changeCount === 3, "typed Graph Excel live update failed", live);
    const added = requests.slice(before);
    assert(added.some((entry) => entry.method === "POST" && entry.path.includes("/tables/RevenueTable/rows/add")), "table row endpoint was not called", { added });
    assert(added.some((entry) => entry.method === "POST" && entry.path.includes("/worksheets/Data/charts/add")), "chart add endpoint was not called", { added });
    assert(added.some((entry) => entry.method === "PATCH" && entry.path.includes("/charts/Chart%201")), "chart update endpoint was not called", { added });
    assert(added.some((entry) => entry.method === "PATCH" && entry.path.endsWith("/title")), "chart title endpoint was not called", { added });
    assert(added.some((entry) => entry.method === "POST" && entry.path.endsWith("/setData")), "chart setData endpoint was not called", { added });
    return { changeCount: live.value.changeCount, graphCalls: added.filter((entry) => entry.path.includes("/workbook/")).length };
  });

  await check("find schemas expose bounded adaptive search concurrency", async () => {
    const toolList = await listTools();
    const findTool = toolList.find((entry) => entry.name === "onedrive_find");
    const findAllTool = toolList.find((entry) => entry.name === "onedrive_find_all");
    for (const entry of [findTool, findAllTool]) {
      const property = entry?.inputSchema?.properties?.searchConcurrency;
      assert(property?.minimum === 1 && property?.maximum === 4 && property?.default === 2, `${entry?.name || "find tool"} has an invalid searchConcurrency schema`, property);
    }
    return { checked: [findTool.name, findAllTool.name] };
  });

  await check("MSA-only common tenant auth retries device login on consumers endpoint", async () => {
    const result = await tool("onedrive_auth_device_start", { tenant: "common" });
    assert(!result.isError, "device-code login should retry on consumers", result);
    assert(result.value.authTenant === "consumers", "device-code login should report consumers fallback", result.value);
    assert(result.value.userCode.startsWith("MOCK-CODE-"), "device-code login should return mocked consumers response", result.value);
    assert(authRequests.some((request) => request.path === "/common/oauth2/v2.0/devicecode"), "common endpoint was not attempted", authRequests);
    assert(authRequests.some((request) => request.path === "/consumers/oauth2/v2.0/devicecode"), "consumers endpoint was not attempted", authRequests);
    return { authTenant: result.value.authTenant, paths: authRequests.map((request) => request.path) };
  });

  await check("logout requires confirmation before deleting Keychain token", async () => {
    const result = await tool("onedrive_logout", { deleteKeychainToken: true });
    assert(!result.isError, "logout should return a structured confirmation response", result);
    assert(result.value.memoryCleared === true, "logout should still clear in-memory auth state", result.value);
    assert(result.value.keychainTokenDeleted === false, "logout should not delete Keychain token without confirmation", result.value);
    assert(result.value.requiredToDelete?.includes("confirmed: true"), "logout should explain the confirmation requirement", result.value);
    return { keychainTokenDeleted: result.value.keychainTokenDeleted };
  });

  await check("token refresh coalesces and logout prevents late auth persistence", async () => {
    const authHome = join(mockHome, "auth-home");
    const fakeBin = join(authHome, "bin");
    const fakeSecurity = join(fakeBin, "security");
    const keychainPath = join(authHome, "keychain-token.json");
    mkdirSync(fakeBin, { recursive: true });
    writeFileSync(fakeSecurity, `#!/usr/bin/env node
const fs = require("node:fs");
const path = require("node:path");
const tokenPath = process.env.ONEDRIVE_TEST_KEYCHAIN_PATH;
const [command, ...args] = process.argv.slice(2);
if (command === "find-generic-password") {
  if (!fs.existsSync(tokenPath)) process.exit(44);
  process.stdout.write(fs.readFileSync(tokenPath, "utf8"));
  process.exit(0);
}
if (command === "add-generic-password") {
  const passwordIndex = args.indexOf("-w");
  fs.mkdirSync(path.dirname(tokenPath), { recursive: true });
  fs.writeFileSync(tokenPath, args[passwordIndex + 1] || "");
  process.exit(0);
}
if (command === "delete-generic-password") {
  fs.rmSync(tokenPath, { force: true });
  process.exit(0);
}
process.exit(2);
`);
    chmodSync(fakeSecurity, 0o755);
    const expiredToken = () => writeFileSync(keychainPath, JSON.stringify({
      token_type: "Bearer",
      access_token: "expired-access",
      refresh_token: "initial-refresh",
      expires_at: 1
    }));
    expiredToken();
    const authClient = createMcpClient({
      ...process.env,
      PATH: `${fakeBin}:${process.env.PATH}`,
      HOME: authHome,
      ONEDRIVE_CLIENT_ID: "mock-client-id",
      ONEDRIVE_TEST_ACCESS_TOKEN: "",
      ONEDRIVE_GRAPH_BASE_URL: graphBaseUrl,
      ONEDRIVE_IDENTITY_BASE_URL: identityBaseUrl,
      ONEDRIVE_TEST_KEYCHAIN_PATH: keychainPath,
      ONEDRIVE_STORAGE_ROOT: join(authHome, ".codex", "onedrive-plugin"),
      ONEDRIVE_CACHE_ROOT: join(authHome, ".codex", "onedrive-plugin", "cache")
    });
    try {
      await authClient.request("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "auth-race-test", version: "1" } });
      refreshResponseDelayMs = 50;
      const refreshBefore = counters.get("identity-refresh") || 0;
      const concurrent = await Promise.all([
        authClient.tool("onedrive_get_info", { itemId: "root-note" }),
        authClient.tool("onedrive_get_info", { itemId: "delete-target" }),
        authClient.tool("onedrive_get_info", { itemId: "big-text" })
      ]);
      assert(concurrent.every((result) => !result.isError), "concurrent Graph reads should share one successful refresh", concurrent);
      assert((counters.get("identity-refresh") || 0) === refreshBefore + 1, "concurrent callers should issue exactly one token refresh", {
        before: refreshBefore,
        after: counters.get("identity-refresh") || 0
      });

      expiredToken();
      await authClient.tool("onedrive_logout", { deleteKeychainToken: false });
      refreshResponseDelayMs = 100;
      const lateRefreshTarget = (counters.get("identity-refresh") || 0) + 1;
      const lateRefresh = authClient.tool("onedrive_get_info", { itemId: "root-note" });
      await waitForCounter("identity-refresh", lateRefreshTarget);
      const logout = await authClient.tool("onedrive_logout", { deleteKeychainToken: true, confirmed: true });
      assert(logout.value.keychainTokenDeleted === true, "confirmed logout should delete the mocked Keychain token", logout);
      const lateRefreshResult = await lateRefresh;
      assert(lateRefreshResult.isError && String(lateRefreshResult.value).includes("authentication state changed"), "late refresh should be discarded after logout", lateRefreshResult);
      assert(!existsSync(keychainPath), "late refresh must not recreate the Keychain token after logout");

      const deviceStart = await authClient.tool("onedrive_auth_device_start", { tenant: "consumers" });
      assert(!deviceStart.isError, "isolated device login should start", deviceStart);
      devicePollShouldSucceed = true;
      devicePollResponseDelayMs = 100;
      const pollTarget = (counters.get("identity-device-poll") || 0) + 1;
      const latePoll = authClient.tool("onedrive_auth_device_poll");
      await waitForCounter("identity-device-poll", pollTarget);
      await authClient.tool("onedrive_logout", { deleteKeychainToken: true, confirmed: true });
      const latePollResult = await latePoll;
      assert(latePollResult.isError && String(latePollResult.value).includes("authentication state changed"), "late device poll should be discarded after logout", latePollResult);
      assert(!existsSync(keychainPath), "late device poll must not recreate the Keychain token after logout");

      deviceStartResponseDelayMs = 100;
      const delayedStartTarget = (counters.get("identity-device-start") || 0) + 1;
      const delayedStart = authClient.tool("onedrive_auth_device_start", { tenant: "consumers" });
      await waitForCounter("identity-device-start", delayedStartTarget);
      deviceStartResponseDelayMs = 0;
      const replacementStart = await authClient.tool("onedrive_auth_device_start", { tenant: "consumers" });
      assert(!replacementStart.isError, "replacement device login should start", replacementStart);
      const supersededStart = await delayedStart;
      assert(supersededStart.isError && String(supersededStart.value).includes("authentication state changed"), "older delayed device start should be superseded", supersededStart);

      devicePollShouldSucceed = true;
      devicePollResponseDelayMs = 100;
      const supersededPollTarget = (counters.get("identity-device-poll") || 0) + 1;
      const supersededPoll = authClient.tool("onedrive_auth_device_poll");
      await waitForCounter("identity-device-poll", supersededPollTarget);
      const newerStart = await authClient.tool("onedrive_auth_device_start", { tenant: "consumers" });
      assert(!newerStart.isError, "new device login should supersede the in-flight poll", newerStart);
      const supersededPollResult = await supersededPoll;
      assert(supersededPollResult.isError && String(supersededPollResult.value).includes("authentication state changed"), "older successful poll should be discarded after a new start", supersededPollResult);
      assert(!existsSync(keychainPath), "superseded poll must not persist a Keychain token");
      return {
        refreshRequests: (counters.get("identity-refresh") || 0) - refreshBefore,
        lateRefreshDiscarded: true,
        lateDevicePollDiscarded: true,
        supersededDeviceStartDiscarded: true,
        supersededDevicePollDiscarded: true
      };
    } finally {
      refreshResponseDelayMs = 0;
      deviceStartResponseDelayMs = 0;
      devicePollResponseDelayMs = 0;
      devicePollShouldSucceed = false;
      await authClient.close();
    }
  });

  await check("safety, revoke, batch, and audit tools are registered", async () => {
    const toolList = await listTools();
    const names = new Set(toolList.map((entry) => entry.name));
    const expected = [
      "onedrive_invite_permission",
      "onedrive_revoke_permission",
      "onedrive_batch_revoke_permissions",
      "onedrive_batch_move",
      "onedrive_audit_recent",
      "onedrive_audit_export",
      "onedrive_audit_clear"
    ];
    const missing = expected.filter((name) => !names.has(name));
    assert(missing.length === 0, "missing safety tools", { missing });
    return { checked: expected };
  });

  await check("GET requests retry 429 once", async () => {
    counters.set("drive", 0);
    const result = await tool("onedrive_drive");
    assert(!result.isError, "onedrive_drive returned an error", result);
    assert(result.value.name === "Mock OneDrive", "drive response did not come from mock Graph", result.value);
    assert(counters.get("drive") === 2, "expected one retry after 429", { count: counters.get("drive") });
    return { attempts: counters.get("drive") };
  });

  await check("doctor reports healthy mock config and Graph access", async () => {
    const result = await tool("onedrive_doctor", { checkRootList: true, rootListLimit: 2 });
    assert(!result.isError, "doctor should succeed", result);
    assert(result.value.ok === true, "doctor should be healthy", result.value);
    assert(result.value.summary.fail === 0, "doctor should have no failed checks", result.value.summary);
    assert(result.value.checks.some((check) => check.name === "root list"), "doctor should include root list check", result.value.checks);
    return result.value.summary;
  });

  await check("read-only Graph batches retry outer and transient inner failures in order", async () => {
    const result = await tool("onedrive_batch_get_info", {
      items: [{ itemId: "root-note" }, { itemId: "batch-flaky" }],
      format: "full"
    });
    assert(!result.isError, "batch get info should recover from transient failures", result);
    assert(result.value.items.map((entry) => entry.id).join(",") === "root-note,batch-flaky", "batch retries should preserve original result order", result.value.items);
    assert((counters.get("batch-outer") || 0) === 3, "expected one outer retry plus one inner retry batch", { count: counters.get("batch-outer") || 0, bodies: graphBodies });
    assert((counters.get("batch-steady") || 0) === 1, "successful subrequest should not be retried", { count: counters.get("batch-steady") || 0 });
    assert((counters.get("batch-flaky") || 0) === 2, "transient subrequest should be retried once", { count: counters.get("batch-flaky") || 0 });
    const finalBatch = graphBodies.at(-1)?.requests || [];
    assert(finalBatch.length === 1 && String(finalBatch[0].url).includes("batch-flaky"), "inner retry should submit only the failed subrequest", finalBatch);
    return {
      outerAttempts: counters.get("batch-outer"),
      steadyAttempts: counters.get("batch-steady"),
      flakyAttempts: counters.get("batch-flaky"),
      ids: result.value.items.map((entry) => entry.id)
    };
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

  await check("repeated nextLink returns controlled pagination cycle error", async () => {
    const result = await tool("onedrive_list_all", { itemId: "cycle-pager", maxItems: 5 });
    assert(result.isError, "pagination cycle should fail");
    assert(String(result.value).includes("pagination cycle detected"), "cycle error should be specific", result);
    assert(!String(result.value).includes("safeDisplayPath"), "cycle error should not expose a ReferenceError", result);
    return { response: result.value };
  });

  await check("scan stops unique empty pagination chains with a page cap", async () => {
    const result = await tool("onedrive_scan", { itemId: "empty-pager", maxItems: 1, maxFolders: 1 });
    assert(result.isError, "empty pagination chain should fail with a controlled cap");
    assert(String(result.value).includes("pagination exceeded"), "page cap error should be specific", result);
    return { response: result.value };
  });

  await check("raw list and search limits are clamped server-side", async () => {
    const beforeList = requests.length;
    const listResult = await tool("onedrive_list", { path: "/", limit: 99999 });
    assert(!listResult.isError, "list should succeed with clamped raw limit", listResult);
    const listRequest = requests.slice(beforeList).find((request) => request.path === "/v1.0/me/drive/root/children");
    assert(listRequest?.url.includes("%24top=200") || listRequest?.url.includes("$top=200"), "list did not clamp $top", { listRequest });

    const beforeSearch = requests.length;
    const searchResult = await tool("onedrive_search", { query: "anything", limit: 99999 });
    assert(!searchResult.isError, "search should succeed with clamped raw limit", searchResult);
    const searchRequest = requests.slice(beforeSearch).find((request) => request.url.includes("/search(q='"));
    assert(searchRequest?.url.includes("%24top=200") || searchRequest?.url.includes("$top=200"), "search did not clamp $top", { searchRequest });
    const searchSelect = new URL(searchRequest.url, "http://127.0.0.1").searchParams.get("$select") || "";
    assert(searchSelect.includes("id") && searchSelect.includes("name") && searchSelect.includes("parentReference"), "search did not request the bounded default field set", { searchRequest, searchSelect });

    const beforeSearchAll = requests.length;
    const searchAllResult = await tool("onedrive_search_all", { query: "anything", pageSize: 200, maxItems: 5 });
    assert(!searchAllResult.isError, "search_all should succeed with a bounded page size", searchAllResult);
    const searchAllRequest = requests.slice(beforeSearchAll).find((request) => request.url.includes("/search(q='"));
    const searchAllSelect = new URL(searchAllRequest.url, "http://127.0.0.1").searchParams.get("$select") || "";
    assert(searchAllRequest?.url.includes("%24top=5") || searchAllRequest?.url.includes("$top=5"), "search_all did not clamp $top to maxItems", { searchAllRequest });
    assert(searchAllSelect === searchSelect, "search and search_all should request the same bounded default field set", { searchSelect, searchAllSelect });
    return { listUrl: listRequest.url, searchUrl: searchRequest.url, searchAllUrl: searchAllRequest.url };
  });

  await check("list_all clears truncation after the final page", async () => {
    const result = await tool("onedrive_list_all", { itemId: "root", pageSize: 2, maxItems: 10 });
    assert(!result.isError, "list_all should succeed across paginated root children", result);
    assert(result.value.count === 3, "list_all should collect both mock pages", result.value);
    assert(result.value.nextLink === null, "list_all should not retain a consumed nextLink", result.value);
    assert(result.value.truncated === false, "list_all should not report truncation after reaching the final page", result.value);
    return { count: result.value.count, truncated: result.value.truncated };
  });

  await check("oversized Graph pages never return a continuation that skips items", async () => {
    await tool("onedrive_cache_clear");
    const result = await tool("onedrive_list_all", { itemId: "oversized-pager", pageSize: 200, maxItems: 1 });
    assert(!result.isError, "oversized bounded list should return a safe partial result", result);
    assert(result.value.count === 1 && result.value.truncated === true, "oversized page should be reported as truncated", result.value);
    assert(result.value.unsafePageTruncation === true, "unsafe page truncation should be explicit", result.value);
    assert(result.value.nextLink === null, "nextLink would skip the unaccepted page remainder", result.value);
    const cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(Boolean(cache.itemsById["oversized-a"]) && !cache.itemsById["oversized-b"], "only accepted items should be cached", cache.itemsById);
    return { count: result.value.count, unsafePageTruncation: result.value.unsafePageTruncation };
  });

  await check("ordinary pagination never seeds delta cursors", async () => {
    await tool("onedrive_cache_clear");
    const boundedList = await tool("onedrive_list_all", { itemId: "root", pageSize: 2, maxItems: 2 });
    assert(!boundedList.isError && boundedList.value.nextLink, "bounded list should expose an ordinary nextLink", boundedList);
    const afterList = await tool("onedrive_sync_status");
    assert(afterList.value.deltaNextLinkAvailable === false && afterList.value.deltaLinkAvailable === false, "ordinary list pagination must not seed delta state", afterList.value);

    const invalidDelta = await tool("onedrive_delta", { nextLink: boundedList.value.nextLink, maxItems: 5 });
    assert(invalidDelta.isError && String(invalidDelta.value).includes("non-delta pagination cursor"), "delta should reject ordinary pagination cursors", invalidDelta);

    const realDelta = await tool("onedrive_delta", { itemId: "folder-a", maxItems: 5 });
    assert(!realDelta.isError && realDelta.value.deltaLink, "real delta should return a deltaLink", realDelta);
    const afterDelta = await tool("onedrive_sync_status");
    assert(afterDelta.value.deltaLinkAvailable === false, "direct delta must not persist a shared cache-refresh cursor", afterDelta.value);
    const refresh = await tool("onedrive_cache_refresh", { itemId: "folder-a", mode: "scan", replaceCache: true, maxItems: 5, maxFolders: 2, maxDepth: 1 });
    assert(!refresh.isError && refresh.value.cache.deltaLinkAvailable === true, "cache_refresh should persist its matching delta cursor", refresh);
    assert(refresh.value.cache.deltaTarget === "itemId:folder-a", "cache_refresh should persist the cursor target", refresh.value.cache);
    await tool("onedrive_cache_clear");

    const legacyHome = join(mockHome, "legacy-cursor-home");
    const legacyCacheRoot = join(legacyHome, ".codex", "onedrive-plugin", "cache");
    const legacyNextLink = `${graphBaseUrl}/me/drive/items/folder-a/delta?token=legacy-scoped`;
    mkdirSync(legacyCacheRoot, { recursive: true });
    writeFileSync(join(legacyCacheRoot, "metadata-cache.json"), JSON.stringify({
      version: 1,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      deltaLink: null,
      deltaNextLink: legacyNextLink,
      scanRoot: { target: "root" },
      itemCount: 0,
      itemsById: {},
      pathsByLower: {}
    }));
    const legacyClient = createMcpClient({
      ...process.env,
      HOME: legacyHome,
      ONEDRIVE_CLIENT_ID: "mock-client-id",
      ONEDRIVE_TEST_ACCESS_TOKEN: "mock-token",
      ONEDRIVE_GRAPH_BASE_URL: graphBaseUrl,
      ONEDRIVE_IDENTITY_BASE_URL: identityBaseUrl,
      ONEDRIVE_STORAGE_ROOT: join(legacyHome, ".codex", "onedrive-plugin"),
      ONEDRIVE_CACHE_ROOT: legacyCacheRoot
    });
    try {
      await legacyClient.request("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "legacy-cursor-test", version: "1" } });
      const beforeLegacy = requests.length;
      const refreshed = await legacyClient.tool("onedrive_cache_refresh", { mode: "auto", maxItems: 2, maxFolders: 1, maxDepth: 0 });
      assert(!refreshed.isError && refreshed.value.effectiveMode === "scan", "legacy unscoped cursor should fall back to scan", refreshed);
      assert(!requests.slice(beforeLegacy).some((request) => request.url.includes("legacy-scoped")), "legacy cursor without deltaTarget must never be fetched", requests.slice(beforeLegacy));
      const legacyStatus = await legacyClient.tool("onedrive_sync_status");
      assert(legacyStatus.value.deltaNextLinkAvailable === false, "legacy non-delta nextLink should be cleared", legacyStatus.value);
    } finally {
      await legacyClient.close();
    }
    return {
      ordinaryNextLinkIgnored: true,
      invalidCursorRejected: true,
      deltaCursorStored: true,
      legacyCursorCleared: true
    };
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

  await check("runtime validation rejects invalid arguments before handlers", async () => {
    const cases = [
      { name: "onedrive_rename", args: { itemId: "delete-target" }, label: "missing required" },
      { name: "onedrive_get_info", args: { itemId: "delete-target", mystery: true }, label: "unknown property" },
      { name: "onedrive_create_sharing_link", args: { itemId: "copy-src", type: "publish" }, label: "bad enum" },
      { name: "onedrive_doctor", args: { rootListLimit: 99999 }, label: "number maximum" },
      { name: "onedrive_batch_move", args: { items: [] }, label: "array minimum" },
      { name: "onedrive_get_info", args: {}, label: "missing anyOf target" },
      { name: "onedrive_get_info", args: { itemId: "root-note", path: "root-note.txt" }, label: "ambiguous item target" },
      {
        name: "onedrive_write_text",
        args: { remotePath: "root-target.txt", remotePreset: "documents", remoteRelativePath: "preset-target.txt", content: "x" },
        label: "ambiguous write target"
      }
    ];
    const before = requests.length;
    const outputs = [];
    for (const entry of cases) {
      const result = await tool(entry.name, entry.args);
      assert(result.isError, `${entry.label} should be rejected`, result);
      assert(result.value.error === "invalid_arguments", `${entry.label} should return structured validation error`, result.value);
      outputs.push({ label: entry.label, details: result.value.details });
    }
    assert(requests.length === before, "validation failures should not reach Graph", { added: requests.slice(before) });
    return { cases: outputs.length };
  });

  await check("remotePreset requires an explicit relative destination", async () => {
    const before = requests.length;
    const result = await tool("onedrive_write_text", { remotePreset: "documents", content: "hello" });
    assert(result.isError, "remotePreset without remoteRelativePath should fail");
    assert(result.value.error === "invalid_arguments", "unexpected remotePreset error", result);
    assert(result.value.details?.some((detail) => detail.message.includes("remotePreset + remoteRelativePath")), "remotePreset validation should explain target options", result);
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

  await check("create_folder returns full child path for drive-scoped parent references", async () => {
    const result = await tool("onedrive_create_folder", { parentPath: "Folder A", name: "Created Child" });
    assert(!result.isError, "create_folder should succeed", result);
    assert(result.value.remotePath === "Folder A/Created Child", "create_folder should reconstruct full child remotePath", result.value);
    return { remotePath: result.value.remotePath };
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
    const result = await toolWithPreview("onedrive_delete", { itemId: "delete-target", expectedName: "delete-me.txt", dryRun: false, confirmed: true });
    assert(!result.isError, "confirmed delete should succeed", result);
    assert(counters.get("delete") === 1, "expected exactly one DELETE", { deleteCount: counters.get("delete") });
    return { deleteCount: counters.get("delete") };
  });

  await check("batch_delete live action preflights expected identity before any DELETE", async () => {
    const beforeDeleteCount = counters.get("delete") || 0;
    const beforeRequests = requests.length;
    const result = await tool("onedrive_batch_delete", {
      items: [
        { itemId: "delete-target", expectedName: "delete-me.txt" },
        { itemId: "root-note" }
      ],
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch_delete guard should return a structured response", result);
    assert(result.value.requiredToDelete?.includes("expectedName or expectedId"), "batch_delete should require expected identity for every item", result.value);
    assert((counters.get("delete") || 0) === beforeDeleteCount, "batch_delete should not delete any item before preflight passes", { beforeDeleteCount, afterDeleteCount: counters.get("delete") || 0 });
    assert(requests.length === beforeRequests, "batch_delete preflight should not touch Graph when an item is missing expected identity", { added: requests.slice(beforeRequests) });
    return { requiredToDelete: result.value.requiredToDelete, deleteCount: counters.get("delete") || 0 };
  });

  await check("batch_delete reports second-item live failure with partial-state warning", async () => {
    const beforeDeleteCount = counters.get("delete") || 0;
    const beforeDeleteFailCount = counters.get("delete-fail") || 0;
    const result = await toolWithPreview("onedrive_batch_delete", {
      items: [
        { itemId: "delete-target", expectedName: "delete-me.txt" },
        { itemId: "delete-fail", expectedName: "delete-fail.txt" }
      ],
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch_delete second-item failure should be structured", result);
    assert(result.value.failed === true && result.value.failedIndex === 1, "batch_delete should report the second failed item", result.value);
    assert(result.value.partialResults?.length === 1, "batch_delete should report the first completed delete", result.value);
    assert(result.value.warnings?.some((warning) => warning.includes("not atomic")), "batch_delete should warn about partial remote state", result.value);
    assert((counters.get("delete") || 0) === beforeDeleteCount + 1, "batch_delete should delete the first item before second failure", { beforeDeleteCount, afterDeleteCount: counters.get("delete") || 0 });
    assert((counters.get("delete-fail") || 0) === beforeDeleteFailCount + 1, "batch_delete should attempt the second item", { beforeDeleteFailCount, afterDeleteFailCount: counters.get("delete-fail") || 0 });
    return { failedIndex: result.value.failedIndex, warnings: result.value.warnings, partialResults: result.value.partialResults };
  });

  await check("rename, move, and copy default to dry-run without mutation", async () => {
    const before = requests.length;
    const rename = await tool("onedrive_rename", { itemId: "delete-target", newName: "renamed.txt" });
    const move = await tool("onedrive_move", { itemId: "root-note", destinationParentItemId: "folder-a" });
    const copy = await tool("onedrive_copy", { itemId: "copy-src", destinationParentItemId: "folder-a" });
    assert(!rename.isError && rename.value.dryRun === true, "rename should dry-run by default", rename);
    assert(!move.isError && move.value.dryRun === true, "move should dry-run by default", move);
    assert(!copy.isError && copy.value.dryRun === true, "copy should dry-run by default", copy);
    const added = requests.slice(before);
    assert(!added.some((request) => request.method === "PATCH" || request.method === "POST"), "dry-runs should not mutate", { added });
    return { graphRequestsAdded: added.length };
  });

  await check("rename, move, and copy live actions require confirmation", async () => {
    const before = requests.length;
    const rename = await tool("onedrive_rename", { itemId: "delete-target", newName: "renamed.txt", expectedName: "delete-me.txt", dryRun: false });
    const move = await tool("onedrive_move", { itemId: "root-note", destinationParentItemId: "folder-a", expectedId: "root-note", dryRun: false });
    const copy = await tool("onedrive_copy", { itemId: "copy-src", destinationParentItemId: "folder-a", expectedId: "copy-src", dryRun: false });
    assert(rename.value.requiredToRename, "rename should require confirmation", rename.value);
    assert(move.value.requiredToMove, "move should require confirmation", move.value);
    assert(copy.value.requiredToCopy, "copy should require confirmation", copy.value);
    const added = requests.slice(before);
    assert(!added.some((request) => request.method === "PATCH" || request.method === "POST"), "unconfirmed live calls should not mutate", { added });
    return { graphRequestsAdded: added.length };
  });

  await check("rename, move, and copy live actions require expected identity", async () => {
    const before = requests.length;
    const rename = await tool("onedrive_rename", { itemId: "delete-target", newName: "renamed.txt", dryRun: false, confirmed: true });
    const move = await tool("onedrive_move", { itemId: "root-note", destinationParentItemId: "folder-a", dryRun: false, confirmed: true });
    const copy = await tool("onedrive_copy", { itemId: "copy-src", destinationParentItemId: "folder-a", dryRun: false, confirmed: true });
    assert(rename.value.requiredToRename?.includes("expectedName or expectedId"), "rename should require expected identity", rename.value);
    assert(move.value.requiredToMove?.includes("expectedName or expectedId"), "move should require expected identity", move.value);
    assert(copy.value.requiredToCopy?.includes("expectedName or expectedId"), "copy should require expected identity", copy.value);
    const added = requests.slice(before);
    assert(!added.some((request) => request.method === "PATCH" || request.method === "POST"), "missing-identity live calls should not mutate", { added });
    return { graphRequestsAdded: added.length };
  });

  await check("confirmed copy sends one copy POST and preserves manual 303 monitor", async () => {
    const result = await tool("onedrive_copy", {
      itemId: "copy-src",
      destinationParentItemId: "folder-a",
      dryRun: false,
      confirmed: true,
      expectedName: "copy-source.txt",
      waitForCompletion: true,
      timeoutSeconds: 5
    });
    assert(!result.isError, "copy should succeed", result);
    assert(counters.get("copy") === 1, "copy should POST exactly once", { copyCount: counters.get("copy") });
    assert(result.value.monitor?.status === 303, "copy monitor did not preserve 303", result.value.monitor);
    assert(result.value.monitor?.resourceLocation?.includes("/v1.0/me/drive/items/copied"), "missing resource location", result.value.monitor);
    assert(!result.value.monitorUrl?.includes("copy-secret"), "copy response should not expose monitor query tokens", result.value);
    assert(!result.value.monitor?.monitorUrl?.includes("copy-secret"), "copy monitor output should not expose monitor query tokens", result.value.monitor);
    assert(!result.value.monitor?.resourceLocation?.includes("downloadToken"), "copy resource location should not expose query tokens", result.value.monitor);
    return result.value.monitor;
  });

  await check("copy monitor rejects untrusted and insecure external URLs", async () => {
    const result = await tool("onedrive_copy", {
      itemId: "copy-evil",
      dryRun: false,
      confirmed: true,
      expectedName: "copy-evil.txt",
      waitForCompletion: true,
      timeoutSeconds: 5
    });
    assert(!result.isError, "copy acceptance should not be converted into a failed mutation", result);
    assert(result.value.accepted === true, "copy should still report accepted mutation", result.value);
    assert(result.value.monitorError?.includes("untrusted copy monitor URL"), "unexpected monitor rejection", result.value);
    const insecure = await tool("onedrive_copy", {
      itemId: "copy-http-sharepoint",
      dryRun: false,
      confirmed: true,
      expectedName: "copy-http-sharepoint.txt",
      waitForCompletion: true,
      timeoutSeconds: 5
    });
    assert(!insecure.isError, "insecure monitor acceptance should not be converted into a failed mutation", insecure);
    assert(insecure.value.monitorError?.includes("untrusted copy monitor URL"), "insecure external monitor should be rejected", insecure.value);
    assert(!insecure.value.monitorError?.includes("copy-secret"), "monitor rejection should not echo query tokens", insecure.value);
    return { monitorError: result.value.monitorError, insecureMonitorError: insecure.value.monitorError };
  });

  await check("sharing dry-run includes before permission audit", async () => {
    const result = await tool("onedrive_create_sharing_link", { itemId: "copy-src", type: "view", scope: "anonymous" });
    assert(!result.isError, "sharing dry-run should succeed", result);
    assert(result.value.dryRun === true, "sharing dry-run should not mutate", result.value);
    assert(result.value.beforePermissionCount === 1, "sharing dry-run should include before permissions", result.value);
    assert(!counters.get("create-link"), "dry-run should not create a link", { createLinkCount: counters.get("create-link") });
    return { beforePermissionCount: result.value.beforePermissionCount };
  });

  await check("organization sharing dry-run warns on personal drives", async () => {
    const result = await tool("onedrive_create_sharing_link", { itemId: "copy-src", type: "view", scope: "organization" });
    assert(!result.isError, "organization sharing dry-run should return structured preview", result);
    assert(result.value.dryRun === true, "organization sharing preview should not mutate", result.value);
    assert(result.value.warnings?.some((warning) => warning.includes("personal OneDrive")), "personal drive organization scope should warn", result.value);
    return { warnings: result.value.warnings };
  });

  await check("sharing live action returns permission diff", async () => {
    const result = await toolWithPreview("onedrive_create_sharing_link", {
      itemId: "copy-src",
      type: "view",
      scope: "anonymous",
      password: "link-secret",
      expirationDateTime: "2026-12-31T00:00:00Z",
      dryRun: false,
      confirmed: true,
      expectedName: "copy-source.txt"
    });
    assert(!result.isError, "confirmed sharing link should succeed", result);
    const createBody = graphBodies.findLast((entry) => entry.key === "create-link")?.body;
    assert(createBody?.password === "link-secret", "createLink should send password when provided", createBody);
    assert(createBody?.expirationDateTime === "2026-12-31T00:00:00Z", "createLink should send expiration when provided", createBody);
    assert(result.value.permissionDiff?.added?.length === 1, "sharing diff should include the new permission", result.value.permissionDiff);
    assert(result.value.permissionDiff?.beforeCount === 1, "sharing diff should track before count", result.value.permissionDiff);
    const audit = auditEntries().findLast((entry) => entry.tool === "onedrive_create_sharing_link" && entry.status === "success");
    assert(audit?.graphRequestId === "mock-create-link-request", "post-verification reads overwrote the mutation response request ID", audit);
    assert(result.value.permissionDiff?.afterCount === 2, "sharing diff should track after count", result.value.permissionDiff);
    return result.value.permissionDiff;
  });

  await check("post-mutation verification failure does not report a successful mutation as failed", async () => {
    const result = await toolWithPreview("onedrive_create_sharing_link", {
      itemId: "postverify-fail",
      type: "view",
      scope: "anonymous",
      expectedId: "postverify-fail",
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "successful createLink must remain successful when the follow-up permission read fails", result);
    assert(result.value.permission?.id === "postverify-link", "remote mutation result should be returned", result.value);
    assert(result.value.verificationIncomplete === true, "response should disclose incomplete post-mutation verification", result.value);
    assert(result.value.localWarnings?.some((warning) => warning.operation === "post-mutation permission verification"), "verification failure should be included as a local warning", result.value);
    assert((counters.get("postverify-create-link") || 0) === 1, "createLink should not be retried after remote success", { count: counters.get("postverify-create-link") });
    return { verificationIncomplete: result.value.verificationIncomplete, localWarnings: result.value.localWarnings };
  });

  await check("invite permission dry-run and live behavior", async () => {
    const beforeInviteCount = counters.get("invite-permission") || 0;
    const dryRun = await tool("onedrive_invite_permission", {
      itemId: "copy-src",
      recipients: [{ email: "person@example.test" }]
    });
    assert(!dryRun.isError, "invite dry-run should succeed", dryRun);
    assert(dryRun.value.dryRun === true, "invite should dry-run by default", dryRun.value);
    assert(dryRun.value.wouldInvite?.sendInvitation === false, "invite should default to silent grant", dryRun.value);
    assert(dryRun.value.wouldInvite?.requireSignIn === true, "invite should require sign-in by default", dryRun.value);
    assert(dryRun.value.wouldInvite?.recipientCount === 1, "invite dry-run should summarize recipients", dryRun.value);
    assert((counters.get("invite-permission") || 0) === beforeInviteCount, "dry-run should not POST invite", { beforeInviteCount, afterInviteCount: counters.get("invite-permission") || 0 });

    const noConfirm = await tool("onedrive_invite_permission", {
      itemId: "copy-src",
      recipients: [{ email: "person@example.test" }],
      expectedId: "copy-src",
      dryRun: false
    });
    assert(!noConfirm.isError, "invite no-confirm should be structured", noConfirm);
    assert(noConfirm.value.requiredToInvite, "invite should require confirmation", noConfirm.value);
    assert((counters.get("invite-permission") || 0) === beforeInviteCount, "unconfirmed invite should not POST", { beforeInviteCount, afterInviteCount: counters.get("invite-permission") || 0 });

    const missingExpected = await tool("onedrive_invite_permission", {
      itemId: "copy-src",
      recipients: [{ email: "person@example.test" }],
      dryRun: false,
      confirmed: true
    });
    assert(!missingExpected.isError, "invite missing expected identity should be structured", missingExpected);
    assert(missingExpected.value.requiredToInvite?.includes("expectedName or expectedId"), "invite should require expected identity", missingExpected.value);
    assert((counters.get("invite-permission") || 0) === beforeInviteCount, "missing expected identity should not POST", { beforeInviteCount, afterInviteCount: counters.get("invite-permission") || 0 });

    const live = await toolWithPreview("onedrive_invite_permission", {
      itemId: "copy-src",
      recipients: [{ email: "person@example.test" }],
      password: "invite-secret",
      expirationDateTime: "2026-12-31T00:00:00Z",
      dryRun: false,
      confirmed: true,
      expectedName: "copy-source.txt"
    });
    assert(!live.isError, "live silent invite should succeed", live);
    assert(counters.get("invite-permission") === beforeInviteCount + 1, "live invite should POST once", { count: counters.get("invite-permission") });
    const inviteBody = graphBodies.findLast((entry) => entry.key === "invite-permission")?.body;
    assert(inviteBody?.recipients?.[0]?.email === "person@example.test", "invite should send recipient email", inviteBody);
    assert(inviteBody?.roles?.[0] === "read", "invite should default to read role", inviteBody);
    assert(inviteBody?.sendInvitation === false, "invite should default to silent direct grant", inviteBody);
    assert(inviteBody?.requireSignIn === true, "invite should default requireSignIn to true", inviteBody);
    assert(inviteBody?.password === "invite-secret", "invite should send password when provided", inviteBody);
    assert(inviteBody?.expirationDateTime === "2026-12-31T00:00:00Z", "invite should send expiration when provided", inviteBody);
    assert(live.value.permissionDiff?.added?.some((permission) => permission.id === "perm-invite"), "invite diff should show added permission", live.value.permissionDiff);
    return live.value.permissionDiff;
  });

  await check("invite permission honors explicit email invitations and reports failures", async () => {
    const liveEmail = await toolWithPreview("onedrive_invite_permission", {
      itemId: "copy-src",
      recipients: [{ email: "person@example.test" }],
      role: "write",
      sendInvitation: true,
      message: "Please review",
      dryRun: false,
      confirmed: true,
      expectedId: "copy-src"
    });
    assert(!liveEmail.isError, "live email invite should succeed", liveEmail);
    const inviteBody = graphBodies.findLast((entry) => entry.key === "invite-permission")?.body;
    assert(inviteBody?.sendInvitation === true, "sendInvitation true should be honored", inviteBody);
    assert(inviteBody?.roles?.[0] === "write", "write role should be honored", inviteBody);
    assert(inviteBody?.message === "Please review", "message should be sent when provided", inviteBody);

    const failed = await toolWithPreview("onedrive_invite_permission", {
      itemId: "invite-fail",
      recipients: [{ email: "person@example.test" }],
      dryRun: false,
      confirmed: true,
      expectedName: "invite-fail.txt"
    });
    assert(failed.isError, "failed invite should return tool error", failed);
    assert(String(failed.value).includes("mock invite failure"), "invite failure should surface Graph message", failed);
    assert(!String(failed.value).includes("person@example.test"), "invite failure should redact recipient email", failed);
    assert(!String(failed.value).includes("11111111-1111-1111-1111-111111111111"), "invite failure should redact object identifiers", failed);
    assert(!String(failed.value).includes("https://example.test"), "invite failure should redact URLs", failed);
    assert(!String(failed.value).includes("abc.def.ghi"), "invite failure should redact bearer-looking tokens", failed);
    return { sendInvitation: inviteBody.sendInvitation, failure: String(failed.value) };
  });

  await check("invite permission recipient validation runs before Graph mutation", async () => {
    const before = requests.length;
    const beforeInviteCount = counters.get("invite-permission") || 0;
    const result = await tool("onedrive_invite_permission", {
      itemId: "copy-src",
      recipients: [{ email: "person@example.test", alias: "person" }]
    });
    assert(result.isError, "ambiguous recipient should fail");
    assert(String(result.value).includes("exactly one of email, alias, or objectId"), "unexpected recipient validation error", result);
    const added = requests.slice(before);
    assert(!added.some((request) => request.method === "POST" && request.path.endsWith("/invite")), "recipient validation should not mutate", { added });
    assert((counters.get("invite-permission") || 0) === beforeInviteCount, "recipient validation should not POST invite", { beforeInviteCount, afterInviteCount: counters.get("invite-permission") || 0 });
    return { graphRequestsAdded: added.length };
  });

  await check("compact permissions include grantedToIdentitiesV2", async () => {
    const result = await tool("onedrive_permissions", { itemId: "copy-src" });
    assert(!result.isError, "permissions should succeed", result);
    const owner = result.value.permissions.find((permission) => permission.id === "perm-owner");
    assert(owner?.grantedToIdentitiesV2?.[0]?.type === "user", "compact permissions should include V2 identity type", owner);
    assert(owner?.grantedToIdentitiesV2?.[0]?.email === "mock@example.test", "compact permissions should include V2 identity email", owner);
    return owner.grantedToIdentitiesV2;
  });

  await check("batch_move preflight prevents partial mutation", async () => {
    const beforeMoveCount = counters.get("move-root-note") || 0;
    const result = await tool("onedrive_batch_move", {
      items: [
        { itemId: "root-note", expectedId: "root-note" },
        { itemId: "delete-target", expectedName: "wrong-name.txt" }
      ],
      destinationParentItemId: "folder-a",
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch_move preflight failure should return structured result", result);
    assert(result.value.preflightFailed === true, "batch_move should report preflight failure", result.value);
    assert((counters.get("move-root-note") || 0) === beforeMoveCount, "batch_move should not PATCH any item after preflight failure", { beforeMoveCount, afterMoveCount: counters.get("move-root-note") || 0 });
    return { errors: result.value.errors };
  });

  await check("batch_move reports second-item live failure with partial-state warning", async () => {
    const beforeMoveCount = counters.get("move-root-note") || 0;
    const beforeMoveFailCount = counters.get("move-fail") || 0;
    const result = await tool("onedrive_batch_move", {
      items: [
        { itemId: "root-note", expectedId: "root-note" },
        { itemId: "move-fail", expectedName: "move-fail.txt" }
      ],
      destinationParentItemId: "folder-a",
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch_move second-item failure should be structured", result);
    assert(result.value.failed === true && result.value.failedIndex === 1, "batch_move should report the second failed item", result.value);
    assert(result.value.partialResults?.length === 1, "batch_move should report the first completed move", result.value);
    assert(result.value.warnings?.some((warning) => warning.includes("not atomic")), "batch_move should warn about partial remote state", result.value);
    assert((counters.get("move-root-note") || 0) === beforeMoveCount + 1, "batch_move should move the first item before second failure", { beforeMoveCount, afterMoveCount: counters.get("move-root-note") || 0 });
    assert((counters.get("move-fail") || 0) === beforeMoveFailCount + 1, "batch_move should attempt the second item", { beforeMoveFailCount, afterMoveFailCount: counters.get("move-fail") || 0 });
    return { failedIndex: result.value.failedIndex, warnings: result.value.warnings, partialResults: result.value.partialResults };
  });

  await check("revoke permission dry-run and live behavior", async () => {
    const dryRun = await tool("onedrive_revoke_permission", { itemId: "revoke-target", permissionId: "perm-public" });
    assert(!dryRun.isError, "revoke dry-run should succeed", dryRun);
    assert(dryRun.value.dryRun === true, "revoke should dry-run by default", dryRun.value);
    assert(dryRun.value.beforePermissions?.some((permission) => permission.id === "perm-public"), "dry-run should include before permissions", dryRun.value);
    assert(!counters.get("revoke-perm-public"), "dry-run should not DELETE permission", { count: counters.get("revoke-perm-public") });

    const noConfirm = await tool("onedrive_revoke_permission", {
      itemId: "revoke-target",
      permissionId: "perm-public",
      expectedName: "shared-doc.txt",
      dryRun: false
    });
    assert(!noConfirm.isError, "revoke no-confirm should be structured", noConfirm);
    assert(noConfirm.value.requiredToRevoke, "revoke should require confirmation", noConfirm.value);
    assert(!counters.get("revoke-perm-public"), "unconfirmed revoke should not DELETE permission", { count: counters.get("revoke-perm-public") });

    const live = await toolWithPreview("onedrive_revoke_permission", {
      itemId: "revoke-target",
      permissionId: "perm-public",
      expectedName: "shared-doc.txt",
      dryRun: false,
      confirmed: true
    });
    assert(!live.isError, "live revoke should succeed", live);
    assert(counters.get("revoke-perm-public") === 1, "live revoke should DELETE once", { count: counters.get("revoke-perm-public") });
    assert(live.value.permissionDiff?.removed?.some((permission) => permission.id === "perm-public"), "live revoke diff should show removed permission", live.value.permissionDiff);
    return live.value.permissionDiff;
  });

  await check("revoke owner permission preview is marked non-revocable", async () => {
    const dryRun = await tool("onedrive_revoke_permission", { itemId: "root-note", permissionId: "perm-owner-root-note" });
    assert(!dryRun.isError, "owner revoke dry-run should return structured preview", dryRun);
    assert(dryRun.value.wouldRevoke.revocable === false, "owner permission should be marked non-revocable", dryRun.value);
    assert(dryRun.value.warnings?.length, "owner permission should include warning", dryRun.value);
    const live = await tool("onedrive_revoke_permission", {
      itemId: "root-note",
      permissionId: "perm-owner-root-note",
      expectedId: "root-note",
      dryRun: false,
      confirmed: true
    });
    assert(!live.isError, "owner live revoke should be refused structurally", live);
    assert(live.value.requiredToRevoke?.includes("Owner"), "owner live revoke should explain refusal", live.value);
    return { warning: dryRun.value.warnings[0], required: live.value.requiredToRevoke };
  });

  await check("batch revoke preflight prevents partial deletion", async () => {
    const beforeRevokeCount = counters.get("revoke-perm-public") || 0;
    const result = await tool("onedrive_batch_revoke_permissions", {
      items: [
        { itemId: "copy-src", permissionId: "perm-owner", expectedId: "copy-src" },
        { itemId: "revoke-target", permissionId: "missing-permission", expectedId: "revoke-target" }
      ],
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch revoke preflight failure should return structured result", result);
    assert(result.value.preflightFailed === true, "batch revoke should report preflight failure", result.value);
    assert((counters.get("revoke-perm-public") || 0) === beforeRevokeCount, "batch revoke should not DELETE after preflight failure", { beforeRevokeCount, afterRevokeCount: counters.get("revoke-perm-public") || 0 });
    return { errors: result.value.errors };
  });

  await check("batch revoke validates permission existence even when permissions are omitted from output", async () => {
    const beforeRevokeCount = counters.get("revoke-perm-a") || 0;
    const result = await tool("onedrive_batch_revoke_permissions", {
      items: [
        { itemId: "revoke-a", permissionId: "perm-a", expectedId: "revoke-a" },
        { itemId: "revoke-b", permissionId: "missing-permission", expectedId: "revoke-b" }
      ],
      includePermissions: false,
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch revoke preflight failure should be structured", result);
    assert(result.value.preflightFailed === true, "batch revoke should fail preflight before mutation", result.value);
    assert((counters.get("revoke-perm-a") || 0) === beforeRevokeCount, "batch revoke should not partially revoke when includePermissions is false", { beforeRevokeCount, afterRevokeCount: counters.get("revoke-perm-a") || 0 });
    return { errors: result.value.errors };
  });

  await check("batch revoke reports second-item live failure with partial-state warning", async () => {
    const beforeRevokeCount = counters.get("revoke-perm-a") || 0;
    const beforeRevokeFailCount = counters.get("revoke-perm-fail") || 0;
    const result = await toolWithPreview("onedrive_batch_revoke_permissions", {
      items: [
        { itemId: "revoke-a", permissionId: "perm-a", expectedId: "revoke-a" },
        { itemId: "revoke-fail", permissionId: "perm-fail", expectedId: "revoke-fail" }
      ],
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "batch revoke second-item failure should be structured", result);
    assert(result.value.failed === true && result.value.failedIndex === 1, "batch revoke should report the second failed item", result.value);
    assert(result.value.partialResults?.length === 1, "batch revoke should report the first completed revoke", result.value);
    assert(result.value.warnings?.some((warning) => warning.includes("not atomic")), "batch revoke should warn about partial remote state", result.value);
    assert((counters.get("revoke-perm-a") || 0) === beforeRevokeCount + 1, "batch revoke should revoke the first permission before second failure", { beforeRevokeCount, afterRevokeCount: counters.get("revoke-perm-a") || 0 });
    assert((counters.get("revoke-perm-fail") || 0) === beforeRevokeFailCount + 1, "batch revoke should attempt the second permission", { beforeRevokeFailCount, afterRevokeFailCount: counters.get("revoke-perm-fail") || 0 });
    return { failedIndex: result.value.failedIndex, warnings: result.value.warnings, partialResults: result.value.partialResults };
  });

  await check("download refuses local OneDrive sync destination by default", async () => {
    const result = await tool("onedrive_download", {
      itemId: "delete-target",
      localPath: join(mockHome, "Library", "CloudStorage", "OneDrive-Personal", "blocked-download.txt")
    });
    assert(result.isError, "download to local OneDrive sync path should fail");
    assert(String(result.value).includes("local OneDrive sync folder"), "unexpected sync-path guard message", result);
    return { blocked: true };
  });

  await check("upload refuses local OneDrive sync source before local or Graph work", async () => {
    const before = requests.length;
    const result = await tool("onedrive_upload", {
      localPath: join(mockHome, "Library", "CloudStorage", "OneDrive-Personal", "blocked-upload.txt"),
      remotePath: "Blocked Upload.txt"
    });
    assert(result.isError, "upload from local OneDrive sync path should fail");
    assert(String(result.value).includes("local OneDrive sync folder"), "unexpected sync-path guard message", result);
    assert(requests.length === before, "upload guard should not reach mock Graph", { before, after: requests.length });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("zero-byte session upload falls back to simple upload", async () => {
    const localPath = join(mockHome, "work", "empty-upload.txt");
    mkdirSync(dirname(localPath), { recursive: true });
    writeFileSync(localPath, "");
    const before = requests.length;
    const result = await tool("onedrive_upload", {
      localPath,
      remotePath: "empty-session.txt",
      uploadMode: "session"
    });
    assert(!result.isError, "zero-byte upload should succeed", result);
    assert(result.value.uploadMode === "simple", "zero-byte session upload should use simple upload", result.value);
    assert(result.value.bytesUploaded === 0, "zero-byte upload should report zero bytes", result.value);
    const added = requests.slice(before);
    assert(added.some((request) => request.method === "PUT" && request.path === "/v1.0/me/drive/root:/empty-session.txt:/content"), "simple upload endpoint was not used", { added });
    assert(!added.some((request) => request.url.includes("createUploadSession")), "zero-byte upload should not create an upload session", { added });
    return { uploadMode: result.value.uploadMode, bytesUploaded: result.value.bytesUploaded };
  });

  await check("upload session URLs are trusted before sending file bytes", async () => {
    const localPath = join(mockHome, "work", "session-upload.txt");
    mkdirSync(dirname(localPath), { recursive: true });
    writeFileSync(localPath, "session upload body", "utf8");
    const trusted = await tool("onedrive_upload", {
      localPath,
      remotePath: "trusted-session.txt",
      uploadMode: "session",
      chunkSize: 327680
    });
    assert(!trusted.isError, "trusted same-origin upload session should succeed", trusted);
    assert(counters.get("trusted-upload-session-put") === 1, "trusted upload session should PUT one chunk", { count: counters.get("trusted-upload-session-put") });

    const beforePut = counters.get("trusted-upload-session-put") || 0;
    const evil = await tool("onedrive_upload", {
      localPath,
      remotePath: "evil-session.txt",
      uploadMode: "session",
      chunkSize: 327680
    });
    assert(evil.isError, "untrusted upload session should fail");
    assert(String(evil.value).includes("untrusted upload session URL"), "unexpected untrusted upload error", evil);
    assert((counters.get("trusted-upload-session-put") || 0) === beforePut, "untrusted upload should not send chunks to mock upload endpoint", { beforePut, afterPut: counters.get("trusted-upload-session-put") || 0 });
    return { trusted: trusted.value.uploadMode, evil: evil.value };
  });

  await check("document PDF export writes converted content", async () => {
    const localPath = join(pluginRoot, "work", "mock-export.pdf");
    const result = await tool("onedrive_export_pdf", { itemId: "quarterly-report", localPath, overwrite: true });
    assert(!result.isError, "PDF export should succeed", result);
    assert(result.value.exportFormat === "pdf", "unexpected export format", result.value);
    assert(result.value.bytesWritten > 0, "PDF export should write bytes", result.value);
    return { bytesWritten: result.value.bytesWritten, localPath: result.value.localPath };
  });

  await check("document text export writes converted content", async () => {
    const localPath = join(pluginRoot, "work", "mock-export.txt");
    const result = await tool("onedrive_export_text", { itemId: "quarterly-report", localPath, overwrite: true });
    assert(!result.isError, "text export should succeed", result);
    assert(result.value.exportFormat === "text", "unexpected export format", result.value);
    assert(result.value.bytesWritten > 0, "text export should write bytes", result.value);
    return { bytesWritten: result.value.bytesWritten, localPath: result.value.localPath };
  });

  await check("Office download helpers reuse resolved metadata", async () => {
    const localPath = join(pluginRoot, "work", "mock-download-word.docx");
    const before = requests.length;
    const result = await tool("onedrive_download_word", { itemId: "quarterly-report", localPath, overwrite: true });
    assert(!result.isError, "Word download helper should succeed", result);
    const added = requests.slice(before);
    const metadataRequests = added.filter((request) => request.method === "GET" && request.path === "/v1.0/me/drive/items/quarterly-report");
    assert(metadataRequests.length === 1, "Word download helper fetched metadata more than once", { metadataRequests, added });
    assert(readFileSync(localPath, "utf8").includes("Quarterly Report raw content"), "Word helper wrote unexpected content", result.value);
    rmSync(localPath, { force: true });
    return { metadataRequests: metadataRequests.length, bytesWritten: result.value.bytesWritten };
  });

  await check("concurrent default downloads and exports reserve distinct destinations", async () => {
    const managedDownloads = join(mockHome, ".codex", "onedrive-plugin", "downloads");
    for (const name of ["same-name.txt", "same-name (2).txt"]) rmSync(join(managedDownloads, name), { force: true });
    const downloads = await Promise.all([
      tool("onedrive_download", { itemId: "dup-a" }),
      tool("onedrive_download", { itemId: "dup-b" })
    ]);
    assert(downloads.every((entry) => !entry.isError), "concurrent default downloads should both succeed", downloads);
    const downloadPaths = downloads.map((entry) => entry.value.localPath);
    assert(new Set(downloadPaths).size === 2, "concurrent downloads returned the same destination", downloadPaths);
    const bodies = downloadPaths.map((path) => readFileSync(path, "utf8")).sort();
    assert(bodies[0].includes("duplicate a") && bodies[1].includes("duplicate b"), "concurrent download bodies were overwritten", bodies);

    const exportRoot = join(managedDownloads, "export");
    for (const name of ["Quarterly Report.pdf", "Quarterly Report (2).pdf"]) rmSync(join(exportRoot, name), { force: true });
    const exports = await Promise.all([
      tool("onedrive_export_pdf", { itemId: "quarterly-report" }),
      tool("onedrive_export_pdf", { itemId: "quarterly-report" })
    ]);
    assert(exports.every((entry) => !entry.isError), "concurrent default exports should both succeed", exports);
    const exportPaths = exports.map((entry) => entry.value.localPath);
    assert(new Set(exportPaths).size === 2, "concurrent exports returned the same destination", exportPaths);
    assert(exportPaths.every((path) => readFileSync(path, "utf8").includes("%PDF")), "concurrent export output was invalid", exportPaths);
    for (const path of [...downloadPaths, ...exportPaths]) rmSync(path, { force: true });
    return { downloadPaths, exportPaths };
  });

  await check("preview returns bounded document text export", async () => {
    const result = await tool("onedrive_preview", { itemId: "quarterly-report", maxBytes: 12 });
    assert(!result.isError, "preview should succeed", result);
    assert(result.value.source === "graph-text-export", "preview should use Graph text export for docx", result.value);
    assert(result.value.preview.length <= 12, "preview should be bounded", result.value);
    assert(result.value.truncated === true, "preview should report truncation", result.value);
    return { source: result.value.source, preview: result.value.preview };
  });

  await check("read_text refuses oversized content without full-read semantics", async () => {
    const result = await tool("onedrive_read_text", { itemId: "big-text", maxBytes: 10 });
    assert(result.isError, "read_text should refuse content above maxBytes", result);
    assert(String(result.value).includes("above maxBytes 10"), "unexpected oversize message", result);
    return { response: result.value };
  });

  await check("read_text honors the advertised ten MiB hard limit", async () => {
    const requestedMaxBytes = 8 * 1024 * 1024;
    const result = await tool("onedrive_read_text", { itemId: "six-meg-text", maxBytes: requestedMaxBytes });
    assert(!result.isError, "a valid maxBytes above the five MiB default should not be silently clamped", result);
    assert(result.value.item.size === 6 * 1024 * 1024, "expected six MiB metadata fixture", result.value.item);
    assert(result.value.content === "bounded six meg metadata test", "unexpected bounded content", result.value);
    return { requestedMaxBytes, itemSize: result.value.item.size };
  });

  await check("preview truncates oversized text content safely", async () => {
    const result = await tool("onedrive_preview", { itemId: "big-text", maxBytes: 10 });
    assert(!result.isError, "preview should succeed", result);
    assert(result.value.source === "text-read", "preview should use text path for text files", result.value);
    assert(result.value.preview === "0123456789", "preview should return bounded prefix", result.value);
    assert(result.value.bytes <= 10, "preview returned bytes should not exceed maxBytes", result.value);
    assert(result.value.bytesRead >= result.value.bytes, "preview should retain total bytesRead for diagnostics", result.value);
    assert(result.value.truncated === true, "preview should report truncation", result.value);
    return { preview: result.value.preview, bytes: result.value.bytes };
  });

  await check("update_file checkout refuses to overwrite existing manifest", async () => {
    const manifestPath = join(mockHome, "work", "existing-manifest.json");
    const localPath = join(mockHome, "work", "checkout-root-note.txt");
    mkdirSync(dirname(manifestPath), { recursive: true });
    writeFileSync(manifestPath, "do not replace", "utf8");
    const before = requests.length;
    const result = await tool("onedrive_update_file", {
      mode: "checkout",
      remotePath: "root-note.txt",
      itemId: "root-note",
      localPath,
      manifestPath
    });
    assert(result.isError, "checkout should refuse existing manifest", result);
    assert(readFileSync(manifestPath, "utf8") === "do not replace", "manifest was overwritten");
    const added = requests.slice(before);
    assert(!added.some((request) => request.path.endsWith("/content")), "checkout should not download after manifest refusal", { added });
    return { response: result.value, graphRequestsAdded: added.length };
  });

  await check("update_file checkout rejects local sync path before Graph", async () => {
    const localPath = join(mockHome, "Library", "CloudStorage", "OneDrive-Personal", "blocked-checkout.txt");
    const before = requests.length;
    const result = await tool("onedrive_update_file", {
      mode: "checkout",
      remotePath: "root-note.txt",
      localPath
    });
    assert(result.isError, "checkout should reject local sync folder targets", result);
    assert(String(result.value).includes("local OneDrive sync folder"), "unexpected checkout sync-folder error", result);
    assert(requests.length === before, "checkout sync-folder validation should not reach Graph", { before, after: requests.length, added: requests.slice(before) });
    return { graphRequestsAdded: requests.length - before };
  });

  await check("update_file commit sends checkout eTag as If-Match", async () => {
    const localPath = join(mockHome, "work", "commit-root-note.txt");
    const manifestPath = `${localPath}.onedrive-update.json`;
    mkdirSync(dirname(localPath), { recursive: true });
    writeFileSync(localPath, "edited root note\n", "utf8");
    writeFileSync(manifestPath, JSON.stringify({
      version: 1,
      checkedOutAt: "2026-07-04T00:00:00.000Z",
      remotePath: "root-note.txt",
      item: item("root-note", "root-note.txt"),
      localPath
    }, null, 2));
    const before = requests.length;
    const result = await tool("onedrive_update_file", {
      mode: "commit",
      remotePath: "root-note.txt",
      localPath,
      manifestPath
    });
    assert(!result.isError, "commit should succeed", result);
    const uploadRequest = requests.slice(before).find((request) => request.method === "PUT" && request.path === "/v1.0/me/drive/root:/root-note.txt:/content");
    assert(uploadRequest, "commit upload request was not observed", { added: requests.slice(before) });
    assert(uploadRequest.headers["if-match"] === "etag-root-note", "commit upload should include checkout eTag If-Match", { headers: uploadRequest.headers });
    const updateAudit = auditEntries().filter((entry) => entry.tool === "onedrive_update_file");
    assert(updateAudit.length >= 1, "commit should be audited as onedrive_update_file", auditEntries());
    return { ifMatch: uploadRequest.headers["if-match"], uploadCount: counters.get("root-note-upload"), updateAuditCount: updateAudit.length };
  });

  await check("batch_download refuses item local sync path before Graph", async () => {
    const before = requests.length;
    const result = await tool("onedrive_batch_download", {
      items: [{
        itemId: "root-note",
        localPath: join(mockHome, "Library", "CloudStorage", "OneDrive-Personal", "blocked-batch-download.txt")
      }]
    });
    assert(!result.isError, "batch_download should report per-item error without failing the whole batch", result);
    assert(result.value.results[0]?.error?.includes("local OneDrive sync folder"), "unexpected batch error", result.value);
    const added = requests.slice(before);
    assert(added.length === 0, "batch_download should reject unsafe local path before Graph", { added });
    return { error: result.value.results[0].error };
  });

  await check("batch_download gives duplicate generated filenames unique targets", async () => {
    const destinationFolder = join(mockHome, "batch-download-dupes");
    const result = await tool("onedrive_batch_download", {
      items: [{ itemId: "dup-a" }, { itemId: "dup-b" }],
      destinationFolder
    });
    assert(!result.isError, "batch_download should succeed", result);
    const paths = result.value.results.map((entry) => entry.localPath);
    assert(paths.length === 2 && new Set(paths).size === 2, "batch_download local paths should be unique", result.value);
    assert(paths[0].endsWith("same-name.txt"), "first duplicate should keep original filename", { paths });
    assert(paths[1].endsWith("same-name (2).txt"), "second duplicate should get deterministic suffix", { paths });
    assert(readFileSync(paths[0], "utf8") === "duplicate a\n", "first duplicate content mismatch", { paths });
    assert(readFileSync(paths[1], "utf8") === "duplicate b\n", "second duplicate content mismatch", { paths });
    return { paths };
  });

  await check("batch_download uniquifies duplicate names in the default destination", async () => {
    const result = await tool("onedrive_batch_download", {
      items: [{ itemId: "dup-a" }, { itemId: "dup-b" }]
    });
    assert(!result.isError, "default batch download should succeed", result);
    const paths = result.value.results.map((entry) => entry.localPath);
    assert(paths.length === 2 && new Set(paths).size === 2, "default batch targets should be unique", result.value);
    assert(paths[0].endsWith("same-name.txt") && paths[1].endsWith("same-name (2).txt"), "default targets should use deterministic suffixes", { paths });
    assert(readFileSync(paths[0], "utf8") === "duplicate a\n", "first default duplicate content mismatch", { paths });
    assert(readFileSync(paths[1], "utf8") === "duplicate b\n", "second default duplicate content mismatch", { paths });
    return { paths };
  });

  await check("batch_move live action requires confirmation and expected identity before Graph", async () => {
    const beforeNoConfirm = requests.length;
    const noConfirm = await tool("onedrive_batch_move", {
      items: [{ itemId: "root-note", expectedId: "root-note" }],
      destinationParentItemId: "folder-a",
      dryRun: false
    });
    assert(!noConfirm.isError, "batch_move no-confirm response should be structured", noConfirm);
    assert(noConfirm.value.requiredToMove, "batch_move should require confirmation", noConfirm.value);
    assert(requests.length === beforeNoConfirm, "batch_move should not touch Graph before confirmation", { added: requests.slice(beforeNoConfirm) });

    const beforeMissingExpected = requests.length;
    const missingExpected = await tool("onedrive_batch_move", {
      items: [{ itemId: "root-note" }],
      destinationParentItemId: "folder-a",
      dryRun: false,
      confirmed: true
    });
    assert(!missingExpected.isError, "batch_move missing-expected response should be structured", missingExpected);
    assert(missingExpected.value.requiredToMove?.includes("expectedName or expectedId"), "batch_move should require expected identity", missingExpected.value);
    assert(requests.length === beforeMissingExpected, "batch_move should not touch Graph before expected identity", { added: requests.slice(beforeMissingExpected) });
    return { noConfirm: noConfirm.value.requiredToMove, missingExpected: missingExpected.value.requiredToMove };
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

  await check("scan maxItems cap limits metadata cache writes", async () => {
    await tool("onedrive_cache_clear");
    const result = await tool("onedrive_scan", { maxItems: 1, maxFolders: 10, maxResults: 10 });
    assert(!result.isError, "capped scan should succeed", result);
    assert(result.value.summary.itemsScanned === 1, "scan should process one item", result.value.summary);
    const status = await tool("onedrive_sync_status");
    assert(!status.isError, "sync status should succeed", status);
    assert(status.value.itemCount === 1, "cache should contain only processed items", status.value);
    return { itemsScanned: result.value.summary.itemsScanned, cacheItems: status.value.itemCount };
  });

  await check("cache refresh auto does not reuse deltaLink for a different target", async () => {
    await tool("onedrive_cache_clear");
    const first = await tool("onedrive_cache_refresh", { itemId: "folder-a", mode: "scan", replaceCache: true, maxItems: 10, maxFolders: 5, maxDepth: 1 });
    assert(!first.isError, "first cache refresh should succeed", first);
    const before = requests.length;
    const second = await tool("onedrive_cache_refresh", { itemId: "folder-b", mode: "auto", maxItems: 10, maxFolders: 5, maxDepth: 1 });
    assert(!second.isError, "second cache refresh should succeed", second);
    assert(second.value.cache.scanRoot.target === "itemId:folder-b", "cache refresh should switch scan root", second.value.cache);
    const added = requests.slice(before);
    assert(!added.some((request) => request.url.includes("/mock/delta/folder-a")), "cache refresh reused old folder delta", { added });
    return { mode: second.value.mode, scanRoot: second.value.cache.scanRoot.target };
  });

  await check("cache refresh merges by default instead of shrinking cache", async () => {
    await tool("onedrive_cache_clear");
    const broad = await tool("onedrive_cache_refresh", { mode: "scan", replaceCache: true, maxItems: 10, maxFolders: 5, maxDepth: 1 });
    assert(!broad.isError, "broad cache refresh should succeed", broad);
    const before = await tool("onedrive_sync_status");
    const narrow = await tool("onedrive_cache_refresh", { itemId: "folder-a", mode: "scan", maxItems: 1, maxFolders: 1, maxDepth: 0 });
    assert(!narrow.isError, "narrow cache refresh should succeed", narrow);
    const after = await tool("onedrive_sync_status");
    assert(after.value.itemCount >= before.value.itemCount, "default cache refresh should merge instead of shrink", { before: before.value, after: after.value, narrow: narrow.value });
    return { before: before.value.itemCount, after: after.value.itemCount, note: narrow.value.note };
  });

  await check("pathless delta records preserve rename and move hierarchy", async () => {
    await tool("onedrive_cache_clear");
    for (const itemId of ["delta-parent", "delta-archive", "delta-child"]) {
      const seeded = await tool("onedrive_get_info", { itemId });
      assert(!seeded.isError, `failed to seed ${itemId}`, seeded);
    }
    const cacheFile = join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json");
    const cachedPath = () => JSON.parse(readFileSync(cacheFile, "utf8")).itemsById["delta-child"]?.remotePath;

    const unchanged = await tool("onedrive_delta", { itemId: "delta-parent", maxItems: 10 });
    assert(!unchanged.isError, "pathless unchanged delta should succeed", unchanged);
    assert(cachedPath() === "Documents/report.txt", "pathless delta shortened a nested cached path", { cachedPath: cachedPath() });

    const renamed = await tool("onedrive_delta", { itemId: "delta-parent", maxItems: 10 });
    assert(!renamed.isError, "pathless rename delta should succeed", renamed);
    assert(cachedPath() === "Documents/renamed-report.txt", "same-parent pathless rename was not reconciled", { cachedPath: cachedPath() });

    const moved = await tool("onedrive_delta", { itemId: "delta-parent", maxItems: 10 });
    assert(!moved.isError, "pathless move delta should succeed", moved);
    const cache = JSON.parse(readFileSync(cacheFile, "utf8"));
    assert(cachedPath() === "Archive/renamed-report.txt", "pathless move did not resolve through cached parent ID", { cachedPath: cachedPath() });
    assert(!Object.hasOwn(cache.pathsByLower, "documents/renamed-report.txt"), "pathless move left a stale path key", cache.pathsByLower);
    return { path: cachedPath(), parentId: cache.itemsById["delta-child"].parentId };
  });

  await check("root delta resolves top-level and hydrated parent paths while reporting unresolved items", async () => {
    await tool("onedrive_cache_clear");
    rootDeltaScenario = "pathless";
    try {
      const result = await tool("onedrive_delta", { maxItems: 10, format: "full" });
      assert(!result.isError, "pathless root delta should succeed", result);
      const byId = new Map(result.value.items.map((entry) => [entry.id, entry]));
      assert(byId.get("top-pathless")?.remotePath === "top-pathless.txt", "top-level root child path was not resolved", byId.get("top-pathless"));
      assert(byId.get("moved-pathless")?.remotePath === "New Parent/moved-pathless.txt", "uncached parent should be hydrated by ID", byId.get("moved-pathless"));
      assert(result.value.unresolvedPathCount === 1, "failed parent hydration should remain explicitly unresolved", result.value);
      const status = await tool("onedrive_sync_status", { includeSamples: true });
      assert(status.value.unresolvedPathCount === 1 && status.value.unresolvedPathSamples?.[0]?.id === "unresolved-pathless", "sync status should expose unresolved path records", status.value);
      return { top: byId.get("top-pathless").remotePath, moved: byId.get("moved-pathless").remotePath, unresolved: result.value.unresolvedPathCount };
    } finally {
      rootDeltaScenario = "empty";
    }
  });

  await check("delta metadata merges use cTag safely and clear omitted hashes", async () => {
    await tool("onedrive_cache_clear");
    await tool("onedrive_content_index_clear");
    await tool("onedrive_get_info", { itemId: "delta-indexed" });
    const indexed = await tool("onedrive_content_index_refresh", { itemId: "delta-indexed", maxFiles: 1, maxBytesPerFile: 4096 });
    assert(!indexed.isError && indexed.value.indexed === 1, "seed content should be indexed", indexed);

    const renamed = await tool("onedrive_delta", { itemId: "delta-index-parent", maxItems: 10 });
    assert(!renamed.isError, "metadata-only delta rename should succeed", renamed);
    let cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(cache.itemsById["delta-indexed"].remotePath === "Indexed/renamed-indexed-report.txt", "rename should repath cached metadata", cache.itemsById["delta-indexed"]);
    assert(cache.itemsById["delta-indexed"].cTag === "index-c1", "explicit unchanged cTag should be retained", cache.itemsById["delta-indexed"]);
    assert(cache.itemsById["delta-indexed"].file?.hashes === undefined, "changed eTag with omitted hashes must clear stale hashes", cache.itemsById["delta-indexed"]);
    const preserved = await tool("onedrive_content_search", { query: "persistent indexed phrase" });
    assert(preserved.value.items?.[0]?.remotePath === "Indexed/renamed-indexed-report.txt", "unchanged cTag should preserve and repath indexed content", preserved.value);

    const changed = await tool("onedrive_delta", { itemId: "delta-index-parent", maxItems: 10 });
    assert(!changed.isError, "second delta should succeed", changed);
    cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(cache.itemsById["delta-indexed"].cTag === undefined, "omitted cTag on changed eTag must not preserve old content proof", cache.itemsById["delta-indexed"]);
    const invalidated = await tool("onedrive_content_search", { query: "persistent indexed phrase" });
    assert(invalidated.value.items.length === 0, "omitted cTag with changed eTag should invalidate indexed content", invalidated.value);
    return { renamedPath: preserved.value.items[0].remotePath, staleHashCleared: true, omittedCTagInvalidated: true };
  });

  await check("path displacement removes descendants and their indexed content", async () => {
    await tool("onedrive_cache_clear");
    await tool("onedrive_content_index_clear");
    await tool("onedrive_get_info", { itemId: "victim-folder" });
    await tool("onedrive_get_info", { itemId: "victim-child" });
    await tool("onedrive_content_index_refresh", { itemId: "victim-child", maxFiles: 1, maxBytesPerFile: 4096 });
    const replacement = await tool("onedrive_get_info", { itemId: "replacement-collision" });
    assert(!replacement.isError, "replacement path item should cache", replacement);
    const cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(!cache.itemsById["victim-folder"] && !cache.itemsById["victim-child"], "displaced folder descendants should be removed", cache.itemsById);
    const search = await tool("onedrive_content_search", { query: "victim indexed phrase" });
    assert(search.value.items.length === 0, "displaced descendants should be removed from content index", search.value);
    return { replacementId: cache.pathsByLower.collision, descendantsRemoved: true };
  });

  await check("metadata and content caches preserve concurrent same-process and cross-process updates", async () => {
    await tool("onedrive_cache_clear");
    const cacheFile = join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json");
    assert(JSON.parse(readFileSync(cacheFile, "utf8")).version === 3, "cleared cache should use the migration-safe version", {});
    const sameProcess = await Promise.all([
      tool("onedrive_get_info", { itemId: "concurrent-a" }),
      tool("onedrive_get_info", { itemId: "concurrent-b" })
    ]);
    assert(sameProcess.every((entry) => !entry.isError), "same-process concurrent reads should succeed", sameProcess);
    const sameCache = JSON.parse(readFileSync(cacheFile, "utf8"));
    assert(sameCache.itemsById["concurrent-a"] && sameCache.itemsById["concurrent-b"], "same-process cache update was lost", sameCache.itemsById);

    const sharedHome = join(mockHome, "shared-process-home");
    const sharedStorage = join(sharedHome, ".codex", "onedrive-plugin");
    const sharedCacheRoot = join(sharedStorage, "cache");
    rmSync(sharedHome, { recursive: true, force: true });
    const env = {
      ...process.env,
      HOME: sharedHome,
      ONEDRIVE_CLIENT_ID: "mock-client-id",
      ONEDRIVE_TEST_ACCESS_TOKEN: "mock-token",
      ONEDRIVE_GRAPH_BASE_URL: graphBaseUrl,
      ONEDRIVE_IDENTITY_BASE_URL: identityBaseUrl,
      ONEDRIVE_STORAGE_ROOT: sharedStorage,
      ONEDRIVE_CACHE_ROOT: sharedCacheRoot
    };
    const first = createMcpClient(env);
    const second = createMcpClient(env);
    try {
      await Promise.all([
        first.request("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "cache-a", version: "1" } }),
        second.request("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "cache-b", version: "1" } })
      ]);
      await second.tool("onedrive_sync_status");
      await first.tool("onedrive_get_info", { itemId: "concurrent-a" });
      const observed = await second.tool("onedrive_sync_status", { includeSamples: true });
      assert(observed.value.samples?.some((entry) => entry.id === "concurrent-a"), "long-lived second process did not reload the first process update", observed.value);
      await Promise.all([
        first.tool("onedrive_get_info", { itemId: "concurrent-a" }),
        second.tool("onedrive_get_info", { itemId: "concurrent-b" })
      ]);
      const sharedCache = JSON.parse(readFileSync(join(sharedCacheRoot, "metadata-cache.json"), "utf8"));
      assert(sharedCache.itemsById["concurrent-a"] && sharedCache.itemsById["concurrent-b"], "cross-process metadata update was lost", sharedCache.itemsById);

      await first.tool("onedrive_content_index_clear");
      const contentResults = await Promise.all([
        first.tool("onedrive_content_index_refresh", { itemId: "concurrent-a", maxFiles: 1, maxBytesPerFile: 4096 }),
        second.tool("onedrive_content_index_refresh", { itemId: "concurrent-b", maxFiles: 1, maxBytesPerFile: 4096 })
      ]);
      assert(contentResults.every((entry) => !entry.isError), "cross-process content refresh should succeed", contentResults);
      const sharedIndex = JSON.parse(readFileSync(join(sharedCacheRoot, "content-index.json"), "utf8"));
      assert(sharedIndex.entriesById["concurrent-a"] && sharedIndex.entriesById["concurrent-b"], "cross-process content-index update was lost", sharedIndex.entriesById);
      JSON.parse(readFileSync(join(sharedCacheRoot, "metadata-cache.json"), "utf8"));
      JSON.parse(readFileSync(join(sharedCacheRoot, "content-index.json"), "utf8"));
      return { sameProcessItems: 2, crossProcessItems: 2, crossProcessIndexEntries: 2 };
    } finally {
      await Promise.all([first.close(), second.close()]);
    }
  });

  await check("sync status reports metadata cache after scan", async () => {
    const result = await tool("onedrive_sync_status", { includeSamples: true });
    assert(!result.isError, "sync status should succeed", result);
    assert(result.value.itemCount >= 2, "sync status should report cached items", result.value);
    assert(result.value.contentIndex?.enabled === true, "sync status should report content index settings", result.value);
    assert(result.value.samples?.length > 0, "sync status should return samples when requested", result.value);
    return { itemCount: result.value.itemCount, sampleCount: result.value.samples.length };
  });

  await check("content index refresh indexes cached text and content search returns snippets", async () => {
    await tool("onedrive_cache_clear");
    await tool("onedrive_content_index_clear");
    const info = await tool("onedrive_get_info", { itemId: "root-note" });
    assert(!info.isError, "get_info should seed metadata cache for indexing", info);
    const refreshed = await tool("onedrive_content_index_refresh", { maxFiles: 5, maxBytesPerFile: 4096 });
    assert(!refreshed.isError, "content index refresh should succeed", refreshed);
    assert(refreshed.value.indexed === 1, "content index should index the cached root note", refreshed.value);
    const searched = await tool("onedrive_content_search", { query: "mock content", maxResults: 5 });
    assert(!searched.isError, "content search should succeed", searched);
    assert(searched.value.items[0]?.id === "root-note", "content search should return root-note first", searched.value.items);
    assert(searched.value.items[0]?.snippet?.includes("mock content"), "content search should include matched snippet", searched.value.items[0]);
    return { refresh: refreshed.value, hit: searched.value.items[0] };
  });

  await check("content index refresh reuses unchanged entries without reading content again", async () => {
    const before = requests.length;
    const refreshed = await tool("onedrive_content_index_refresh", { maxFiles: 5, maxBytesPerFile: 4096 });
    assert(!refreshed.isError, "second content index refresh should succeed", refreshed);
    assert(refreshed.value.reused === 1, "unchanged indexed file should be reused", refreshed.value);
    const added = requests.slice(before);
    assert(!added.some((request) => request.path.endsWith("/content")), "unchanged refresh should not download content", { added });
    return { reused: refreshed.value.reused, addedRequests: added.length };
  });

  await check("content index refresh prioritizes missing entries under maxFiles cap", async () => {
    const beta = await tool("onedrive_get_info", { itemId: "beta-note" });
    assert(!beta.isError, "get_info should seed a second cache item", beta);
    const refreshed = await tool("onedrive_content_index_refresh", { maxFiles: 1, maxBytesPerFile: 4096 });
    assert(!refreshed.isError, "capped content index refresh should succeed", refreshed);
    assert(refreshed.value.indexed === 1, "missing beta-note index entry should be selected before fresh reusable entries", refreshed.value);
    assert(refreshed.value.reused === 0, "fresh reusable entries should not consume the capped refresh slot first", refreshed.value);
    const searched = await tool("onedrive_content_search", { query: "beta launch", maxResults: 5 });
    assert(!searched.isError, "content search should succeed after capped refresh", searched);
    assert(searched.value.items[0]?.id === "beta-note", "capped refresh should index beta-note", searched.value.items);
    return { indexed: refreshed.value.indexed, reused: refreshed.value.reused, selected: refreshed.value.selected };
  });

  await check("content search matches whole tokens instead of substrings", async () => {
    await tool("onedrive_cache_clear");
    await tool("onedrive_content_index_clear");
    const tibetan = await tool("onedrive_get_info", { itemId: "tibetan-note" });
    const beta = await tool("onedrive_get_info", { itemId: "beta-note" });
    assert(!tibetan.isError && !beta.isError, "test notes should seed metadata", { tibetan, beta });
    const refreshed = await tool("onedrive_content_index_refresh", { maxFiles: 5, maxBytesPerFile: 4096 });
    assert(!refreshed.isError, "content index refresh should succeed", refreshed);
    const searched = await tool("onedrive_content_search", { query: "beta", maxResults: 5 });
    assert(!searched.isError, "content search should succeed", searched);
    const ids = searched.value.items.map((entry) => entry.id);
    assert(ids.includes("beta-note"), "content search should include whole-token beta match", searched.value.items);
    assert(!ids.includes("tibetan-note"), "content search should not match beta inside Tibetan", searched.value.items);
    return { ids };
  });

  await check("content index removes deleted items but survives partial metadata scans", async () => {
    const deleted = await toolWithPreview("onedrive_delete", {
      itemId: "beta-note",
      expectedId: "beta-note",
      dryRun: false,
      confirmed: true
    });
    assert(!deleted.isError, "indexed beta note delete should succeed", deleted);
    const afterDelete = await tool("onedrive_content_search", { query: "beta launch", maxResults: 5 });
    assert(!afterDelete.isError, "content search after delete should succeed", afterDelete);
    assert(!afterDelete.value.items.some((entry) => entry.id === "beta-note"), "deleted item remained in the content index", afterDelete.value.items);

    const reseeded = await tool("onedrive_get_info", { itemId: "beta-note" });
    assert(!reseeded.isError, "mock beta note should be reseedable", reseeded);
    const reindexed = await tool("onedrive_content_index_refresh", { itemId: "beta-note", maxFiles: 1, maxBytesPerFile: 4096 });
    assert(!reindexed.isError && reindexed.value.indexed === 1, "beta note should reindex after reseed", reindexed);
    await tool("onedrive_cache_clear");
    const partial = await tool("onedrive_scan", { maxItems: 1, maxFolders: 1, maxResults: 1 });
    assert(!partial.isError, "partial metadata scan should succeed", partial);
    const afterPartial = await tool("onedrive_content_search", { query: "beta launch", maxResults: 5 });
    assert(afterPartial.value.items.some((entry) => entry.id === "beta-note"), "partial scan incorrectly pruned an unrelated indexed item", afterPartial.value.items);
    return { deletedIndexHits: afterDelete.value.items.length, preservedAfterPartialScan: true };
  });

  await check("find uses content index without fetching file bodies", async () => {
    await tool("onedrive_cache_clear");
    await tool("onedrive_content_index_clear");
    const info = await tool("onedrive_get_info", { itemId: "root-note" });
    assert(!info.isError, "get_info should seed root-note metadata for find", info);
    const refreshed = await tool("onedrive_content_index_refresh", { maxFiles: 5, maxBytesPerFile: 4096 });
    assert(!refreshed.isError, "content index refresh should prepare root-note for find", refreshed);
    const before = requests.length;
    const result = await tool("onedrive_find", {
      query: "mock content",
      maxResults: 3,
      scanFallback: false
    });
    assert(!result.isError, "find should succeed with content index", result);
    assert(result.value.summary.localIndexUsed === true, "find should report content index usage", result.value.summary);
    assert(result.value.items[0]?.id === "root-note", "content-indexed find should return root-note", result.value.items);
    assert(result.value.items[0]?.snippets?.[0]?.includes("mock content"), "find result should include content snippet", result.value.items[0]);
    const added = requests.slice(before);
    assert(!added.some((request) => request.path.endsWith("/content")), "find should not fetch content bodies", { added });
    return { summary: result.value.summary, found: result.value.items[0], graphRequestsAdded: added.length };
  });

  await check("rename dry-run previews without PATCH", async () => {
    const before = requests.length;
    const result = await tool("onedrive_rename", { itemId: "delete-target", newName: "renamed.txt", dryRun: true });
    assert(!result.isError, "rename dry-run should succeed", result);
    assert(result.value.dryRun === true, "rename dry-run should not mutate", result.value);
    const added = requests.slice(before);
    assert(!added.some((request) => request.method === "PATCH"), "rename dry-run should not PATCH", { added });
    return { graphRequestsAdded: added.length };
  });

  await check("rename updates cache path keys", async () => {
    await tool("onedrive_cache_clear");
    const info = await tool("onedrive_get_info", { itemId: "delete-target" });
    assert(!info.isError, "get_info should seed cache", info);
    const result = await tool("onedrive_rename", {
      itemId: "delete-target",
      newName: "renamed-cache.txt",
      expectedName: "delete-me.txt",
      dryRun: false,
      confirmed: true
    });
    assert(!result.isError, "rename should succeed", result);
    const cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(!Object.hasOwn(cache.pathsByLower, "delete-me.txt"), "old cache path key should be removed", cache.pathsByLower);
    assert(cache.pathsByLower["renamed-cache.txt"] === "delete-target", "new cache path key should be present", cache.pathsByLower);
    return { paths: cache.pathsByLower };
  });

  await check("folder rename repaths cached descendants", async () => {
    await tool("onedrive_cache_clear");
    const scanResult = await tool("onedrive_scan", { itemId: "folder-a", maxItems: 10, maxFolders: 5, maxDepth: 1 });
    assert(!scanResult.isError, "scan should seed folder descendants", scanResult);
    let cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(cache.pathsByLower["folder a/deep summary deck.pptx"] === "deep-deck", "expected descendant cache path before rename", cache.pathsByLower);

    const renamed = await tool("onedrive_rename", {
      itemId: "folder-a",
      newName: "Folder Renamed",
      expectedName: "Folder A",
      dryRun: false,
      confirmed: true
    });
    assert(!renamed.isError, "folder rename should succeed", renamed);
    cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(!Object.hasOwn(cache.pathsByLower, "folder a/deep summary deck.pptx"), "folder rename should remove stale descendant cache path", cache.pathsByLower);
    assert(cache.pathsByLower["folder renamed"] === "folder-a", "folder rename should cache the renamed folder path", cache.pathsByLower);
    assert(cache.pathsByLower["folder renamed/deep summary deck.pptx"] === "deep-deck", "folder rename should preserve descendants at their new path", cache.pathsByLower);
    return { paths: cache.pathsByLower };
  });

  await check("batch_move updates metadata cache path keys", async () => {
    await tool("onedrive_cache_clear");
    const info = await tool("onedrive_get_info", { itemId: "root-note" });
    assert(!info.isError, "get_info should seed cache before batch move", info);
    const moved = await tool("onedrive_batch_move", {
      items: [{ itemId: "root-note", expectedId: "root-note" }],
      destinationParentItemId: "folder-a",
      dryRun: false,
      confirmed: true
    });
    assert(!moved.isError, "batch_move should succeed", moved);
    const cache = JSON.parse(readFileSync(join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json"), "utf8"));
    assert(!Object.hasOwn(cache.pathsByLower, "root-note.txt"), "batch_move should remove stale source cache path", cache.pathsByLower);
    assert(cache.pathsByLower["folder a/root-note.txt"] === "root-note", "batch_move should cache destination path", cache.pathsByLower);
    const content = await tool("onedrive_content_search", { query: "mock content", maxResults: 5 });
    assert(!content.isError, "content search after move should succeed", content);
    assert(content.value.items.find((entry) => entry.id === "root-note")?.remotePath === "Folder A/root-note.txt", "content index metadata did not follow the moved file", content.value.items);
    return { paths: cache.pathsByLower, indexedPath: content.value.items.find((entry) => entry.id === "root-note")?.remotePath };
  });

  await check("large_files evaluates scanned files beyond returned scan result cap", async () => {
    const result = await tool("onedrive_large_files", {
      itemId: "bulk-large",
      minBytes: 1000,
      maxItems: 6000,
      maxFolders: 2,
      limit: 5
    });
    assert(!result.isError, "large_files should succeed", result);
    assert(result.value.items.some((entry) => entry.id === "huge-after-cap"), "large_files should include huge file after first 5000 scan matches", result.value);
    return { count: result.value.count, ids: result.value.items.map((entry) => entry.id), scanned: result.value.scanSummary.itemsScanned };
  });

  await check("duplicates evaluates scanned files beyond returned scan result cap", async () => {
    const result = await tool("onedrive_duplicates", {
      itemId: "bulk-dupes",
      maxItems: 6005,
      maxFolders: 2,
      limit: 5
    });
    assert(!result.isError, "duplicates should succeed", result);
    const group = result.value.groups.find((entry) => entry.items.some((item) => item.id === "dup-a"));
    assert(group?.items.some((item) => item.id === "dup-b"), "duplicates should include pair after first 5000 scan matches", result.value);
    return { duplicateGroups: result.value.duplicateGroups, group };
  });

  await check("find retains canonical Graph content hits but rejects unrelated expansion hits", async () => {
    const before = requests.length;
    const result = await tool("onedrive_find", {
      query: "canonical hidden phrase",
      useCache: false,
      useContentIndex: false,
      scanFallback: false,
      maxSearchTerms: 2,
      searchConcurrency: 2,
      maxResults: 5
    });
    assert(!result.isError, "canonical content-hit find should succeed", result);
    assert(result.value.items.some((entry) => entry.id === "canonical-content-hit"), "canonical Graph content/metadata hit was dropped", result.value);
    assert(!result.value.items.some((entry) => entry.id === "unrelated-expansion-hit"), "unrelated expansion-only hit should remain gated", result.value);
    assert(result.value.items.find((entry) => entry.id === "canonical-content-hit")?.reasons?.includes("canonical Graph content/metadata match"), "canonical result should explain why it was retained", result.value.items);
    assert(result.value.summary.searchTermsExecuted === 2, "expected canonical plus one expansion term", result.value.summary);
    assert(result.value.summary.searchTermsSkipped === 0, "low-confidence canonical content hit should not skip the planned expansion", result.value.summary);
    assert(result.value.note.includes("Recursive scan fallback was disabled."), "find note should truthfully report disabled scan fallback", result.value.note);
    assert(!result.value.note.includes("metadata cache") && !result.value.note.includes("content index"), "find note should not claim disabled local sources were used", result.value.note);
    const searchRequests = requests.slice(before).filter((request) => decodeURIComponent(request.url).includes("/search(q='"));
    assert(searchRequests.length === 2, "find should execute exactly the instrumented search terms", { searchRequests, searchPlan: result.value.searchPlan });
    return {
      ids: result.value.items.map((entry) => entry.id),
      summary: result.value.summary,
      searchPlan: result.value.searchPlan
    };
  });

  await check("find reports actual paginated Graph search calls", async () => {
    const result = await tool("onedrive_find", {
      query: "Paged Research",
      maxSearchTerms: 1,
      scanFallback: false,
      useCache: false,
      useContentIndex: false
    });
    assert(!result.isError, "paged find should succeed", result);
    assert(result.value.summary.graphSearchCalls === 2, "graphSearchCalls should count both fetched pages", result.value.summary);
    assert(result.value.items.length === 2, "both paginated hits should be returned", result.value.items);
    return { graphSearchCalls: result.value.summary.graphSearchCalls, ids: result.value.items.map((entry) => entry.id) };
  });

  await check("find_all executes every planned term and preserves earlier specific folder hints", async () => {
    const result = await tool("onedrive_find_all", {
      query: "Quarterly Research",
      folderHints: ["Personal/Documents/Health"],
      maxSearchTerms: 3,
      maxResults: 10,
      useCache: false,
      useContentIndex: false
    });
    assert(!result.isError, "broad find should succeed", result);
    assert(result.value.searchPlan.stopReason === "all-terms-executed", "find_all should not confidence-stop", result.value.searchPlan);
    assert(result.value.searchPlan.executed.length === result.value.searchPlan.planned.length, "find_all skipped planned terms", result.value.searchPlan);
    assert(result.value.items.some((entry) => entry.id === "quarterly-research-broad"), "expansion-only relevant hit was omitted", result.value.items);
    assert(result.value.folderPlan.includes("Personal/Documents/Health") && !result.value.folderPlan.includes("Personal/Documents"), "later default ancestor replaced the explicit specific hint", result.value.folderPlan);
    return { plan: result.value.searchPlan, folderPlan: result.value.folderPlan, ids: result.value.items.map((entry) => entry.id) };
  });

  await check("find batches live search cache persistence and useCache false performs no metadata I/O", async () => {
    await tool("onedrive_cache_clear");
    const batched = await tool("onedrive_find", {
      query: "Batch Cache Research",
      maxSearchTerms: 2,
      minConfidenceForSearchOnly: 100,
      scanFallback: false,
      useCache: true,
      useContentIndex: false
    });
    assert(!batched.isError, "batched cache find should succeed", batched);
    assert(batched.value.summary.searchTermsExecuted === 2, "test should execute multiple search terms", batched.value.summary);
    assert(batched.value.summary.metadataCacheWrites === 1, "accepted live search results should persist in one cache write", batched.value.summary);

    const cacheFile = join(mockHome, ".codex", "onedrive-plugin", "cache", "metadata-cache.json");
    const validCache = readFileSync(cacheFile, "utf8");
    writeFileSync(cacheFile, "{invalid-cache-json");
    const liveOnly = await tool("onedrive_find", {
      query: "No Cache Live",
      maxSearchTerms: 1,
      scanFallback: false,
      useCache: false,
      useContentIndex: false
    });
    const untouched = readFileSync(cacheFile, "utf8");
    writeFileSync(cacheFile, validCache);
    assert(!liveOnly.isError, "fully live find should ignore an unreadable metadata cache", liveOnly);
    assert(liveOnly.value.summary.metadataCacheWrites === 0, "useCache false should not write metadata", liveOnly.value.summary);
    assert(untouched === "{invalid-cache-json", "useCache false modified the metadata cache", { untouched });
    assert(!liveOnly.value.localWarnings?.some((warning) => warning.operation.includes("metadata cache")), "useCache false read the invalid metadata cache", liveOnly.value.localWarnings);
    return { batchedWrites: batched.value.summary.metadataCacheWrites, liveOnlyWrites: liveOnly.value.summary.metadataCacheWrites };
  });

  await check("find adaptively skips expansion terms after a confident canonical hit", async () => {
    const before = requests.length;
    const result = await tool("onedrive_find", {
      query: "Adaptive Exact Match",
      useCache: false,
      useContentIndex: false,
      scanFallback: false,
      maxSearchTerms: 8,
      searchConcurrency: 4,
      maxResults: 5
    });
    assert(!result.isError, "adaptive confident find should succeed", result);
    assert(result.value.items[0]?.id === "adaptive-exact-hit", "confident canonical result should rank first", result.value.items);
    assert(result.value.summary.graphSearchCalls === 1, "confident canonical result should require one Graph search", result.value.summary);
    assert(result.value.summary.searchTermsExecuted === 1, "only the canonical term should execute", result.value.summary);
    assert(result.value.summary.searchTermsSkipped > 0, "expansion terms should be recorded as skipped", result.value.summary);
    assert(result.value.searchPlan.stopReason === "high-confidence-canonical", "unexpected adaptive stop reason", result.value.searchPlan);
    const searchRequests = requests.slice(before).filter((request) => decodeURIComponent(request.url).includes("/search(q='"));
    assert(searchRequests.length === 1, "adaptive stop should make exactly one Graph search request", { searchRequests, searchPlan: result.value.searchPlan });
    return {
      found: result.value.items[0],
      summary: result.value.summary,
      searchPlan: result.value.searchPlan
    };
  });

  await check("find ignores unrelated cache-only no-match results", async () => {
    const result = await tool("onedrive_find", {
      query: "qwertyuiopasdf",
      maxResults: 3,
      scanFallback: false
    });
    assert(!result.isError, "no-match find should succeed", result);
    assert(result.value.items.length === 0, "no-match find should not return unrelated cached items", result.value);
    return { summary: result.value.summary };
  });

  await check("find root fallback does not rescan completed hinted subtrees", async () => {
    const before = requests.length;
    const result = await tool("onedrive_find", {
      query: "qwertyuiopasdf",
      folderHints: ["Folder A", ""],
      maxSearchTerms: 1,
      maxResults: 3,
      scanConcurrency: 1,
      scanMaxItems: 50,
      scanMaxFolders: 20,
      useCache: false,
      useContentIndex: false
    });
    assert(!result.isError, "find subtree-pruning fallback should succeed", result);
    const added = requests.slice(before);
    const folderAChildrenRequests = added.filter((request) => request.url.includes("/items/folder-a/children"));
    assert(folderAChildrenRequests.length === 1, "root fallback rescanned the completed Folder A subtree", {
      folderAChildrenRequests,
      scanRuns: result.value.scanRuns
    });
    const rootRun = result.value.scanRuns.find((run) => run.folder === "root" && run.summary);
    assert(rootRun?.summary?.foldersSkipped >= 1, "root fallback did not report the completed subtree exclusion", result.value.scanRuns);
    return {
      folderAChildrenRequests: folderAChildrenRequests.length,
      rootFoldersSkipped: rootRun.summary.foldersSkipped,
      scanAttempts: result.value.summary.scanAttempts
    };
  });

  await check("find drops stale cache-only hits when live scan cannot confirm", async () => {
    const seeded = await tool("onedrive_get_info", { itemId: "stale-cache" });
    assert(!seeded.isError, "get_info should seed stale cache item", seeded);
    const result = await tool("onedrive_find", {
      query: "Stale Cache Deck",
      maxResults: 3,
      scanMaxItems: 20,
      scanMaxFolders: 10
    });
    assert(!result.isError, "stale-cache find should succeed", result);
    assert(result.value.summary.usedScanFallback === true, "stale cache hit should not suppress scan", result.value.summary);
    assert(!result.value.items.some((item) => item.id === "stale-cache"), "stale cache-only item should not be returned", result.value.items);
    return { summary: result.value.summary, ids: result.value.items.map((item) => item.id) };
  });

  await check("find confirms cached nested deck with live scan", async () => {
    const result = await tool("onedrive_find", {
      query: "Deep Summary Deck",
      maxResults: 3,
      scanMaxItems: 20,
      scanMaxFolders: 10
    });
    assert(!result.isError, "find should succeed", result);
    assert(result.value.summary.localIndexUsed === false, "find must not use a local index", result.value.summary);
    assert(result.value.summary.persistentCacheUsed === true, "find should use persistent cache after earlier scans populate it", result.value.summary);
    assert(result.value.summary.usedScanFallback === true, "find should confirm cache-only hit with scan fallback", result.value.summary);
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
    assert(result.value.summary.persistentCacheUsed === true, "report find should use cache when available", result.value.summary);
    assert(result.value.summary.usedScanFallback === true, "report find should confirm cache-only hit with scan fallback", result.value.summary);
    assert(result.value.items[0]?.id === "quarterly-report", "report find should return the docx report", result.value.items);
    assert(!result.value.inferred?.strictExtensions?.includes(".pdf"), "report should not infer a strict PDF filter", result.value.inferred);
    return {
      summary: result.value.summary,
      found: result.value.items[0],
      inferred: result.value.inferred
    };
  });

  await check("find_all returns broader ranked results with cache acceleration", async () => {
    const result = await tool("onedrive_find_all", {
      query: "Quarterly Report",
      maxResults: 20,
      scanMaxItems: 20,
      scanMaxFolders: 10
    });
    assert(!result.isError, "find_all should succeed", result);
    assert(result.value.strategy === "broad-cache-assisted-remote-first", "unexpected strategy", result.value);
    assert(result.value.summary.localIndexUsed === false, "find_all must not use a local index", result.value.summary);
    assert(result.value.summary.persistentCacheUsed === true, "find_all should use persistent cache when available", result.value.summary);
    assert(!result.value.note.includes("persistent cache was used"), "find_all note should not contradict cache use", result.value);
    assert(result.value.items.some((item) => item.id === "quarterly-report"), "find_all should include docx report", result.value.items);
    assert(result.value.folderPlan?.includes("root"), "find_all should include root in folder plan", result.value.folderPlan);
    assert(result.value.folderPlan?.at(-1) === "root", "root should remain the final broad-search fallback", result.value.folderPlan);
    const nonRootFolderKeys = result.value.folderPlan
      .filter((folder) => folder !== "root")
      .map((folder) => String(folder).replace(/^\/+|\/+$/g, "").toLowerCase());
    assert(!nonRootFolderKeys.some((folder, index) => nonRootFolderKeys.some((possibleParent, parentIndex) =>
      index !== parentIndex && possibleParent && folder.startsWith(`${possibleParent}/`)
    )), "find_all folder plan should prune nested hints regardless of input order", result.value.folderPlan);
    return {
      summary: result.value.summary,
      folderPlan: result.value.folderPlan,
      names: result.value.items.map((item) => item.name)
    };
  });

  await check("sharing audits ignore owner-only private permissions", async () => {
    const shared = await tool("onedrive_shared_by_me", {
      maxItems: 20,
      maxFolders: 10,
      maxDepth: 2,
      limit: 10
    });
    assert(!shared.isError, "shared_by_me should succeed", shared);
    const sharedIds = shared.value.items.map((entry) => entry.item.id);
    assert(sharedIds.includes("deep-deck"), "shared_by_me should include explicitly shared deck", shared.value);
    assert(sharedIds.includes("deep-pdf"), "shared_by_me should include a direct grant represented only by grantedToIdentitiesV2", shared.value);
    assert(!sharedIds.includes("root-note"), "shared_by_me should exclude owner-only private files", shared.value);
    assert(shared.value.auditedCount > 0 && shared.value.errorCount === 0 && shared.value.incomplete === false, "successful sharing audit should report complete counters", shared.value);

    const publicLinks = await tool("onedrive_public_links", {
      maxItems: 20,
      maxFolders: 10,
      maxDepth: 2,
      limit: 10
    });
    assert(!publicLinks.isError, "public_links should succeed", publicLinks);
    const publicIds = publicLinks.value.items.map((entry) => entry.item.id);
    assert(publicIds.includes("deep-deck"), "public_links should include anonymous link", publicLinks.value);
    assert(!publicIds.includes("deep-pdf"), "public_links should exclude direct V2 grants without anonymous links", publicLinks.value);
    assert(!publicIds.includes("root-note"), "public_links should exclude owner-only private files", publicLinks.value);
    return { sharedIds, publicIds };
  });

  await check("sharing audits surface exhausted permission failures", async () => {
    const before = counters.get("batch-permission-error") || 0;
    const result = await tool("onedrive_shared_by_me", {
      itemId: "sharing-error-root",
      maxItems: 5,
      maxFolders: 2,
      maxDepth: 1,
      limit: 5
    });
    assert(!result.isError, "sharing audit should return structured partial results", result);
    assert(result.value.auditedCount === 0 && result.value.errorCount === 1, "sharing audit should count failed permission reads", result.value);
    assert(result.value.incomplete === true && result.value.errors?.length === 1, "sharing audit should mark permission failures incomplete", result.value);
    assert((counters.get("batch-permission-error") || 0) === before + 4, "transient permission subrequest should stop after bounded retries", {
      before,
      after: counters.get("batch-permission-error") || 0
    });
    return {
      auditedCount: result.value.auditedCount,
      errorCount: result.value.errorCount,
      incomplete: result.value.incomplete
    };
  });

  await check("audit log records live successes and safe failures", async () => {
    const failed = await toolWithPreview("onedrive_delete", {
      itemId: "delete-fail",
      expectedName: "delete-fail.txt",
      dryRun: false,
      confirmed: true
    });
    assert(failed.isError, "failed live delete should return tool error", failed);
    const recent = await tool("onedrive_audit_recent", { limit: 20 });
    assert(!recent.isError, "audit recent should succeed", recent);
    const entries = recent.value.entries || [];
    const tools = entries.map((entry) => entry.tool);
    assert(tools.includes("onedrive_delete"), "audit should include live delete", entries);
    assert(tools.includes("onedrive_create_sharing_link"), "audit should include live sharing link", entries);
    assert(tools.includes("onedrive_invite_permission"), "audit should include live invite", entries);
    assert(tools.includes("onedrive_revoke_permission"), "audit should include live revoke", entries);
    assert(entries.some((entry) => entry.tool === "onedrive_delete" && entry.status === "failed" && entry.error?.message?.includes("mock delete failure")), "audit should include safe failed delete info", entries);
    const scoped = await tool("onedrive_audit_recent", { tool: "onedrive_delete", status: "failed", pathContains: "delete-fail", limit: 5 });
    assert(!scoped.isError, "scoped audit recent should succeed", scoped);
    assert(scoped.value.entries.length >= 1, "scoped audit should include delete-fail entry", scoped.value);
    assert(scoped.value.entries.every((entry) => entry.tool === "onedrive_delete" && entry.status === "failed"), "scoped audit should filter by tool and status", scoped.value);
    const serialized = JSON.stringify(entries);
    assert(!serialized.includes("mock-token"), "audit should not include access token", entries);
    assert(!serialized.includes("https://example.test/share/link"), "audit should not include sharing webUrl", entries);
    assert(!serialized.includes("link-secret"), "audit should not include sharing link password", entries);
    assert(!serialized.includes("invite-secret"), "audit should not include invite password", entries);
    assert(!serialized.includes("Please review"), "audit should not include invitation message", entries);
    assert(!serialized.includes("person@example.test"), "audit should not include recipient email", entries);
    assert(!serialized.includes("11111111-1111-1111-1111-111111111111"), "audit should not include object identifiers echoed in errors", entries);
    assert(!serialized.includes("https://example.test/invite"), "audit should not include URLs echoed in errors", entries);
    assert(!serialized.includes("abc.def.ghi"), "audit should not include bearer-looking tokens echoed in errors", entries);
    return { count: entries.length, tools };
  });

  await check("plugin-managed OneDrive data is private on disk", async () => {
    const storage = join(mockHome, ".codex", "onedrive-plugin");
    const paths = {
      storage,
      cache: join(storage, "cache"),
      audit: join(storage, "audit"),
      metadata: join(storage, "cache", "metadata-cache.json"),
      contentIndex: join(storage, "cache", "content-index.json"),
      auditLog: join(storage, "audit", "mutations.jsonl"),
      exportedPdf: join(pluginRoot, "work", "mock-export.pdf")
    };
    const mode = (path) => statSync(path).mode & 0o777;
    for (const path of [paths.storage, paths.cache, paths.audit]) {
      assert(mode(path) === 0o700, `private directory mode should be 0700: ${path}`, { path, mode: mode(path).toString(8) });
    }
    for (const path of [paths.metadata, paths.contentIndex, paths.auditLog, paths.exportedPdf]) {
      assert(mode(path) === 0o600, `private file mode should be 0600: ${path}`, { path, mode: mode(path).toString(8) });
    }
    return Object.fromEntries(Object.entries(paths).map(([name, path]) => [name, mode(path).toString(8)]));
  });

  await check("successful remote mutation reports local bookkeeping failure as a warning", async () => {
    const auditDirectory = join(mockHome, ".codex", "onedrive-plugin", "audit");
    const auditLogPath = join(auditDirectory, "mutations.jsonl");
    const previousAuditLog = readFileSync(auditLogPath);
    rmSync(auditDirectory, { recursive: true, force: true });
    writeFileSync(auditDirectory, "blocks audit directory creation", "utf8");
    try {
      const result = await tool("onedrive_create_folder", { name: "Bookkeeping Warning Folder" });
      assert(!result.isError, "remote create_folder success must not be reported as failed when local audit persistence fails", result);
      assert(result.value.id === "created-child-folder", "create_folder should return the remote result", result.value);
      assert(result.value.localWarnings?.some((warning) => warning.operation === "mutation audit write"), "local audit failure should be disclosed", result.value);
      return { itemId: result.value.id, localWarnings: result.value.localWarnings };
    } finally {
      rmSync(auditDirectory, { force: true });
      mkdirSync(auditDirectory, { recursive: true, mode: 0o700 });
      writeFileSync(auditLogPath, previousAuditLog, { mode: 0o600 });
    }
  });

  await check("audit_clear requires explicit confirmation", async () => {
    const noConfirm = await tool("onedrive_audit_clear");
    assert(!noConfirm.isError, "audit_clear no-confirm should be structured", noConfirm);
    assert(noConfirm.value.requiredToClear, "audit_clear should require confirmation", noConfirm.value);
    assert(auditEntries().length > 0, "audit log should remain before confirmed clear");
    const cleared = await tool("onedrive_audit_clear", { confirmed: true });
    assert(!cleared.isError, "confirmed audit_clear should succeed", cleared);
    assert(auditEntries().length === 0, "audit log should be removed after confirmed clear", auditEntries());
    return { cleared: cleared.value.cleared };
  });
} finally {
  child.stdin.end();
  child.kill("SIGTERM");
  graph.close();
  identity.close();
}

const failCount = results.filter((result) => result.status === "fail").length;
console.log(JSON.stringify({ graphBaseUrl, results, stderr: stderr.join(""), summary: { total: results.length, failCount } }, null, 2));
if (failCount === 0 && !keepWork) {
  rmSync(mockHome, { recursive: true, force: true });
  rmSync(join(pluginRoot, "work", "mock-export.pdf"), { force: true });
  rmSync(join(pluginRoot, "work", "mock-export.txt"), { force: true });
}
if (failCount > 0) process.exitCode = 1;
