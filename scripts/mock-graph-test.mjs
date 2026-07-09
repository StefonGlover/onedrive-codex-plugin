#!/usr/bin/env node

import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");
const mockHome = join(pluginRoot, "work", "mock-home");
const keepWork = process.argv.includes("--keep-work");
rmSync(mockHome, { recursive: true, force: true });
mkdirSync(mockHome, { recursive: true });

const requests = [];
const authRequests = [];
const graphBodies = [];
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

  if (req.method === "POST" && decodedUrl.includes("/v1.0/me/drive/root:/Folder A:/children")) {
    const body = await readJsonBody(req);
    return json(res, 201, folder("created-child-folder", body.name || "Created Child", {
      parentReference: { path: "/drives/drive/root:/Folder A" }
    }));
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

const identity = createServer((req, res) => {
  const url = new URL(req.url, "http://127.0.0.1");
  authRequests.push({ method: req.method, path: url.pathname });

  if (req.method === "POST" && url.pathname === "/common/oauth2/v2.0/devicecode") {
    return json(res, 400, {
      error: "invalid_request",
      error_description: "AADSTS9002331: Application is configured for use by Microsoft Account users only. Please use the /consumers endpoint to serve this request."
    });
  }

  if (req.method === "POST" && url.pathname === "/consumers/oauth2/v2.0/devicecode") {
    return json(res, 200, {
      user_code: "MOCK-CODE",
      device_code: "mock-device-code",
      verification_uri: "https://microsoft.com/devicelogin",
      verification_uri_complete: "https://microsoft.com/devicelogin?user_code=MOCK-CODE",
      expires_in: 900,
      interval: 5,
      message: "Use code MOCK-CODE"
    });
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

  await check("MSA-only common tenant auth retries device login on consumers endpoint", async () => {
    const result = await tool("onedrive_auth_device_start", { tenant: "common" });
    assert(!result.isError, "device-code login should retry on consumers", result);
    assert(result.value.authTenant === "consumers", "device-code login should report consumers fallback", result.value);
    assert(result.value.userCode === "MOCK-CODE", "device-code login should return mocked consumers response", result.value);
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
    return { listUrl: listRequest.url, searchUrl: searchRequest.url };
  });

  await check("list_all clears truncation after the final page", async () => {
    const result = await tool("onedrive_list_all", { itemId: "root", pageSize: 2, maxItems: 10 });
    assert(!result.isError, "list_all should succeed across paginated root children", result);
    assert(result.value.count === 3, "list_all should collect both mock pages", result.value);
    assert(result.value.nextLink === null, "list_all should not retain a consumed nextLink", result.value);
    assert(result.value.truncated === false, "list_all should not report truncation after reaching the final page", result.value);
    return { count: result.value.count, truncated: result.value.truncated };
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
    assert(result.value.permissionDiff?.afterCount === 2, "sharing diff should track after count", result.value.permissionDiff);
    return result.value.permissionDiff;
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

  await check("sync status reports metadata cache after scan", async () => {
    const result = await tool("onedrive_sync_status", { includeSamples: true });
    assert(!result.isError, "sync status should succeed", result);
    assert(result.value.itemCount >= 3, "sync status should report cached items", result.value);
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

  await check("folder rename invalidates cached descendant paths", async () => {
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
    return { paths: cache.pathsByLower };
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
    assert(!sharedIds.includes("root-note"), "shared_by_me should exclude owner-only private files", shared.value);

    const publicLinks = await tool("onedrive_public_links", {
      maxItems: 20,
      maxFolders: 10,
      maxDepth: 2,
      limit: 10
    });
    assert(!publicLinks.isError, "public_links should succeed", publicLinks);
    const publicIds = publicLinks.value.items.map((entry) => entry.item.id);
    assert(publicIds.includes("deep-deck"), "public_links should include anonymous link", publicLinks.value);
    assert(!publicIds.includes("root-note"), "public_links should exclude owner-only private files", publicLinks.value);
    return { sharedIds, publicIds };
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
if (failCount === 0) rmSync(mockHome, { recursive: true, force: true });
if (failCount > 0) process.exitCode = 1;
