#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { createServer } from "node:https";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const host = "127.0.0.1";
const port = 3443;
const pluginRoot = join(dirname(fileURLToPath(import.meta.url)), "..");
const allowedFiles = new Map([
  ["/office-addin/taskpane.html", { path: join(pluginRoot, "office-addin", "taskpane.html"), contentType: "text/html; charset=utf-8" }],
  ["/office-addin/taskpane.js", { path: join(pluginRoot, "office-addin", "taskpane.js"), contentType: "text/javascript; charset=utf-8" }],
  ["/office-addin/icon-16.png", { path: join(pluginRoot, "office-addin", "icon-16.png"), contentType: "image/png" }],
  ["/office-addin/icon-32.png", { path: join(pluginRoot, "office-addin", "icon-32.png"), contentType: "image/png" }],
  ["/office-addin/icon-64.png", { path: join(pluginRoot, "office-addin", "icon-64.png"), contentType: "image/png" }],
  ["/office-addin/icon-80.png", { path: join(pluginRoot, "office-addin", "icon-80.png"), contentType: "image/png" }]
]);

function writeResponse(request, response, pathname, status, headers = {}, body) {
  const userAgent = String(request.headers["user-agent"] || "")
    .replace(/[\r\n\u0000-\u001f\u007f]/g, " ")
    .slice(0, 256);
  console.log(JSON.stringify({
    event: "office-companion-request",
    method: String(request.method || "UNKNOWN").slice(0, 16),
    path: pathname,
    status,
    userAgent
  }));
  response.writeHead(status, headers).end(request.method === "HEAD" ? undefined : body);
}

function option(name) {
  const equalsPrefix = `--${name}=`;
  const equalsArgument = process.argv.slice(2).find((argument) => argument.startsWith(equalsPrefix));
  if (equalsArgument) return equalsArgument.slice(equalsPrefix.length);
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

if (process.argv.includes("--help")) {
  console.log("Usage: node office-addin/serve.mjs --cert <certificate.pem> --key <private-key.pem>");
  console.log(`Serves only the Office companion assets at https://${host}:${port}.`);
  process.exit(0);
}

const certPath = option("cert");
const keyPath = option("key");
if (!certPath || !keyPath) {
  console.error("Both --cert and --key are required. See office-addin/README.md for temporary certificate setup and cleanup.");
  process.exit(2);
}

let credentials;
try {
  credentials = { cert: readFileSync(certPath), key: readFileSync(keyPath) };
} catch (error) {
  console.error(`Could not read the HTTPS certificate or key: ${error.message}`);
  process.exit(2);
}

const server = createServer({ ...credentials, maxHeaderSize: 16 * 1024 }, (request, response) => {
  let pathname;
  try {
    pathname = new URL(request.url, `https://${host}:${port}`).pathname;
  } catch {
    writeResponse(request, response, "<invalid>", 400, {}, "Bad request.\n");
    return;
  }
  if (pathname === "/") {
    writeResponse(request, response, pathname, 302, { Location: "/office-addin/taskpane.html", "Cache-Control": "no-store" });
    return;
  }
  if (!allowedFiles.has(pathname)) {
    writeResponse(request, response, pathname, 404, { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" }, "Not found.\n");
    return;
  }
  if (request.method !== "GET" && request.method !== "HEAD") {
    writeResponse(request, response, pathname, 405, { Allow: "GET, HEAD", "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" }, "Method not allowed.\n");
    return;
  }
  const asset = allowedFiles.get(pathname);
  try {
    const body = readFileSync(asset.path);
    writeResponse(request, response, pathname, 200, {
      "Content-Type": asset.contentType,
      "Content-Length": body.length,
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
      "Referrer-Policy": "no-referrer"
    }, body);
  } catch {
    writeResponse(request, response, pathname, 500, { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" }, "Asset unavailable.\n");
  }
});

server.on("clientError", (_error, socket) => socket.end("HTTP/1.1 400 Bad Request\r\nConnection: close\r\n\r\n"));
server.on("error", (error) => {
  console.error(`Office companion HTTPS server failed: ${error.message}`);
  process.exitCode = 1;
});
server.listen(port, host, () => console.log(`Office companion host ready at https://${host}:${port}/office-addin/taskpane.html (pid ${process.pid}).`));

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => server.close(() => process.exit(0)));
}
