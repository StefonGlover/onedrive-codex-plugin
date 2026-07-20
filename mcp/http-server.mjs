#!/usr/bin/env node

import { createServer } from "node:http";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import {
  authorizeMcpRequest,
  oauthSettings,
  protectedResourceMetadata,
  validateOAuthConfiguration
} from "./oauth.mjs";
import { processMcpMessage, shutdownOneDriveServer } from "./server.mjs";

const maxRequestBytes = 1024 * 1024;

function listenAddress(env = process.env) {
  const host = String(env.ONEDRIVE_MCP_HTTP_HOST || "127.0.0.1").trim();
  const port = Number(env.ONEDRIVE_MCP_HTTP_PORT || 3001);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("ONEDRIVE_MCP_HTTP_PORT must be an integer from 1 to 65535.");
  }
  return { host, port };
}

function setCommonHeaders(response) {
  response.setHeader("Cache-Control", "no-store");
  response.setHeader("X-Content-Type-Options", "nosniff");
  response.setHeader("Access-Control-Allow-Origin", "https://chatgpt.com");
  response.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  response.setHeader(
    "Access-Control-Allow-Headers",
    "Authorization, Content-Type, Accept, MCP-Protocol-Version, MCP-Session-Id, Last-Event-ID"
  );
  response.setHeader("Access-Control-Expose-Headers", "MCP-Session-Id, WWW-Authenticate");
}

function sendJson(response, status, payload, extraHeaders = {}) {
  setCommonHeaders(response);
  for (const [name, value] of Object.entries(extraHeaders)) response.setHeader(name, value);
  const body = `${JSON.stringify(payload)}\n`;
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  response.end(body);
}

function sendText(response, status, body, extraHeaders = {}) {
  setCommonHeaders(response);
  for (const [name, value] of Object.entries(extraHeaders)) response.setHeader(name, value);
  response.writeHead(status, {
    "Content-Type": "text/plain; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  response.end(body);
}

async function readJsonBody(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxRequestBytes) {
      const error = new Error(`MCP request body exceeds ${maxRequestBytes} bytes.`);
      error.status = 413;
      throw error;
    }
    chunks.push(chunk);
  }
  if (!chunks.length) {
    const error = new Error("MCP request body is empty.");
    error.status = 400;
    throw error;
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch (error) {
    const invalid = new Error(`Invalid JSON request body: ${error.message}`);
    invalid.status = 400;
    throw invalid;
  }
}

function isToolCall(message) {
  return message?.method === "tools/call";
}

async function requestAuthorization(request, messages) {
  const settings = oauthSettings();
  if (settings.mode !== "oauth") return { authMode: "noauth" };
  if (!messages.some(isToolCall)) return null;
  try {
    return await authorizeMcpRequest(request.headers.authorization, { requireGraph: true });
  } catch (error) {
    return { authMode: "oauth_error", error };
  }
}

async function handleMcp(request, response) {
  if (request.method !== "POST") {
    sendText(response, 405, "Method Not Allowed\n", { Allow: "POST, OPTIONS" });
    return;
  }
  const payload = await readJsonBody(request);
  const messages = Array.isArray(payload) ? payload : [payload];
  if (!messages.length || messages.some((message) => !message || typeof message !== "object" || message.jsonrpc !== "2.0")) {
    sendJson(response, 400, {
      jsonrpc: "2.0",
      id: null,
      error: { code: -32600, message: "Invalid JSON-RPC 2.0 request." }
    });
    return;
  }
  const auth = await requestAuthorization(request, messages);
  const results = (await Promise.all(messages.map((message) => processMcpMessage(message, auth)))).filter(Boolean);
  if (!results.length) {
    setCommonHeaders(response);
    response.writeHead(202);
    response.end();
    return;
  }
  sendJson(response, 200, Array.isArray(payload) ? results : results[0]);
}

export function createOneDriveHttpServer(env = process.env) {
  validateOAuthConfiguration(env);
  return createServer(async (request, response) => {
    try {
      const url = new URL(request.url || "/", "http://localhost");
      if (request.method === "OPTIONS") {
        setCommonHeaders(response);
        response.writeHead(204);
        response.end();
        return;
      }
      if (url.pathname === "/healthz" && request.method === "GET") {
        const settings = oauthSettings(env);
        sendJson(response, 200, {
          ok: true,
          server: "onedrive",
          transport: "streamable-http",
          authMode: settings.mode
        });
        return;
      }
      if (["/.well-known/oauth-protected-resource", "/.well-known/oauth-protected-resource/mcp"].includes(url.pathname)
        && request.method === "GET") {
        const metadata = protectedResourceMetadata(env);
        if (!metadata) {
          sendText(response, 404, "OAuth protected-resource metadata is disabled.\n");
          return;
        }
        sendJson(response, 200, metadata);
        return;
      }
      if (url.pathname === "/mcp") {
        await handleMcp(request, response);
        return;
      }
      sendText(response, 404, "Not Found\n");
    } catch (error) {
      sendJson(response, error.status || 500, {
        error: error.status && error.status < 500 ? "invalid_request" : "server_error",
        message: error.status && error.status < 500 ? error.message : "The OneDrive MCP server could not process the request."
      });
    }
  });
}

async function main() {
  const { host, port } = listenAddress();
  const server = createOneDriveHttpServer();
  await new Promise((resolvePromise, reject) => {
    server.once("error", reject);
    server.listen(port, host, resolvePromise);
  });
  process.stderr.write(`OneDrive MCP HTTP server listening on http://${host}:${port}/mcp\n`);
  const stop = async () => {
    await new Promise((resolvePromise) => server.close(resolvePromise));
    await shutdownOneDriveServer();
  };
  for (const signal of ["SIGINT", "SIGTERM"]) {
    process.once(signal, async () => {
      await stop();
      process.exit(0);
    });
  }
}

const isMainModule = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMainModule) {
  main().catch((error) => {
    process.stderr.write(`${error.stack || error.message}\n`);
    process.exit(1);
  });
}
