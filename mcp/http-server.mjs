#!/usr/bin/env node

import { createServer } from "node:http";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import {
  authorizeMcpRequest,
  OAuthError,
  oauthChallenge,
  oauthSettings,
  protectedResourceMetadata,
  validateOAuthConfiguration
} from "./oauth.mjs";
import { processMcpMessage, shutdownOneDriveServer } from "./server.mjs";

const maxRequestBytes = 1024 * 1024;
let lastAuthFailure = null;
let lastToolFailure = null;
let lastToolCall = null;

function authFailureDiagnostic(authMode, error) {
  return {
    at: new Date().toISOString(),
    authMode,
    name: error?.name || "Error",
    code: error?.code || "server_error",
    status: Number.isInteger(error?.status) ? error.status : 500,
    message: String(error?.message || "OAuth authorization failed.").slice(0, 1024)
  };
}

function recordToolDiagnostic(messages, results) {
  const toolCalls = messages.filter(isToolCall);
  const toolResults = results.filter((result) => result?.result);
  for (let index = 0; index < Math.min(toolCalls.length, toolResults.length); index += 1) {
    const toolName = String(toolCalls[index]?.params?.name || "unknown").slice(0, 128);
    const toolResult = toolResults[index].result;
    lastToolCall = {
      at: new Date().toISOString(),
      tool: toolName,
      isError: toolResult?.isError === true
    };
    if (toolResult?.isError) {
      const error = toolResult?.structuredContent?.error || {};
      lastToolFailure = {
        at: new Date().toISOString(),
        tool: toolName,
        code: String(error.code || "tool_error").slice(0, 128),
        message: String(error.message || toolResult?.content?.[0]?.text || "Tool call failed.").slice(0, 1024),
        ...(Number.isInteger(error.graphStatus) ? { graphStatus: error.graphStatus } : {})
      };
    } else {
      lastToolFailure = null;
    }
  }
}

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
  const authorization = request.headers.authorization;
  try {
    const settings = oauthSettings();
    if (settings.mode !== "oauth") return { authMode: "noauth" };
    if (!messages.some(isToolCall)) return null;
    const authorized = await authorizeMcpRequest(authorization, { requireGraph: true });
    lastAuthFailure = null;
    return authorized;
  } catch (error) {
    const unauthenticatedToolCall = !authorization
      && error instanceof OAuthError
      && error.status === 401;
    const authMode = unauthenticatedToolCall
      ? "oauth_required"
      : error instanceof OAuthError && error.status === 401
        ? "oauth_error"
        : "oauth_server_error";
    lastAuthFailure = authFailureDiagnostic(authMode, error);
    return {
      authMode,
      error
    };
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
  recordToolDiagnostic(messages, results);
  if (!results.length) {
    setCommonHeaders(response);
    response.writeHead(202);
    response.end();
    return;
  }
  if (auth?.authMode === "oauth_required") {
    // Keep the MCP auth challenge in the successful transport response so
    // Secure MCP Tunnel can deliver the tool result's mcp/www_authenticate
    // metadata to ChatGPT. A transport-level 401 is reserved for a bearer
    // credential that was actually supplied but failed validation.
    sendJson(response, 200, Array.isArray(payload) ? results : results[0]);
    return;
  }
  if (auth?.authMode === "oauth_error") {
    sendJson(response, 401, Array.isArray(payload) ? results : results[0], {
      "WWW-Authenticate": oauthChallenge({
        error: auth.error?.code || "invalid_token",
        description: auth.error?.message
      })
    });
    return;
  }
  if (auth?.authMode === "oauth_server_error") {
    sendJson(response, 503, Array.isArray(payload) ? results : results[0]);
    return;
  }
  sendJson(response, 200, Array.isArray(payload) ? results : results[0]);
}

export function createOneDriveHttpServer(env = process.env) {
  validateOAuthConfiguration(env);
  const resourceMetadata = protectedResourceMetadata(env);
  const resourceMetadataRoutes = new Set();
  if (resourceMetadata) {
    resourceMetadataRoutes.add("/.well-known/oauth-protected-resource");
    resourceMetadataRoutes.add("/.well-known/oauth-protected-resource/mcp");
    const resourcePath = new URL(resourceMetadata.resource).pathname.replace(/\/+$/, "");
    if (resourcePath && resourcePath !== "/") {
      resourceMetadataRoutes.add(`/.well-known/oauth-protected-resource${resourcePath}`);
    }
  }
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
          authMode: settings.mode,
          ...(lastAuthFailure ? { lastAuthFailure } : {}),
          ...(lastToolFailure ? { lastToolFailure } : {}),
          ...(lastToolCall ? { lastToolCall } : {})
        });
        return;
      }
      if (resourceMetadataRoutes.has(url.pathname) && request.method === "GET") {
        sendJson(response, 200, resourceMetadata);
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
