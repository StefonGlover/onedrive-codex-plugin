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
import {
  createResourceReadAdmissionController,
  holdResourceReadUntilResponseDeadline
} from "./resource-read-admission.mjs";
import { activeServerRelease, activeToolProfile, processMcpMessage, shutdownOneDriveServer } from "./server.mjs";

const maxRequestBytes = 1024 * 1024;
const maxBatchMessages = 16;
const maxJsonResponseBytes = 40 * 1024 * 1024;
const defaultResourceReadResponseDeadlineMs = 60_000;
const hostedLoopbackHosts = new Set(["127.0.0.1", "::1", "localhost"]);
let lastAuthFailure = null;
let lastToolFailure = null;
let lastToolCall = null;
const serverStartedAt = Date.now();
const recentToolDurationsMs = [];
let toolCallCount = 0;
let toolErrorCount = 0;
let graphThrottleCount = 0;
let oauthFailureCount = 0;
const defaultResourceReadAdmission = createResourceReadAdmissionController();

function percentile(values, probability) {
  if (!values.length) return null;
  const sorted = [...values].sort((left, right) => left - right);
  return Math.round(sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * probability) - 1))] * 10) / 10;
}

function observabilitySnapshot() {
  return {
    release: activeServerRelease,
    uptimeSeconds: Math.max(0, Math.floor((Date.now() - serverStartedAt) / 1000)),
    toolCalls: toolCallCount,
    toolErrors: toolErrorCount,
    graphThrottles: graphThrottleCount,
    oauthFailures: oauthFailureCount,
    latency: {
      sampleCount: recentToolDurationsMs.length,
      windowSize: 100,
      p50Ms: percentile(recentToolDurationsMs, 0.5),
      p95Ms: percentile(recentToolDurationsMs, 0.95)
    }
  };
}

function validateHostedToolProfile(env = process.env) {
  const requestedProfile = String(env.ONEDRIVE_TOOL_PROFILE || "").trim().toLowerCase();
  if (requestedProfile !== "chatgpt" || activeToolProfile !== "chatgpt") {
    throw new Error(
      "OneDrive Streamable HTTP requires ONEDRIVE_TOOL_PROFILE=chatgpt. "
      + "The full profile is restricted to trusted local stdio because it exposes local filesystem maintenance tools."
    );
  }
}

function hostedHttpHost(env = process.env) {
  const requestedHost = String(env.ONEDRIVE_MCP_HTTP_HOST ?? "").trim().toLowerCase();
  const host = requestedHost || "127.0.0.1";
  if (!hostedLoopbackHosts.has(host)) {
    throw new Error(
      "ONEDRIVE_MCP_HTTP_HOST must be a canonical loopback host: 127.0.0.1, ::1, or localhost. "
      + "Wildcard, LAN, and other hostname bindings are refused."
    );
  }
  return host;
}

function authFailureDiagnostic(authMode, error) {
  return {
    at: new Date().toISOString(),
    authMode,
    name: error?.name || "Error",
    code: error?.code || "server_error",
    status: Number.isInteger(error?.status) ? error.status : 500
  };
}

function recordToolDiagnostic(messages, results, fallbackDurationMs = null) {
  const toolCalls = messages.filter(isToolCall);
  const resultById = new Map(results.map((result) => [JSON.stringify([typeof result?.id, result?.id]), result]));
  for (const toolCall of toolCalls) {
    const toolResult = resultById.get(JSON.stringify([typeof toolCall?.id, toolCall?.id]))?.result;
    if (!toolResult) continue;
    const toolName = String(toolCall?.params?.name || "unknown").slice(0, 128);
    toolCallCount += 1;
    const reportedDurationMs = Number(toolResult?._meta?.["onedrive/performance"]?.totalMs);
    const durationMs = Number.isFinite(reportedDurationMs) && reportedDurationMs >= 0
      ? reportedDurationMs
      : fallbackDurationMs;
    if (Number.isFinite(durationMs) && durationMs >= 0) {
      recentToolDurationsMs.push(durationMs);
      if (recentToolDurationsMs.length > 100) recentToolDurationsMs.shift();
    }
    lastToolCall = {
      at: new Date().toISOString(),
      tool: toolName,
      isError: toolResult?.isError === true
    };
    if (toolResult?.isError) {
      toolErrorCount += 1;
      const error = toolResult?.structuredContent?.error || {};
      if (error.graphStatus === 429 || error.code === "rate_limited") graphThrottleCount += 1;
      lastToolFailure = {
        at: new Date().toISOString(),
        tool: toolName,
        code: String(error.code || "tool_error").slice(0, 128),
        ...(Number.isInteger(error.graphStatus) ? { graphStatus: error.graphStatus } : {})
      };
    } else {
      lastToolFailure = null;
    }
  }
}

function listenAddress(env = process.env) {
  const host = hostedHttpHost(env);
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
  let serialized;
  try {
    serialized = JSON.stringify(payload);
    if (typeof serialized !== "string") throw new Error("JSON serialization returned no value.");
  } catch {
    status = 500;
    serialized = JSON.stringify({
      jsonrpc: "2.0",
      id: null,
      error: { code: -32603, message: "The MCP response could not be serialized safely." }
    });
  }
  let body = `${serialized}\n`;
  if (Buffer.byteLength(body) > maxJsonResponseBytes) {
    status = 500;
    body = `${JSON.stringify({
      jsonrpc: "2.0",
      id: null,
      error: { code: -32003, message: `The MCP JSON response exceeds the ${maxJsonResponseBytes}-byte limit.` }
    })}\n`;
  }
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

function isStaticSkillResourceRead(message) {
  return message?.method === "resources/read"
    && typeof message?.params?.uri === "string"
    && message.params.uri.startsWith("skill://onedrive/");
}

function requiresOAuth(message) {
  return ["tools/list", "resources/list"].includes(message?.method)
    || (message?.method === "resources/read" && !isStaticSkillResourceRead(message))
    || isToolCall(message);
}

async function requestAuthorization(request, messages) {
  const authorization = request.headers.authorization;
  try {
    const settings = oauthSettings();
    if (settings.mode !== "oauth") return { authMode: "noauth" };
    const requiresAuthorization = Boolean(authorization) || messages.some(requiresOAuth);
    if (!requiresAuthorization) {
      lastAuthFailure = null;
      return { authMode: "oauth_discovery" };
    }
    const authorized = await authorizeMcpRequest(authorization, {
      requireGraph: messages.some(isToolCall)
    });
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
    if (authMode !== "oauth_required") oauthFailureCount += 1;
    return {
      authMode,
      error
    };
  }
}

function resourceReadAdmissionSubject(auth = {}) {
  if (typeof auth.authContextId === "string" && auth.authContextId) return auth.authContextId;
  if (auth.authMode === "noauth") return "noauth-local-server";
  return null;
}

function resourceReadResponseDeadlineMs(env = process.env) {
  const value = Number(env.ONEDRIVE_MCP_RESOURCE_RESPONSE_TIMEOUT_MS || defaultResourceReadResponseDeadlineMs);
  if (!Number.isSafeInteger(value) || value < 1_000 || value > 300_000) {
    throw new Error("ONEDRIVE_MCP_RESOURCE_RESPONSE_TIMEOUT_MS must be an integer from 1000 through 300000.");
  }
  return value;
}

async function handleMcp(
  request,
  response,
  resourceReadAdmission = defaultResourceReadAdmission,
  resourceReadResponseTimeoutMs = defaultResourceReadResponseDeadlineMs
) {
  if (request.method !== "POST") {
    sendText(response, 405, "Method Not Allowed\n", { Allow: "POST, OPTIONS" });
    return;
  }
  const payload = await readJsonBody(request);
  const isBatch = Array.isArray(payload);
  const messages = Array.isArray(payload) ? payload : [payload];
  if (isBatch && messages.length > maxBatchMessages) {
    sendJson(response, 400, {
      jsonrpc: "2.0",
      id: null,
      error: { code: -32600, message: `JSON-RPC batches support at most ${maxBatchMessages} messages.` }
    });
    return;
  }
  if (!messages.length || messages.some((message) => !message || typeof message !== "object" || message.jsonrpc !== "2.0")) {
    sendJson(response, 400, {
      jsonrpc: "2.0",
      id: null,
      error: { code: -32600, message: "Invalid JSON-RPC 2.0 request." }
    });
    return;
  }
  if (isBatch && messages.length > 1 && messages.some((message) => message.method === "resources/read" && !isStaticSkillResourceRead(message))) {
    sendJson(response, 400, {
      jsonrpc: "2.0",
      id: null,
      error: { code: -32600, message: "resources/read must be sent as a standalone JSON-RPC message." }
    });
    return;
  }
  const auth = await requestAuthorization(request, messages);
  if (auth?.authMode === "oauth_required" || auth?.authMode === "oauth_error") {
    const results = messages.map((message) => ({
      jsonrpc: "2.0",
      id: message.id ?? null,
      error: {
        code: -32001,
        message: auth.error?.message || "OneDrive authentication is required."
      }
    }));
    sendJson(response, 401, Array.isArray(payload) ? results : results[0], {
      "WWW-Authenticate": oauthChallenge({
        error: auth.error?.code || "invalid_token",
        description: auth.error?.message
      })
    });
    return;
  }
  if (messages.length === 1
    && messages[0].method === "resources/read"
    && !isStaticSkillResourceRead(messages[0])
    && auth?.authMode !== "oauth_server_error") {
    const subject = resourceReadAdmissionSubject(auth);
    if (!subject) {
      sendJson(response, 503, {
        jsonrpc: "2.0",
        id: messages[0].id ?? null,
        error: { code: -32005, message: "Resource reads are unavailable without an isolated authentication context." }
      });
      return;
    }
    const lease = resourceReadAdmission.acquire(subject);
    if (!lease.admitted) {
      sendJson(response, 429, {
        jsonrpc: "2.0",
        id: messages[0].id ?? null,
        error: {
          code: -32004,
          message: "A large OneDrive resource response is already in progress. Retry shortly.",
          data: {
            retryable: true,
            retryAfterSeconds: lease.retryAfterSeconds
          }
        }
      }, { "Retry-After": String(lease.retryAfterSeconds) });
      return;
    }
    try {
      holdResourceReadUntilResponseDeadline({
        lease,
        response,
        deadlineMs: resourceReadResponseTimeoutMs
      });
    } catch (error) {
      lease.release();
      throw error;
    }
  }
  const processingStartedAt = performance.now();
  const results = (await Promise.all(messages.map((message) => processMcpMessage(message, auth)))).filter(Boolean);
  const processingDurationMs = Math.round((performance.now() - processingStartedAt) * 10) / 10;
  recordToolDiagnostic(messages, results, processingDurationMs);
  if (!results.length) {
    setCommonHeaders(response);
    response.writeHead(202);
    response.end();
    return;
  }
  if (auth?.authMode === "oauth_server_error") {
    sendJson(response, 503, Array.isArray(payload) ? results : results[0]);
    return;
  }
  sendJson(response, 200, Array.isArray(payload) ? results : results[0]);
}

export function createOneDriveHttpServer(env = process.env, { resourceReadAdmission = defaultResourceReadAdmission } = {}) {
  validateHostedToolProfile(env);
  hostedHttpHost(env);
  validateOAuthConfiguration(env);
  if (!resourceReadAdmission || typeof resourceReadAdmission.acquire !== "function") {
    throw new Error("createOneDriveHttpServer requires a resource-read admission controller.");
  }
  const resourceReadResponseTimeoutMs = resourceReadResponseDeadlineMs(env);
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
          observability: observabilitySnapshot(),
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
        await handleMcp(request, response, resourceReadAdmission, resourceReadResponseTimeoutMs);
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
