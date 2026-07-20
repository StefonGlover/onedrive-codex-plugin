#!/usr/bin/env node

import { createServer } from "node:http";
import { generateKeyPairSync, sign } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

function assert(condition, message, details = undefined) {
  if (!condition) {
    const error = new Error(message);
    error.details = details;
    throw error;
  }
}

function json(response, status, value) {
  const body = JSON.stringify(value);
  response.writeHead(status, { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) });
  response.end(body);
}

async function listen(server) {
  await new Promise((resolvePromise, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolvePromise);
  });
  return server.address().port;
}

async function close(server) {
  await new Promise((resolvePromise) => server.close(resolvePromise));
}

const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
const publicJwk = publicKey.export({ format: "jwk" });
const keyId = "oauth-http-test-key";
const apiClientId = "11111111-1111-4111-8111-111111111111";
const resource = `api://${apiClientId}`;
const apiScope = `${resource}/access_as_user`;
const oauthRequests = [];
let issuer;

const identityServer = createServer(async (request, response) => {
  const url = new URL(request.url, "http://localhost");
  if (url.pathname === "/.well-known/openid-configuration") {
    json(response, 200, {
      issuer,
      authorization_endpoint: `${issuer}/authorize`,
      token_endpoint: `${issuer}/token`,
      jwks_uri: `${issuer}/keys`,
      code_challenge_methods_supported: ["S256"],
      token_endpoint_auth_methods_supported: ["client_secret_post"]
    });
    return;
  }
  if (url.pathname === "/keys") {
    json(response, 200, { keys: [{ ...publicJwk, kid: keyId, use: "sig", alg: "RS256" }] });
    return;
  }
  if (url.pathname === "/token" && request.method === "POST") {
    const chunks = [];
    for await (const chunk of request) chunks.push(chunk);
    const form = new URLSearchParams(Buffer.concat(chunks).toString("utf8"));
    oauthRequests.push(Object.fromEntries(form));
    if (form.get("client_id") !== apiClientId
      || form.get("client_secret") !== "test-client-secret"
      || form.get("requested_token_use") !== "on_behalf_of"
      || form.get("scope") !== "https://graph.microsoft.com/.default") {
      json(response, 400, { error: "invalid_request", error_description: "Unexpected OBO request." });
      return;
    }
    json(response, 200, { access_token: "mock-graph-access-token", token_type: "Bearer", expires_in: 3600 });
    return;
  }
  response.writeHead(404);
  response.end();
});

const identityPort = await listen(identityServer);
issuer = `http://127.0.0.1:${identityPort}`;
const storageRoot = await mkdtemp(join(tmpdir(), "onedrive-oauth-http-test-"));
Object.assign(process.env, {
  ONEDRIVE_MCP_AUTH_MODE: "oauth",
  ONEDRIVE_MCP_RESOURCE: resource,
  ONEDRIVE_MCP_PUBLIC_BASE_URL: "https://onedrive-mcp.example.test",
  ONEDRIVE_MCP_OAUTH_API_CLIENT_ID: apiClientId,
  ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET: "test-client-secret",
  ONEDRIVE_MCP_OAUTH_API_SCOPE: apiScope,
  ONEDRIVE_MCP_OAUTH_SCOPE_CLAIM: "access_as_user",
  ONEDRIVE_MCP_OAUTH_AUDIENCE: apiClientId,
  ONEDRIVE_MCP_OAUTH_ISSUER: issuer,
  ONEDRIVE_MCP_OAUTH_AUTHORITY: issuer,
  ONEDRIVE_MCP_OAUTH_DISCOVERY_URL: `${issuer}/.well-known/openid-configuration`,
  ONEDRIVE_MCP_OAUTH_GRAPH_SCOPES: "https://graph.microsoft.com/.default",
  ONEDRIVE_STORAGE_ROOT: storageRoot,
  ONEDRIVE_CACHE_ROOT: join(storageRoot, "cache"),
  ONEDRIVE_CLIENT_ID: apiClientId
});

function bearerToken(overrides = {}) {
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: "RS256", typ: "JWT", kid: keyId })).toString("base64url");
  const claims = Buffer.from(JSON.stringify({
    iss: issuer,
    aud: apiClientId,
    sub: "oauth-test-user",
    oid: "22222222-2222-4222-8222-222222222222",
    scp: "access_as_user",
    iat: now,
    nbf: now - 5,
    exp: now + 3600,
    ...overrides
  })).toString("base64url");
  const signingInput = `${header}.${claims}`;
  return `${signingInput}.${sign("RSA-SHA256", Buffer.from(signingInput), privateKey).toString("base64url")}`;
}

let mcpServer;
try {
  const oauth = await import("../mcp/oauth.mjs");
  const token = bearerToken();
  const verified = await oauth.verifyBearerToken(`Bearer ${token}`);
  assert(verified.claims.sub === "oauth-test-user", "JWT claims were not verified.", verified.claims);
  assert(verified.authContextId.length === 64, "OAuth identity was not converted to an opaque context ID.");
  const graphToken = await oauth.exchangeForGraphToken(verified);
  assert(graphToken === "mock-graph-access-token", "OBO exchange did not return the Graph token.");
  assert(oauthRequests.length === 1, "OBO exchange should be cached for the same assertion.", oauthRequests);
  assert(await oauth.exchangeForGraphToken(verified) === graphToken, "Cached OBO token changed unexpectedly.");
  assert(oauthRequests.length === 1, "Cached OBO exchange called the token endpoint twice.", oauthRequests);
  await assertRejects(() => oauth.verifyBearerToken(`Bearer ${bearerToken({ aud: "wrong-audience" })}`), "not minted for this MCP API");
  await assertRejects(() => oauth.verifyBearerToken(`Bearer ${bearerToken({ scp: "wrong.scope" })}`), "missing the required");
  await assertRejects(() => oauth.verifyBearerToken(`Bearer ${bearerToken({ exp: 1 })}`), "expired");

  const { createOneDriveHttpServer } = await import("../mcp/http-server.mjs");
  mcpServer = createOneDriveHttpServer();
  const mcpPort = await listen(mcpServer);
  const baseUrl = `http://127.0.0.1:${mcpPort}`;

  const health = await fetch(`${baseUrl}/healthz`).then((response) => response.json());
  assert(health.ok && health.authMode === "oauth", "HTTP health route did not report OAuth mode.", health);

  const metadata = await fetch(`${baseUrl}/.well-known/oauth-protected-resource/mcp`).then((response) => response.json());
  assert(metadata.resource === resource, "Protected-resource metadata has the wrong resource.", metadata);
  assert(metadata.authorization_servers?.[0] === issuer, "Protected-resource metadata has the wrong issuer.", metadata);
  assert(metadata.scopes_supported?.[0] === apiScope, "Protected-resource metadata has the wrong scope.", metadata);

  const initialize = await mcpCall(baseUrl, {
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2025-11-25", capabilities: {}, clientInfo: { name: "oauth-test", version: "1" } }
  });
  assert(initialize.result?.serverInfo?.name === "onedrive", "HTTP MCP initialize failed.", initialize);

  const listed = await mcpCall(baseUrl, { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
  assert(listed.result?.tools?.length === 84, "OAuth HTTP server did not expose the exact tool contract.", listed.result?.tools?.length);
  assert(listed.result.tools.every((tool) => tool.securitySchemes?.[0]?.type === "oauth2"), "A tool is missing oauth2 security metadata.");
  assert(listed.result.tools.every((tool) => tool.securitySchemes?.[0]?.scopes?.[0] === apiScope), "A tool advertises the wrong OAuth scope.");

  const unlinked = await mcpCall(baseUrl, {
    jsonrpc: "2.0",
    id: 3,
    method: "tools/call",
    params: { name: "onedrive_config", arguments: {} }
  });
  assert(unlinked.result?.isError === true, "Unlinked OAuth tool call was not rejected.", unlinked);
  assert(unlinked.result?._meta?.["mcp/www_authenticate"]?.[0]?.startsWith("Bearer "), "OAuth challenge metadata is missing.", unlinked);

  const linked = await mcpCall(baseUrl, {
    jsonrpc: "2.0",
    id: 4,
    method: "tools/call",
    params: { name: "onedrive_config", arguments: {} }
  }, token);
  assert(linked.result?.isError === false, "Linked OAuth tool call failed.", linked);
  assert(JSON.parse(linked.result.content[0].text).clientIdConfigured === true, "Linked config result is malformed.", linked);

  console.log(JSON.stringify({
    ok: true,
    checks: {
      protectedResourceMetadata: true,
      jwtSignatureIssuerAudienceExpiryAndScope: true,
      onBehalfOfExchangeAndCache: true,
      streamableHttpInitialize: true,
      oauthToolDescriptors: 84,
      runtimeChallenge: true,
      authenticatedToolCall: true
    }
  }, null, 2));
} finally {
  if (mcpServer) await close(mcpServer);
  await close(identityServer);
  await rm(storageRoot, { recursive: true, force: true });
}

async function assertRejects(action, messageFragment) {
  try {
    await action();
  } catch (error) {
    assert(String(error.message).includes(messageFragment), `Expected rejection containing ${messageFragment}.`, error.message);
    return;
  }
  throw new Error(`Expected rejection containing ${messageFragment}.`);
}

async function mcpCall(baseUrl, body, token = null) {
  const response = await fetch(`${baseUrl}/mcp`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      ...(token ? { Authorization: `Bearer ${token}` } : {})
    },
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const details = await response.text();
    assert(false, `MCP HTTP request failed with ${response.status}.`, details);
  }
  return response.json();
}
