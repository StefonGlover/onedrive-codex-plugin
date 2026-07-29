#!/usr/bin/env node

import { createServer } from "node:http";
import { generateKeyPairSync, sign } from "node:crypto";
import { chmod, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { issueFacadeAccessToken } from "../mcp/oauth-facade-token.mjs";

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

const primaryKeys = generateKeyPairSync("rsa", { modulusLength: 2048 });
const rotatedKeys = generateKeyPairSync("rsa", { modulusLength: 2048 });
const publicJwk = primaryKeys.publicKey.export({ format: "jwk" });
const rotatedPublicJwk = rotatedKeys.publicKey.export({ format: "jwk" });
const keyId = "oauth-http-test-key";
const rotatedKeyId = "oauth-http-test-rotated-key";
const apiClientId = "11111111-1111-4111-8111-111111111111";
const chatGptClientId = "33333333-3333-4333-8333-333333333333";
const apiResource = `api://${apiClientId}`;
const protectedResource = "https://onedrive-mcp.example.test/v1/mcp/tunnel_test";
const resourceMetadataUrl = "https://onedrive-mcp.example.test/.well-known/oauth-protected-resource/v1/mcp/tunnel_test";
const apiScope = `${apiResource}/access_as_user`;
const resourceScopes = [apiScope];
const oauthRequests = [];
let activeJwks = [];
let jwksRequestCount = 0;
let discoveryFailureStatus = 0;
let omitDiscoveryIssuer = false;
let jwksFailureStatus = 0;
let jwksPayloadOverride = null;
let tokenFailure = null;
let issuer;

const identityServer = createServer(async (request, response) => {
  const url = new URL(request.url, "http://localhost");
  if (url.pathname === "/.well-known/openid-configuration") {
    if (discoveryFailureStatus) {
      json(response, discoveryFailureStatus, { error: "metadata_unavailable" });
      return;
    }
    json(response, 200, {
      ...(omitDiscoveryIssuer ? {} : { issuer }),
      authorization_endpoint: `${issuer}/authorize`,
      token_endpoint: `${issuer}/token`,
      jwks_uri: `${issuer}/keys`,
      code_challenge_methods_supported: ["S256"],
      token_endpoint_auth_methods_supported: ["client_secret_post"]
    });
    return;
  }
  if (url.pathname === "/keys") {
    jwksRequestCount += 1;
    if (jwksFailureStatus) {
      json(response, jwksFailureStatus, { error: "jwks_unavailable" });
      return;
    }
    json(response, 200, jwksPayloadOverride || { keys: activeJwks });
    return;
  }
  if (url.pathname === "/token" && request.method === "POST") {
    const chunks = [];
    for await (const chunk of request) chunks.push(chunk);
    const form = new URLSearchParams(Buffer.concat(chunks).toString("utf8"));
    oauthRequests.push(Object.fromEntries(form));
    if (tokenFailure) {
      json(response, tokenFailure.status, tokenFailure.payload);
      return;
    }
    if (form.get("client_id") !== apiClientId || form.get("client_secret") !== "test-client-secret") {
      json(response, 401, { error: "invalid_client", error_description: "The test API client credentials are invalid." });
      return;
    }
    if (form.get("requested_token_use") !== "on_behalf_of"
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
const testJwk = (jwk, kid, jwkIssuer = issuer) => ({
  ...jwk,
  kid,
  use: "sig",
  alg: "RS256",
  issuer: jwkIssuer
});
activeJwks = [testJwk(publicJwk, keyId)];
const storageRoot = await mkdtemp(join(tmpdir(), "onedrive-oauth-http-test-"));
const facadeAccessTokenKeyFile = join(storageRoot, "facade-access.key");
await writeFile(facadeAccessTokenKeyFile, "11".repeat(32), { mode: 0o600 });
await chmod(facadeAccessTokenKeyFile, 0o600);
for (const key of [
  "ONEDRIVE_MCP_RESOURCE",
  "ONEDRIVE_MCP_PUBLIC_BASE_URL",
  "ONEDRIVE_MCP_OAUTH_API_RESOURCE",
  "ONEDRIVE_MCP_PROTECTED_RESOURCE",
  "ONEDRIVE_MCP_RESOURCE_METADATA_URL",
  "ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS"
]) {
  delete process.env[key];
}
Object.assign(process.env, {
  ONEDRIVE_MCP_AUTH_MODE: "oauth",
  ONEDRIVE_MCP_OAUTH_API_RESOURCE: apiResource,
  ONEDRIVE_MCP_PROTECTED_RESOURCE: protectedResource,
  ONEDRIVE_MCP_RESOURCE_METADATA_URL: resourceMetadataUrl,
  ONEDRIVE_MCP_OAUTH_API_CLIENT_ID: apiClientId,
  ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET: "test-client-secret",
  ONEDRIVE_MCP_OAUTH_API_SCOPE: apiScope,
  ONEDRIVE_MCP_OAUTH_SCOPE_CLAIM: "access_as_user",
  ONEDRIVE_MCP_OAUTH_AUDIENCE: apiClientId,
  ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS: chatGptClientId,
  ONEDRIVE_MCP_OAUTH_ISSUER: issuer,
  ONEDRIVE_MCP_OAUTH_AUTHORITY: issuer,
  ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER: "https://oauth-adapter.example.test",
  ONEDRIVE_MCP_OAUTH_DISCOVERY_URL: `${issuer}/.well-known/openid-configuration`,
  ONEDRIVE_MCP_OAUTH_OBO_TOKEN_ENDPOINT: `${issuer}/token`,
  ONEDRIVE_MCP_OAUTH_GRAPH_SCOPES: "https://graph.microsoft.com/.default",
  ONEDRIVE_MCP_OAUTH_FACADE_ACCESS_TOKEN_KEY_FILE: facadeAccessTokenKeyFile,
  ONEDRIVE_STORAGE_ROOT: storageRoot,
  ONEDRIVE_CACHE_ROOT: join(storageRoot, "cache"),
  ONEDRIVE_CLIENT_ID: apiClientId
});

function bearerToken(overrides = {}, signingKeys = primaryKeys, kid = keyId) {
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: "RS256", typ: "JWT", kid })).toString("base64url");
  const claims = Buffer.from(JSON.stringify({
    iss: issuer,
    aud: apiClientId,
    azp: chatGptClientId,
    ver: "2.0",
    tid: "9188040d-6c67-4c5b-b112-36a304b66dad",
    sub: "oauth-test-user",
    oid: "22222222-2222-4222-8222-222222222222",
    scp: "access_as_user",
    iat: now,
    nbf: now - 5,
    exp: now + 3600,
    ...overrides
  })).toString("base64url");
  const signingInput = `${header}.${claims}`;
  return `${signingInput}.${sign("RSA-SHA256", Buffer.from(signingInput), signingKeys.privateKey).toString("base64url")}`;
}

let mcpServer;
try {
  const oauth = await import("../mcp/oauth.mjs");
  const validConfiguration = { ...process.env };
  const withoutAllowedClient = { ...validConfiguration };
  delete withoutAllowedClient.ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS;
  assertThrows(
    () => oauth.validateOAuthConfiguration(withoutAllowedClient),
    "ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS"
  );
  assertThrows(
    () => oauth.validateOAuthConfiguration({
      ...validConfiguration,
      ONEDRIVE_MCP_PROTECTED_RESOURCE: "http://onedrive-mcp.example.test/mcp"
    }),
    "must use HTTPS"
  );
  assertThrows(
    () => oauth.validateOAuthConfiguration({
      ...validConfiguration,
      ONEDRIVE_MCP_PROTECTED_RESOURCE: "https://REPLACE_WITH_TUNNEL_ID.example.test/mcp"
    }),
    "deployment placeholder"
  );
  assertThrows(
    () => oauth.validateOAuthConfiguration({
      ...validConfiguration,
      ONEDRIVE_MCP_OAUTH_API_SCOPE: "api://wrong-resource/access_as_user"
    }),
    "must equal"
  );
  assertThrows(
    () => oauth.validateOAuthConfiguration({
      ...validConfiguration,
      ONEDRIVE_MCP_OAUTH_AUDIENCE: "66666666-6666-4666-8666-666666666666"
    }),
    "must equal ONEDRIVE_MCP_OAUTH_API_CLIENT_ID"
  );
  assertThrows(
    () => oauth.validateOAuthConfiguration({
      ...validConfiguration,
      ONEDRIVE_MCP_PROTECTED_RESOURCE: `${protectedResource}?tenant=unsafe`
    }),
    "must not include a query string"
  );
  const derivedMetadataConfiguration = { ...validConfiguration };
  delete derivedMetadataConfiguration.ONEDRIVE_MCP_RESOURCE_METADATA_URL;
  assert(
    oauth.oauthSettings(derivedMetadataConfiguration).resourceMetadataUrl === resourceMetadataUrl,
    "The default protected-resource metadata URL does not follow RFC 9728 path insertion."
  );

  const organizationTenantId = "44444444-4444-4444-8444-444444444444";
  const consumerTenantId = "9188040d-6c67-4c5b-b112-36a304b66dad";
  const issuerMetadata = { issuer: "https://login.microsoftonline.com/{tenantid}/v2.0" };
  const issuerClaims = (tenantId) => ({
    iss: `https://login.microsoftonline.com/${tenantId}/v2.0`,
    tid: tenantId
  });
  assert(
    oauth.issuerMatchesForTests(issuerClaims(organizationTenantId), issuerMetadata, { tenant: "common", issuer: "" }),
    "The common authority rejected a matching organizational issuer."
  );
  assert(
    oauth.issuerMatchesForTests(issuerClaims(consumerTenantId), issuerMetadata, { tenant: "consumers", issuer: "" }),
    "The consumers authority rejected the Microsoft account issuer."
  );
  assert(
    !oauth.issuerMatchesForTests(issuerClaims(organizationTenantId), issuerMetadata, { tenant: "consumers", issuer: "" }),
    "The consumers authority accepted an organizational issuer."
  );
  assert(
    oauth.issuerMatchesForTests(issuerClaims(organizationTenantId), issuerMetadata, { tenant: "organizations", issuer: "" }),
    "The organizations authority rejected an organizational issuer."
  );
  assert(
    !oauth.issuerMatchesForTests(issuerClaims(consumerTenantId), issuerMetadata, { tenant: "organizations", issuer: "" }),
    "The organizations authority accepted the Microsoft account issuer."
  );
  assert(
    !oauth.issuerMatchesForTests({
      iss: `https://login.microsoftonline.com/${organizationTenantId}/v2.0`,
      tid: consumerTenantId
    }, issuerMetadata, { tenant: "common", issuer: "" }),
    "Issuer validation accepted a mismatched tenant claim."
  );
  assert(
    !oauth.issuerMatchesForTests(
      issuerClaims(consumerTenantId),
      issuerMetadata,
      { tenant: "organizations", issuer: `https://login.microsoftonline.com/${consumerTenantId}/v2.0` }
    ),
    "An explicit issuer override bypassed the organizations tenant policy."
  );
  assert(
    !oauth.issuerMatchesForTests(
      { iss: "https://login.microsoftonline.com/not-a-guid/v2.0", tid: "not-a-guid" },
      issuerMetadata,
      { tenant: "common", issuer: "" }
    ),
    "The common authority accepted a non-GUID tenant claim."
  );
  assert(
    !oauth.issuerMatchesForTests(
      issuerClaims(organizationTenantId),
      { issuer: `https://login.microsoftonline.com/${consumerTenantId}/v2.0` },
      { tenant: "common", issuer: "" }
    ),
    "Issuer validation fell back after a nonmatching discovery issuer."
  );

  const token = bearerToken();
  const verified = await oauth.verifyBearerToken(`Bearer ${token}`);
  assert(verified.claims.sub === "oauth-test-user", "JWT claims were not verified.", verified.claims);
  assert(verified.authContextId.length === 64, "OAuth identity was not converted to an opaque context ID.");
  assert(
    oauth.oboTokenEndpointForClaims(
      { tid: consumerTenantId },
      { oboTokenEndpoint: "" }
    ) === `https://login.microsoftonline.com/${consumerTenantId}/oauth2/v2.0/token`,
    "OBO exchange did not derive the tenant-specific Microsoft token endpoint."
  );
  assertThrows(
    () => oauth.oboTokenEndpointForClaims({ tid: "consumers" }, { oboTokenEndpoint: "" }),
    "valid tenant"
  );
  const graphToken = await oauth.exchangeForGraphToken(verified);
  assert(graphToken === "mock-graph-access-token", "OBO exchange did not return the Graph token.");
  assert(oauthRequests.length === 1, "OBO exchange should be cached for the same assertion.", oauthRequests);
  assert(await oauth.exchangeForGraphToken(verified) === graphToken, "Cached OBO token changed unexpectedly.");
  assert(oauthRequests.length === 1, "Cached OBO exchange called the token endpoint twice.", oauthRequests);
  const facadeUpstreamToken = bearerToken({ jti: "facade-upstream-token" });
  const facadeToken = issueFacadeAccessToken({
    providerAccessToken: facadeUpstreamToken,
    issuer: "https://oauth-adapter.example.test",
    audience: protectedResource,
    clientId: chatGptClientId,
    scope: apiScope,
    expiresIn: 3600,
    keyFile: facadeAccessTokenKeyFile
  }).accessToken;
  const facadeVerified = await oauth.verifyBearerToken(`Bearer ${facadeToken}`);
  assert(
    facadeVerified.upstreamAccessToken === facadeUpstreamToken
      && facadeVerified.claims.aud === protectedResource
      && facadeVerified.claims.azp === chatGptClientId,
    "The resource-bound facade token was not verified or unwrapped correctly.",
    facadeVerified.claims
  );
  assert(
    await oauth.exchangeForGraphToken(facadeVerified) === "mock-graph-access-token",
    "The facade token could not complete an OBO exchange."
  );
  assert(
    oauthRequests.at(-1)?.assertion === facadeUpstreamToken,
    "The OBO exchange did not use the Microsoft token securely carried by the facade token."
  );
  const canonicalTunnelResource =
    "https://api.openai.com/v1/mcp/tunnel_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
  const chatGptTunnelResourceAlias =
    "https://tunnel-service.gateway.unified-0.internal.api.openai.org/v1/mcp/tunnel_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
  process.env.ONEDRIVE_MCP_PROTECTED_RESOURCE = canonicalTunnelResource;
  process.env.ONEDRIVE_MCP_PROTECTED_RESOURCE_ALIASES =
    chatGptTunnelResourceAlias;
  try {
    const aliasFacadeToken = issueFacadeAccessToken({
      providerAccessToken: facadeUpstreamToken,
      issuer: "https://oauth-adapter.example.test",
      audience: chatGptTunnelResourceAlias,
      clientId: chatGptClientId,
      scope: apiScope,
      expiresIn: 3600,
      keyFile: facadeAccessTokenKeyFile
    }).accessToken;
    const aliasFacadeVerified = await oauth.verifyBearerToken(
      `Bearer ${aliasFacadeToken}`
    );
    assert(
      aliasFacadeVerified.claims.aud === chatGptTunnelResourceAlias,
      "The exact same-tunnel ChatGPT gateway audience was not accepted.",
      aliasFacadeVerified.claims
    );
  } finally {
    process.env.ONEDRIVE_MCP_PROTECTED_RESOURCE = protectedResource;
    delete process.env.ONEDRIVE_MCP_PROTECTED_RESOURCE_ALIASES;
  }
  await assertRejects(() => oauth.verifyBearerToken(`Bearer ${bearerToken({ aud: "wrong-audience" })}`), "not minted for this MCP API");
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ azp: "55555555-5555-4555-8555-555555555555" })}`),
    "allowed ChatGPT client application"
  );
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ ver: "1.0", azp: undefined, appid: chatGptClientId })}`),
    "v2 token format"
  );
  await assertRejects(() => oauth.verifyBearerToken(`Bearer ${bearerToken({ scp: "wrong.scope" })}`), "missing the required");
  await assertRejects(() => oauth.verifyBearerToken(`Bearer ${bearerToken({ exp: 1 })}`), "expired");

  activeJwks = [
    testJwk(rotatedPublicJwk, keyId, "https://login.microsoftonline.com/wrong-tenant/v2.0"),
    testJwk(publicJwk, keyId)
  ];
  oauth.resetOAuthCachesForTests();
  const duplicateKidVerified = await oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "duplicate-kid" })}`);
  assert(duplicateKidVerified.claims.jti === "duplicate-kid", "Duplicate signing key IDs were not filtered by JWK issuer.");

  activeJwks = [testJwk(publicJwk, keyId, "https://login.microsoftonline.com/wrong-tenant/v2.0")];
  oauth.resetOAuthCachesForTests();
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "wrong-jwk-issuer" })}`),
    "No matching OAuth signing key"
  );

  activeJwks = [{ ...publicJwk, kid: keyId, use: "sig", alg: "RS256" }];
  oauth.resetOAuthCachesForTests();
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "missing-jwk-issuer" })}`),
    "No matching OAuth signing key"
  );

  activeJwks = [testJwk(publicJwk, keyId)];
  oauth.resetOAuthCachesForTests();
  const refreshCountBeforeUnknownKids = jwksRequestCount;
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "unknown-kid-one" }, primaryKeys, "unknown-kid-one")}`),
    "No matching OAuth signing key"
  );
  const refreshCountAfterFirstUnknownKid = jwksRequestCount;
  assert(
    refreshCountAfterFirstUnknownKid === refreshCountBeforeUnknownKids + 2,
    "The first unknown key ID should perform one normal JWKS read and one bounded refresh.",
    { refreshCountBeforeUnknownKids, refreshCountAfterFirstUnknownKid }
  );
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "unknown-kid-one-repeat" }, primaryKeys, "unknown-kid-one")}`),
    "No matching OAuth signing key"
  );
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "unknown-kid-two" }, primaryKeys, "unknown-kid-two")}`),
    "No matching OAuth signing key"
  );
  assert(jwksRequestCount === refreshCountAfterFirstUnknownKid, "Unknown key IDs bypassed the JWKS refresh cooldown.");
  await assertRejects(
    () => oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "oversized-kid" }, primaryKeys, "k".repeat(257))}`),
    "no longer than 256 characters"
  );

  activeJwks = [testJwk(publicJwk, keyId)];
  oauth.resetOAuthCachesForTests();
  await oauth.verifyBearerToken(`Bearer ${bearerToken({ jti: "prime-jwks-cache" })}`);
  const rotationRefreshCount = jwksRequestCount;
  activeJwks = [testJwk(rotatedPublicJwk, rotatedKeyId)];
  const rotatedToken = bearerToken({ jti: "rotated-key" }, rotatedKeys, rotatedKeyId);
  const rotatedVerified = await oauth.verifyBearerToken(`Bearer ${rotatedToken}`);
  assert(rotatedVerified.claims.jti === "rotated-key", "JWKS rotation did not refresh and retry an unknown key ID.");
  assert(jwksRequestCount === rotationRefreshCount + 1, "JWKS rotation performed more than one refresh.");
  activeJwks = [
    testJwk(publicJwk, keyId),
    testJwk(rotatedPublicJwk, rotatedKeyId)
  ];
  oauth.resetOAuthCachesForTests();

  const { createOneDriveHttpServer } = await import("../mcp/http-server.mjs");
  mcpServer = createOneDriveHttpServer();
  const mcpPort = await listen(mcpServer);
  const baseUrl = `http://127.0.0.1:${mcpPort}`;

  const health = await fetch(`${baseUrl}/healthz`).then((response) => response.json());
  assert(health.ok && health.authMode === "oauth", "HTTP health route did not report OAuth mode.", health);

  const metadataResponse = await fetch(`${baseUrl}/.well-known/oauth-protected-resource/v1/mcp/tunnel_test`);
  assert(metadataResponse.status === 200, "Configured resource-path metadata route failed.", metadataResponse.status);
  const metadata = await metadataResponse.json();
  assert(metadata.resource === protectedResource, "Protected-resource metadata has the wrong resource.", metadata);
  assert(
    metadata.authorization_servers?.[0] === "https://oauth-adapter.example.test",
    "Protected-resource metadata has the wrong authorization server.",
    metadata
  );
  assert(
    JSON.stringify(metadata.scopes_supported) === JSON.stringify(resourceScopes),
    "Protected-resource metadata has the wrong resource scope set.",
    metadata
  );
  const rootMetadataResponse = await fetch(`${baseUrl}/.well-known/oauth-protected-resource`);
  assert(rootMetadataResponse.status === 200, "Root protected-resource metadata route failed.", rootMetadataResponse.status);
  const rootMetadata = await rootMetadataResponse.json();
  const mcpMetadataResponse = await fetch(`${baseUrl}/.well-known/oauth-protected-resource/mcp`);
  assert(mcpMetadataResponse.status === 200, "Standard MCP protected-resource metadata route failed.", mcpMetadataResponse.status);
  const mcpMetadata = await mcpMetadataResponse.json();
  assert(
    JSON.stringify(rootMetadata) === JSON.stringify(metadata)
      && JSON.stringify(mcpMetadata) === JSON.stringify(metadata),
    "Root, standard MCP, and configured resource-path metadata routes disagree.",
    { rootMetadata, mcpMetadata, metadata }
  );

  const unauthenticatedInitialize = await mcpRequest(baseUrl, {
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2025-11-25", capabilities: {}, clientInfo: { name: "oauth-test", version: "1" } }
  });
  assert(
    unauthenticatedInitialize.status === 200
      && unauthenticatedInitialize.body?.result?.serverInfo?.name === "onedrive",
    "Unauthenticated MCP initialize was not available for tunnel discovery.",
    unauthenticatedInitialize
  );
  assert(
    !unauthenticatedInitialize.headers.has("www-authenticate"),
    "Unauthenticated MCP initialize unexpectedly returned an authentication challenge.",
    [...unauthenticatedInitialize.headers]
  );

  const unauthenticatedToolList = await mcpRequest(
    baseUrl,
    { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} }
  );
  assert(
    unauthenticatedToolList.status === 401,
    "Unauthenticated tools/list did not trigger OAuth discovery.",
    unauthenticatedToolList
  );
  const toolListChallenge =
    unauthenticatedToolList.headers.get("www-authenticate") || "";
  assert(
    toolListChallenge.includes(`resource_metadata="${resourceMetadataUrl}"`),
    "Unauthenticated tools/list advertises the wrong metadata URL.",
    toolListChallenge
  );

  const initialize = await mcpCall(baseUrl, {
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2025-11-25", capabilities: {}, clientInfo: { name: "oauth-test", version: "1" } }
  }, bearerToken({ jti: "linked-initialize" }));
  assert(initialize.result?.serverInfo?.name === "onedrive", "HTTP MCP initialize failed.", initialize);

  const listed = await mcpCall(
    baseUrl,
    { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} },
    bearerToken({ jti: "linked-tools-list" })
  );
  const expectedToolCount = process.env.ONEDRIVE_TOOL_PROFILE === "chatgpt" ? 21 : 84;
  assert(listed.result?.tools?.length === expectedToolCount, "OAuth HTTP server did not expose the exact tool contract.", listed.result?.tools?.length);
  assert(listed.result.tools.every((tool) => tool.securitySchemes?.[0]?.type === "oauth2"), "A tool is missing oauth2 security metadata.");
  assert(
    listed.result.tools.every((tool) =>
      JSON.stringify(tool.securitySchemes?.[0]?.scopes) === JSON.stringify(resourceScopes)
    ),
    "A tool advertises the wrong OAuth scope set."
  );
  assert(listed.result.tools.every((tool) => tool._meta?.securitySchemes?.[0]?.type === "oauth2"), "A tool is missing mirrored OAuth security metadata.");
  assert(
    listed.result.tools.every((tool) =>
      JSON.stringify(tool._meta?.securitySchemes?.[0]?.scopes) === JSON.stringify(resourceScopes)
    ),
    "A tool mirrors the wrong OAuth scope set."
  );

  const unlinked = await mcpRequest(baseUrl, {
    jsonrpc: "2.0",
    id: 3,
    method: "tools/call",
    params: { name: "onedrive_config", arguments: {} }
  });
  assert(unlinked.status === 401, "Unlinked OAuth tool call did not return HTTP 401.", unlinked);
  const unlinkedChallenge = unlinked.headers.get("www-authenticate") || "";
  assert(unlinkedChallenge.startsWith("Bearer "), "OAuth challenge header is missing.", unlinked.body);
  assert(unlinkedChallenge.includes(`resource_metadata="${resourceMetadataUrl}"`), "MCP OAuth challenge advertises the wrong metadata URL.", unlinkedChallenge);
  assert(
    unlinkedChallenge.includes(`scope="${resourceScopes.join(" ")}"`),
    "MCP OAuth challenge advertises the wrong scope set.",
    unlinkedChallenge
  );
  assert(unlinkedChallenge.includes('error="invalid_token"'), "MCP OAuth challenge is missing the OAuth error.", unlinkedChallenge);
  assert(unlinkedChallenge.includes("error_description="), "MCP OAuth challenge is missing the OAuth error description.", unlinkedChallenge);

  const invalidBearer = await mcpRequest(baseUrl, {
    jsonrpc: "2.0",
    id: 4,
    method: "tools/call",
    params: { name: "onedrive_config", arguments: {} }
  }, "not-a-jwt");
  assert(invalidBearer.status === 401, "An invalid supplied bearer token did not return HTTP 401.", invalidBearer);
  const challenge = invalidBearer.headers.get("www-authenticate") || "";
  assert(challenge.startsWith("Bearer "), "Invalid-token HTTP OAuth challenge header is missing.", challenge);
  assert(challenge.includes(`resource_metadata="${resourceMetadataUrl}"`), "Invalid-token HTTP challenge advertises the wrong metadata URL.", challenge);
  assert(
    challenge.includes(`scope="${resourceScopes.join(" ")}"`),
    "Invalid-token HTTP challenge advertises the wrong scope set.",
    challenge
  );

  discoveryFailureStatus = 404;
  oauth.resetOAuthCachesForTests();
  try {
    assertServiceUnavailableResponse(await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 4,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "discovery-failure-test" })), "OAuth discovery failure");
  } finally {
    discoveryFailureStatus = 0;
    oauth.resetOAuthCachesForTests();
  }

  omitDiscoveryIssuer = true;
  try {
    assertServiceUnavailableResponse(await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 5,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "malformed-discovery-test" })), "Malformed OAuth discovery");
  } finally {
    omitDiscoveryIssuer = false;
    oauth.resetOAuthCachesForTests();
  }

  jwksFailureStatus = 404;
  try {
    assertServiceUnavailableResponse(await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 6,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "jwks-failure-test" })), "JWKS failure");
  } finally {
    jwksFailureStatus = 0;
    oauth.resetOAuthCachesForTests();
  }

  jwksPayloadOverride = { signing_keys: [] };
  try {
    assertServiceUnavailableResponse(await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 7,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "malformed-jwks-test" })), "Malformed JWKS");
  } finally {
    jwksPayloadOverride = null;
    oauth.resetOAuthCachesForTests();
  }

  tokenFailure = {
    status: 429,
    payload: { error: "temporarily_unavailable", error_description: "Test identity provider throttling." }
  };
  try {
    assertServiceUnavailableResponse(await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 8,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "obo-throttle-test" })), "OBO throttling");
  } finally {
    tokenFailure = null;
  }

  const originalAllowedClientIds = process.env.ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS;
  process.env.ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS = "";
  try {
    assertServiceUnavailableResponse(await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 9,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "runtime-configuration-test" })), "Runtime OAuth configuration failure");
  } finally {
    process.env.ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS = originalAllowedClientIds;
  }

  const originalSecret = process.env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET;
  process.env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET = "wrong-client-secret";
  try {
    const unavailable = await mcpRequest(baseUrl, {
      jsonrpc: "2.0",
      id: 10,
      method: "tools/call",
      params: { name: "onedrive_config", arguments: {} }
    }, bearerToken({ jti: "invalid-client-test" }));
    assertServiceUnavailableResponse(unavailable, "OBO provider configuration failure");
  } finally {
    process.env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET = originalSecret;
  }

  const linked = await mcpCall(baseUrl, {
    jsonrpc: "2.0",
    id: 11,
    method: "tools/call",
    params: { name: "onedrive_config", arguments: {} }
  }, bearerToken({ jti: "linked-http-call" }));
  assert(linked.result?.isError === false, "Linked OAuth tool call failed.", linked);
  assert(JSON.parse(linked.result.content[0].text).clientIdConfigured === true, "Linked config result is malformed.", linked);

  console.log(JSON.stringify({
    ok: true,
    checks: {
      failClosedConfiguration: true,
      rfc9728ProtectedResourceMetadataUrl: true,
      protectedResourceMetadata: true,
      rootAndPathProtectedResourceMetadataRoutes: true,
      issuerTenantMatrix: true,
      jwtSignatureIssuerAudienceAuthorizedClientExpiryAndScope: true,
      jwkIssuerAndDuplicateKidFiltering: true,
      jwksRotationRefresh: true,
      boundedUnknownKidRefresh: true,
      onBehalfOfExchangeAndCache: true,
      resourceBoundFacadeTokenAndOboExchange: true,
      streamableHttpInitialize: true,
      oauthToolDescriptors: expectedToolCount,
      http401AndRuntimeChallenge: true,
      anonymousDiscoveryAndProtectedToolCalls: true,
      invalidSuppliedTokenUsesHttp401: true,
      providerAndConfigurationFailuresAre503: true,
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

function assertThrows(action, messageFragment) {
  try {
    action();
  } catch (error) {
    assert(String(error.message).includes(messageFragment), `Expected rejection containing ${messageFragment}.`, error.message);
    return;
  }
  throw new Error(`Expected rejection containing ${messageFragment}.`);
}

function assertServiceUnavailableResponse(response, label) {
  assert(response.status === 503, `${label} did not return HTTP 503.`, response);
  assert(!response.headers.has("www-authenticate"), `${label} incorrectly prompted the user to reconnect.`);
  assert(response.body.result?.structuredContent?.error?.code === "service_unavailable", `${label} was not safely classified.`, response.body);
  assert(!response.body.result?._meta?.["mcp/www_authenticate"], `${label} included relink metadata.`, response.body);
}

async function mcpRequest(baseUrl, body, token = null) {
  const response = await fetch(`${baseUrl}/mcp`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      ...(token ? { Authorization: `Bearer ${token}` } : {})
    },
    body: JSON.stringify(body)
  });
  const text = await response.text();
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    assert(false, `MCP HTTP response ${response.status} was not JSON.`, text);
  }
  return { status: response.status, headers: response.headers, body: parsed };
}

async function mcpCall(baseUrl, body, token = null) {
  const response = await mcpRequest(baseUrl, body, token);
  assert(response.status >= 200 && response.status < 300, `MCP HTTP request failed with ${response.status}.`, response.body);
  return response.body;
}
