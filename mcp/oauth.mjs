import { createHash, createPublicKey, verify as verifySignature } from "node:crypto";
import { readFileSync } from "node:fs";
import {
  isFacadeAccessToken,
  validateFacadeAccessTokenKeyFile,
  verifyFacadeAccessToken
} from "./oauth-facade-token.mjs";

const DEFAULT_MSA_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad";
const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const PROVIDER_CONFIGURATION_ERRORS = new Set([
  "invalid_client",
  "unauthorized_client",
  "invalid_resource",
  "invalid_scope",
  "unsupported_grant_type",
  "invalid_request"
]);
const PROVIDER_TRANSIENT_ERRORS = new Set([
  "temporarily_unavailable",
  "server_error"
]);
const MAX_JWT_KEY_ID_LENGTH = 256;
const MAX_NEGATIVE_KEY_IDS_PER_JWKS = 64;
const JWKS_REFRESH_COOLDOWN_MS = 30_000;
const JWKS_NEGATIVE_KID_TTL_MS = 60_000;
const discoveryCache = new Map();
const jwksCache = new Map();
const jwksRefreshState = new Map();
const oboCache = new Map();

export class OAuthError extends Error {
  constructor(message, { code = "invalid_token", status = 401, scope } = {}) {
    super(message);
    this.name = "OAuthError";
    this.code = code;
    this.status = status;
    this.scope = scope;
  }
}

function requiredSetting(value, name) {
  if (!value) throw new Error(`${name} is required when ONEDRIVE_MCP_AUTH_MODE is oauth.`);
  return value;
}

function secretSetting(value, file, name) {
  if (value) return value;
  if (file) {
    try {
      const secret = readFileSync(file, "utf8").trim();
      if (secret) return secret;
    } catch (error) {
      throw new Error(`Could not read ${name} from ${file}: ${error.message}`);
    }
  }
  throw new Error(`${name} or ${name}_FILE is required when ONEDRIVE_MCP_AUTH_MODE is oauth.`);
}

function normalizedAuthority(value) {
  return String(value || "").replace(/\/+$/, "");
}

function protectedResourceMetadataUrl(resource) {
  if (!resource) return "";
  let parsed;
  try {
    parsed = new URL(resource);
  } catch {
    return "";
  }
  const resourcePath = parsed.pathname === "/" ? "" : parsed.pathname.replace(/\/+$/, "");
  return `${parsed.origin}/.well-known/oauth-protected-resource${resourcePath}`;
}

function splitClientIds(value) {
  return [...new Set(String(value || "").split(/[\s,]+/).map((entry) => entry.trim().toLowerCase()).filter(Boolean))];
}

function rejectPlaceholder(value, name) {
  if (/REPLACE_WITH|YOUR[-_ ]/i.test(String(value || ""))) {
    throw new Error(`${name} still contains a deployment placeholder.`);
  }
}

function validateGuid(value, name) {
  rejectPlaceholder(value, name);
  if (!GUID_PATTERN.test(String(value || ""))) throw new Error(`${name} must be a Microsoft Entra application client ID UUID.`);
}

function validateUrl(value, name, { httpsOnly = true, allowQuery = true } = {}) {
  rejectPlaceholder(value, name);
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    throw new Error(`${name} must be an absolute URL.`);
  }
  const loopbackHttp = parsed.protocol === "http:" && new Set(["127.0.0.1", "localhost", "::1", "[::1]"]).has(parsed.hostname);
  if ((httpsOnly && parsed.protocol !== "https:") || (!httpsOnly && parsed.protocol !== "https:" && !loopbackHttp)) {
    throw new Error(`${name} must use HTTPS${httpsOnly ? "" : " (temporary loopback HTTP is allowed only for tests)"}.`);
  }
  if (parsed.username || parsed.password || parsed.hash) throw new Error(`${name} must not include credentials or a fragment.`);
  if (!allowQuery && parsed.search) throw new Error(`${name} must not include a query string.`);
  return parsed;
}

export function oauthSettings(env = process.env) {
  const mode = String(env.ONEDRIVE_MCP_AUTH_MODE || "noauth").trim().toLowerCase();
  if (!new Set(["noauth", "oauth"]).has(mode)) {
    throw new Error("ONEDRIVE_MCP_AUTH_MODE must be noauth or oauth.");
  }
  const tenant = String(env.ONEDRIVE_MCP_OAUTH_TENANT || env.ONEDRIVE_TENANT || "consumers").trim().toLowerCase();
  const apiClientId = String(env.ONEDRIVE_MCP_OAUTH_API_CLIENT_ID || env.ONEDRIVE_CLIENT_ID || "").trim();
  const apiResource = String(env.ONEDRIVE_MCP_OAUTH_API_RESOURCE || (apiClientId ? `api://${apiClientId}` : "")).trim();
  const resource = String(env.ONEDRIVE_MCP_PROTECTED_RESOURCE || "").trim().replace(/\/+$/, "");
  const apiScope = String(
    env.ONEDRIVE_MCP_OAUTH_API_SCOPE
    || (apiResource ? `${apiResource.replace(/\/+$/, "")}/access_as_user` : "")
  ).trim();
  const resourceMetadataUrl = String(
    env.ONEDRIVE_MCP_RESOURCE_METADATA_URL
    || protectedResourceMetadataUrl(resource)
  ).trim();
  const authority = normalizedAuthority(
    env.ONEDRIVE_MCP_OAUTH_AUTHORITY || `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/v2.0`
  );
  const authorizationServer = normalizedAuthority(
    env.ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER || authority
  );
  return {
    mode,
    tenant,
    apiClientId,
    apiResource,
    resource,
    resourceMetadataUrl,
    apiScope,
    scopeClaim: String(env.ONEDRIVE_MCP_OAUTH_SCOPE_CLAIM || apiScope.split("/").filter(Boolean).at(-1) || "access_as_user").trim(),
    authority,
    authorizationServer,
    facadeAccessTokenKeyFile: String(
      env.ONEDRIVE_MCP_OAUTH_FACADE_ACCESS_TOKEN_KEY_FILE
      || env.ONEDRIVE_TOKEN_ENCRYPTION_KEY_FILE
      || ""
    ).trim(),
    audience: String(env.ONEDRIVE_MCP_OAUTH_AUDIENCE || apiClientId).trim(),
    allowedClientIds: splitClientIds(env.ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS),
    issuer: String(env.ONEDRIVE_MCP_OAUTH_ISSUER || "").trim(),
    discoveryUrl: String(env.ONEDRIVE_MCP_OAUTH_DISCOVERY_URL || `${authority}/.well-known/openid-configuration`).trim(),
    oboTokenEndpoint: String(env.ONEDRIVE_MCP_OAUTH_OBO_TOKEN_ENDPOINT || "").trim(),
    graphScopes: String(env.ONEDRIVE_MCP_OAUTH_GRAPH_SCOPES || "https://graph.microsoft.com/.default").trim(),
    clientSecret: mode === "oauth"
      ? secretSetting(
          env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET,
          env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET_FILE,
          "ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET"
        )
      : ""
  };
}

export function validateOAuthConfiguration(env = process.env) {
  const settings = oauthSettings(env);
  if (settings.mode === "noauth") return settings;
  requiredSetting(settings.apiClientId, "ONEDRIVE_MCP_OAUTH_API_CLIENT_ID (or ONEDRIVE_CLIENT_ID)");
  requiredSetting(settings.apiResource, "ONEDRIVE_MCP_OAUTH_API_RESOURCE");
  requiredSetting(settings.resource, "ONEDRIVE_MCP_PROTECTED_RESOURCE");
  requiredSetting(settings.resourceMetadataUrl, "ONEDRIVE_MCP_RESOURCE_METADATA_URL");
  requiredSetting(settings.apiScope, "ONEDRIVE_MCP_OAUTH_API_SCOPE");
  requiredSetting(settings.audience, "ONEDRIVE_MCP_OAUTH_AUDIENCE");
  requiredSetting(settings.authority, "ONEDRIVE_MCP_OAUTH_AUTHORITY");
  requiredSetting(settings.authorizationServer, "ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER");
  requiredSetting(settings.graphScopes, "ONEDRIVE_MCP_OAUTH_GRAPH_SCOPES");
  requiredSetting(
    settings.facadeAccessTokenKeyFile,
    "ONEDRIVE_MCP_OAUTH_FACADE_ACCESS_TOKEN_KEY_FILE"
  );
  validateFacadeAccessTokenKeyFile(
    settings.facadeAccessTokenKeyFile,
    "ONEDRIVE_MCP_OAUTH_FACADE_ACCESS_TOKEN_KEY_FILE"
  );
  validateGuid(settings.apiClientId, "ONEDRIVE_MCP_OAUTH_API_CLIENT_ID");
  if (!new Set(["common", "consumers", "organizations"]).has(settings.tenant)) {
    validateGuid(settings.tenant, "ONEDRIVE_MCP_OAUTH_TENANT");
  }
  if (!settings.allowedClientIds.length) {
    throw new Error("ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS must contain the ChatGPT Entra client application ID.");
  }
  for (const clientId of settings.allowedClientIds) validateGuid(clientId, "ONEDRIVE_MCP_OAUTH_ALLOWED_CLIENT_IDS");
  rejectPlaceholder(settings.apiResource, "ONEDRIVE_MCP_OAUTH_API_RESOURCE");
  rejectPlaceholder(settings.apiScope, "ONEDRIVE_MCP_OAUTH_API_SCOPE");
  rejectPlaceholder(settings.audience, "ONEDRIVE_MCP_OAUTH_AUDIENCE");
  if (settings.audience.toLowerCase() !== settings.apiClientId.toLowerCase()) {
    throw new Error("ONEDRIVE_MCP_OAUTH_AUDIENCE must equal ONEDRIVE_MCP_OAUTH_API_CLIENT_ID for Entra v2 access tokens.");
  }
  const expectedScope = `${settings.apiResource.replace(/\/+$/, "")}/${settings.scopeClaim}`;
  if (settings.apiScope !== expectedScope) {
    throw new Error(`ONEDRIVE_MCP_OAUTH_API_SCOPE must equal ${expectedScope}.`);
  }
  validateUrl(settings.resource, "ONEDRIVE_MCP_PROTECTED_RESOURCE", { allowQuery: false });
  validateUrl(settings.resourceMetadataUrl, "ONEDRIVE_MCP_RESOURCE_METADATA_URL", { allowQuery: false });
  validateUrl(settings.authority, "ONEDRIVE_MCP_OAUTH_AUTHORITY", { httpsOnly: false });
  validateUrl(settings.authorizationServer, "ONEDRIVE_MCP_OAUTH_AUTHORIZATION_SERVER", { httpsOnly: false, allowQuery: false });
  validateUrl(settings.discoveryUrl, "ONEDRIVE_MCP_OAUTH_DISCOVERY_URL", { httpsOnly: false });
  if (settings.oboTokenEndpoint) {
    validateUrl(settings.oboTokenEndpoint, "ONEDRIVE_MCP_OAUTH_OBO_TOKEN_ENDPOINT", { httpsOnly: false });
  }
  if (settings.issuer) validateUrl(settings.issuer.replace("{tenantid}", DEFAULT_MSA_TENANT_ID), "ONEDRIVE_MCP_OAUTH_ISSUER", { httpsOnly: false });
  return settings;
}

function resourceScopes(settings) {
  return [
    requiredSetting(settings.apiScope, "ONEDRIVE_MCP_OAUTH_API_SCOPE")
  ];
}

export function toolSecuritySchemes(env = process.env) {
  const settings = oauthSettings(env);
  return settings.mode === "oauth"
    ? [{
        type: "oauth2",
        scopes: resourceScopes(settings)
      }]
    : [{ type: "noauth" }];
}

export function protectedResourceMetadata(env = process.env) {
  const settings = validateOAuthConfiguration(env);
  if (settings.mode !== "oauth") return null;
  return {
    resource: settings.resource,
    authorization_servers: [settings.authorizationServer],
    scopes_supported: resourceScopes(settings),
    bearer_methods_supported: ["header"]
  };
}

function base64UrlJson(value, label) {
  try {
    return JSON.parse(Buffer.from(value, "base64url").toString("utf8"));
  } catch {
    throw new OAuthError(`The bearer token has an invalid ${label}.`);
  }
}

function splitJwt(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3 || parts.some((part) => !part)) {
    throw new OAuthError("The bearer token is not a signed JWT.");
  }
  return {
    header: base64UrlJson(parts[0], "header"),
    claims: base64UrlJson(parts[1], "claims"),
    signingInput: `${parts[0]}.${parts[1]}`,
    signature: Buffer.from(parts[2], "base64url")
  };
}

async function fetchJson(url, { method = "GET", headers, body, providerPhase = "metadata" } = {}) {
  let response;
  try {
    response = await fetch(url, {
      method,
      headers,
      body,
      signal: AbortSignal.timeout ? AbortSignal.timeout(15_000) : undefined
    });
  } catch (error) {
    throw new OAuthError(`OAuth provider request failed: ${error.message}`, { code: "temporarily_unavailable", status: 503 });
  }
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const description = payload.error_description || payload.error || `${response.status} ${response.statusText}`;
    const providerCode = String(payload.error || "invalid_token");
    const configurationFailure = providerPhase === "obo" && PROVIDER_CONFIGURATION_ERRORS.has(providerCode);
    const providerFailure = providerPhase !== "obo";
    const transientFailure = response.status === 429
      || response.status >= 500
      || PROVIDER_TRANSIENT_ERRORS.has(providerCode);
    throw new OAuthError(`OAuth provider rejected the request: ${description}`, {
      code: configurationFailure || providerFailure || transientFailure ? "server_error" : providerCode,
      status: configurationFailure || providerFailure || transientFailure ? 503 : 401
    });
  }
  return payload;
}

async function cachedJson(cache, url, maxAgeMs) {
  const cached = cache.get(url);
  if (cached && cached.expiresAt > Date.now()) return cached.value;
  const value = await fetchJson(url);
  cache.set(url, { value, expiresAt: Date.now() + maxAgeMs });
  return value;
}

async function discovery(settings) {
  const metadata = await cachedJson(discoveryCache, settings.discoveryUrl, 60 * 60 * 1000);
  if (!metadata.issuer || !metadata.jwks_uri || !metadata.token_endpoint) {
    throw new OAuthError("The OAuth discovery document is missing issuer, jwks_uri, or token_endpoint.", {
      code: "server_error",
      status: 503
    });
  }
  return metadata;
}

export function issuerMatchesForTests(claims, metadata, settings) {
  const issuer = String(claims.iss || "");
  if (!issuer) return false;
  const tenantId = String(claims.tid || "").toLowerCase();
  if (!GUID_PATTERN.test(tenantId)) return false;
  if (settings.tenant === "consumers" && tenantId !== DEFAULT_MSA_TENANT_ID) return false;
  if (settings.tenant === "organizations" && (!tenantId || tenantId === DEFAULT_MSA_TENANT_ID)) return false;
  if (!["common", "consumers", "organizations"].includes(settings.tenant)
    && tenantId !== String(settings.tenant || "").toLowerCase()) {
    return false;
  }
  if (settings.issuer) return issuer === settings.issuer.replace("{tenantid}", tenantId);
  const discovered = String(metadata.issuer || "");
  if (discovered) {
    const expected = discovered.replace("{tenantid}", tenantId);
    return issuer === expected;
  }
  let parsed;
  try {
    parsed = new URL(issuer);
  } catch {
    return false;
  }
  if (parsed.protocol !== "https:" || parsed.hostname !== "login.microsoftonline.com" || !parsed.pathname.endsWith("/v2.0")) {
    return false;
  }
  if (!tenantId) return false;
  const issuerTenant = String(parsed.pathname.split("/").filter(Boolean)[0] || "").toLowerCase();
  if (issuerTenant !== tenantId) return false;
  if (settings.tenant === "consumers") return tenantId === DEFAULT_MSA_TENANT_ID;
  if (settings.tenant === "organizations") return tenantId !== DEFAULT_MSA_TENANT_ID;
  if (!["common", "organizations"].includes(settings.tenant)) {
    return issuerTenant === settings.tenant.toLowerCase() && tenantId === settings.tenant.toLowerCase();
  }
  return true;
}

function issuerMatches(claims, metadata, settings) {
  return issuerMatchesForTests(claims, metadata, settings);
}

function jwkIssuerMatches(jwk, claims) {
  if (typeof jwk?.issuer !== "string" || !jwk.issuer) return false;
  const tenantId = String(claims.tid || "").toLowerCase();
  return String(jwk.issuer).replace("{tenantid}", tenantId) === String(claims.iss || "");
}

function validateJwksDocument(jwks) {
  if (!Array.isArray(jwks?.keys) || jwks.keys.length === 0) {
    throw new OAuthError("The OAuth provider returned a malformed or empty JWKS document.", {
      code: "server_error",
      status: 503
    });
  }
  return jwks;
}

function matchingJwks(jwks, header, claims) {
  return Array.isArray(jwks?.keys)
    ? jwks.keys.filter((candidate) =>
        candidate?.kid === header.kid
        && (!candidate.use || candidate.use === "sig")
        && (!candidate.alg || candidate.alg === "RS256")
        && jwkIssuerMatches(candidate, claims))
    : [];
}

async function refreshJwksForUnknownKid(url, kid, claims) {
  const now = Date.now();
  let state = jwksRefreshState.get(url);
  if (!state) {
    state = { refreshedAt: 0, promise: null, negativeKids: new Map() };
    jwksRefreshState.set(url, state);
  }
  for (const [negativeKey, expiresAt] of state.negativeKids.entries()) {
    if (expiresAt <= now) state.negativeKids.delete(negativeKey);
  }
  const negativeKey = `${kid}\0${String(claims.iss || "")}`;
  if ((state.negativeKids.get(negativeKey) || 0) > now) return null;
  if (state.promise) return await state.promise;
  if (state.refreshedAt > now - JWKS_REFRESH_COOLDOWN_MS) return null;
  state.promise = (async () => {
    const value = validateJwksDocument(await fetchJson(url));
    jwksCache.set(url, { value, expiresAt: Date.now() + 15 * 60 * 1000 });
    state.refreshedAt = Date.now();
    return value;
  })();
  try {
    const value = await state.promise;
    if (!matchingJwks(value, { kid, alg: "RS256" }, claims).length) {
      while (state.negativeKids.size >= MAX_NEGATIVE_KEY_IDS_PER_JWKS) {
        state.negativeKids.delete(state.negativeKids.keys().next().value);
      }
      state.negativeKids.set(negativeKey, Date.now() + JWKS_NEGATIVE_KID_TTL_MS);
    }
    return value;
  } finally {
    state.promise = null;
  }
}

function audienceMatches(claims, audience) {
  const values = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
  const expected = String(audience || "").toLowerCase();
  return values.filter(Boolean).some((value) => String(value).toLowerCase() === expected);
}

function scopeSet(claims) {
  return new Set(String(claims.scp || "").split(/\s+/).filter(Boolean));
}

function tokenFromHeader(header) {
  const match = String(header || "").match(/^Bearer\s+([^\s]+)$/i);
  if (!match) throw new OAuthError("A valid Authorization: Bearer header is required.");
  return match[1];
}

export async function verifyBearerToken(authorization, env = process.env) {
  const settings = validateOAuthConfiguration(env);
  if (settings.mode !== "oauth") return null;
  const token = tokenFromHeader(authorization);
  if (isFacadeAccessToken(token)) {
    let verified;
    try {
      verified = verifyFacadeAccessToken(token, {
        issuer: settings.authorizationServer,
        audience: settings.resource,
        clientId: settings.allowedClientIds[0],
        requiredScope: settings.apiScope,
        keyFile: settings.facadeAccessTokenKeyFile
      });
    } catch {
      throw new OAuthError("The facade bearer token is invalid.");
    }
    const subject = String(verified.claims.sub || "");
    return {
      token,
      upstreamAccessToken: verified.upstreamAccessToken,
      claims: verified.claims,
      authContextId: createHash("sha256")
        .update(`${verified.claims.iss}\0${subject}`)
        .digest("hex"),
      expiresAt: verified.claims.exp * 1000,
      settings,
      metadata: null
    };
  }
  const parsed = splitJwt(token);
  if (parsed.header.alg !== "RS256"
    || typeof parsed.header.kid !== "string"
    || !parsed.header.kid
    || parsed.header.kid.length > MAX_JWT_KEY_ID_LENGTH) {
    throw new OAuthError(`The bearer token must use RS256 and include a key ID no longer than ${MAX_JWT_KEY_ID_LENGTH} characters.`);
  }
  const metadata = await discovery(settings);
  if (parsed.claims.ver !== "2.0") {
    throw new OAuthError("The bearer token must use the Microsoft identity platform v2 token format.");
  }
  if (!issuerMatches(parsed.claims, metadata, settings)) {
    throw new OAuthError("The bearer token issuer is not trusted.");
  }
  let jwks = validateJwksDocument(await cachedJson(jwksCache, metadata.jwks_uri, 15 * 60 * 1000));
  let candidates = matchingJwks(jwks, parsed.header, parsed.claims);
  if (!candidates.length) {
    jwks = await refreshJwksForUnknownKid(metadata.jwks_uri, parsed.header.kid, parsed.claims);
    candidates = matchingJwks(jwks, parsed.header, parsed.claims);
    if (!candidates.length) throw new OAuthError("No matching OAuth signing key was found after bounded provider key discovery.");
  }
  let signatureValid = false;
  let importableKeyFound = false;
  for (const jwk of candidates) {
    try {
      const key = createPublicKey({ key: jwk, format: "jwk" });
      importableKeyFound = true;
      if (verifySignature("RSA-SHA256", Buffer.from(parsed.signingInput), key, parsed.signature)) {
        signatureValid = true;
        break;
      }
    } catch {
      // Try the next issuer-compatible key when Microsoft publishes duplicate key IDs.
    }
  }
  if (!importableKeyFound) {
    throw new OAuthError("The OAuth provider returned an invalid signing key.", {
      code: "server_error",
      status: 503
    });
  }
  if (!signatureValid) {
    throw new OAuthError("The bearer token signature is invalid.");
  }
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(parsed.claims.exp) || parsed.claims.exp <= now - 60) {
    throw new OAuthError("The bearer token is expired.");
  }
  if (Number.isFinite(parsed.claims.nbf) && parsed.claims.nbf > now + 60) {
    throw new OAuthError("The bearer token is not active yet.");
  }
  if (!audienceMatches(parsed.claims, settings.audience)) {
    throw new OAuthError("The bearer token was not minted for this MCP API.");
  }
  const authorizedClientId = String(parsed.claims.azp || "").toLowerCase();
  if (!authorizedClientId || !settings.allowedClientIds.includes(authorizedClientId)) {
    throw new OAuthError("The bearer token was not issued to an allowed ChatGPT client application.");
  }
  if (!scopeSet(parsed.claims).has(settings.scopeClaim)) {
    throw new OAuthError(`The bearer token is missing the required ${settings.scopeClaim} scope.`, {
      code: "insufficient_scope",
      scope: settings.apiScope
    });
  }
  const subject = String(parsed.claims.oid || parsed.claims.sub || "");
  if (!subject) throw new OAuthError("The bearer token is missing a stable user subject.");
  return {
    token,
    claims: parsed.claims,
    authContextId: createHash("sha256").update(`${parsed.claims.iss}\0${subject}`).digest("hex"),
    expiresAt: parsed.claims.exp * 1000,
    settings,
    metadata
  };
}

function pruneOboCache(now = Date.now()) {
  for (const [key, value] of oboCache.entries()) {
    if (!value || value.expiresAt <= now + 60_000) oboCache.delete(key);
  }
}

export function oboTokenEndpointForClaims(claims, settings) {
  if (settings?.oboTokenEndpoint) return settings.oboTokenEndpoint;
  const tenantId = String(claims?.tid || "").trim().toLowerCase();
  if (!GUID_PATTERN.test(tenantId)) {
    throw new OAuthError("The verified bearer token is missing a valid tenant for downstream token exchange.");
  }
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
}

export async function exchangeForGraphToken(verified, env = process.env) {
  if (!verified?.token) throw new OAuthError("A verified bearer token is required for Graph token exchange.");
  const settings = validateOAuthConfiguration(env);
  const cacheKey = createHash("sha256").update(verified.token).update("\0").update(settings.graphScopes).digest("hex");
  pruneOboCache();
  const cached = oboCache.get(cacheKey);
  if (cached) return cached.accessToken;
  const tokenEndpoint = oboTokenEndpointForClaims(verified.claims, settings);
  const payload = await fetchJson(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    providerPhase: "obo",
    body: new URLSearchParams({
      client_id: settings.apiClientId,
      client_secret: settings.clientSecret,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: verified.upstreamAccessToken || verified.token,
      requested_token_use: "on_behalf_of",
      scope: settings.graphScopes
    })
  });
  if (!payload.access_token) {
    throw new OAuthError("The on-behalf-of exchange did not return a Microsoft Graph access token.", {
      code: "server_error",
      status: 503
    });
  }
  const expiresIn = Number(payload.expires_in);
  const expiresAt = Date.now() + (Number.isFinite(expiresIn) && expiresIn > 0 ? expiresIn * 1000 : 5 * 60 * 1000);
  oboCache.set(cacheKey, { accessToken: payload.access_token, expiresAt });
  return payload.access_token;
}

export async function authorizeMcpRequest(authorization, { requireGraph = false, env = process.env } = {}) {
  const settings = validateOAuthConfiguration(env);
  if (settings.mode !== "oauth") return { authMode: "noauth" };
  const verified = await verifyBearerToken(authorization, env);
  return {
    authMode: "oauth",
    authContextId: verified.authContextId,
    oauthClaims: verified.claims,
    graphAccessToken: requireGraph ? await exchangeForGraphToken(verified, env) : null
  };
}

export function oauthChallenge({ error = "invalid_token", description, env = process.env } = {}) {
  const settings = oauthSettings(env);
  const metadataUrl = settings.resourceMetadataUrl;
  const fields = [];
  if (metadataUrl.startsWith("https://")) fields.push(`resource_metadata="${metadataUrl}"`);
  if (settings.apiScope) fields.push(`scope="${resourceScopes(settings).join(" ")}"`);
  if (error) fields.push(`error="${String(error).replace(/["\\]/g, "")}"`);
  if (description) fields.push(`error_description="${String(description).replace(/["\\\r\n]/g, " ").slice(0, 180)}"`);
  return `Bearer ${fields.join(", ")}`;
}

export function resetOAuthCachesForTests() {
  discoveryCache.clear();
  jwksCache.clear();
  jwksRefreshState.clear();
  oboCache.clear();
}
