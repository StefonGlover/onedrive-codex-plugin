import { createHash, createPublicKey, verify as verifySignature } from "node:crypto";
import { readFileSync } from "node:fs";

const DEFAULT_MSA_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad";
const discoveryCache = new Map();
const jwksCache = new Map();
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

export function oauthSettings(env = process.env) {
  const mode = String(env.ONEDRIVE_MCP_AUTH_MODE || "noauth").trim().toLowerCase();
  if (!new Set(["noauth", "oauth"]).has(mode)) {
    throw new Error("ONEDRIVE_MCP_AUTH_MODE must be noauth or oauth.");
  }
  const tenant = String(env.ONEDRIVE_MCP_OAUTH_TENANT || env.ONEDRIVE_TENANT || "consumers").trim();
  const apiClientId = String(env.ONEDRIVE_MCP_OAUTH_API_CLIENT_ID || env.ONEDRIVE_CLIENT_ID || "").trim();
  const resource = String(env.ONEDRIVE_MCP_RESOURCE || (apiClientId ? `api://${apiClientId}` : "")).trim();
  const apiScope = String(env.ONEDRIVE_MCP_OAUTH_API_SCOPE || (resource ? `${resource.replace(/\/+$/, "")}/access_as_user` : "")).trim();
  const authority = normalizedAuthority(
    env.ONEDRIVE_MCP_OAUTH_AUTHORITY || `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/v2.0`
  );
  return {
    mode,
    tenant,
    apiClientId,
    resource,
    apiScope,
    scopeClaim: String(env.ONEDRIVE_MCP_OAUTH_SCOPE_CLAIM || apiScope.split("/").filter(Boolean).at(-1) || "access_as_user").trim(),
    authority,
    audience: String(env.ONEDRIVE_MCP_OAUTH_AUDIENCE || apiClientId).trim(),
    issuer: String(env.ONEDRIVE_MCP_OAUTH_ISSUER || "").trim(),
    discoveryUrl: String(env.ONEDRIVE_MCP_OAUTH_DISCOVERY_URL || `${authority}/.well-known/openid-configuration`).trim(),
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
  requiredSetting(settings.resource, "ONEDRIVE_MCP_RESOURCE");
  requiredSetting(settings.apiScope, "ONEDRIVE_MCP_OAUTH_API_SCOPE");
  requiredSetting(settings.audience, "ONEDRIVE_MCP_OAUTH_AUDIENCE");
  requiredSetting(settings.authority, "ONEDRIVE_MCP_OAUTH_AUTHORITY");
  requiredSetting(settings.graphScopes, "ONEDRIVE_MCP_OAUTH_GRAPH_SCOPES");
  return settings;
}

export function toolSecuritySchemes(env = process.env) {
  const settings = oauthSettings(env);
  return settings.mode === "oauth"
    ? [{ type: "oauth2", scopes: [requiredSetting(settings.apiScope, "ONEDRIVE_MCP_OAUTH_API_SCOPE")] }]
    : [{ type: "noauth" }];
}

export function protectedResourceMetadata(env = process.env) {
  const settings = validateOAuthConfiguration(env);
  if (settings.mode !== "oauth") return null;
  return {
    resource: settings.resource,
    authorization_servers: [settings.authority],
    scopes_supported: [settings.apiScope],
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

async function fetchJson(url, { method = "GET", headers, body } = {}) {
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
    throw new OAuthError(`OAuth provider rejected the request: ${description}`, {
      code: payload.error || "invalid_token",
      status: response.status >= 500 ? 503 : 401
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
  if (!metadata.jwks_uri || !metadata.token_endpoint) {
    throw new OAuthError("The OAuth discovery document is missing jwks_uri or token_endpoint.", {
      code: "server_error",
      status: 503
    });
  }
  return metadata;
}

function issuerMatches(claims, metadata, settings) {
  const issuer = String(claims.iss || "");
  if (!issuer) return false;
  if (settings.issuer) return issuer === settings.issuer.replace("{tenantid}", String(claims.tid || ""));
  const discovered = String(metadata.issuer || "");
  if (discovered) {
    const expected = discovered.replace("{tenantid}", String(claims.tid || ""));
    if (issuer === expected) return true;
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
  if (settings.tenant === "consumers") return claims.tid === DEFAULT_MSA_TENANT_ID;
  if (!["common", "organizations"].includes(settings.tenant)) {
    return parsed.pathname.split("/").filter(Boolean)[0] === settings.tenant;
  }
  return true;
}

function audienceMatches(claims, audience) {
  const values = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
  return values.filter(Boolean).includes(audience);
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
  const parsed = splitJwt(token);
  if (parsed.header.alg !== "RS256" || !parsed.header.kid) {
    throw new OAuthError("The bearer token must use RS256 and include a key ID.");
  }
  const metadata = await discovery(settings);
  const jwks = await cachedJson(jwksCache, metadata.jwks_uri, 15 * 60 * 1000);
  const jwk = Array.isArray(jwks.keys)
    ? jwks.keys.find((candidate) => candidate.kid === parsed.header.kid && (!candidate.use || candidate.use === "sig"))
    : null;
  if (!jwk) {
    jwksCache.delete(metadata.jwks_uri);
    throw new OAuthError("No matching OAuth signing key was found. Retry after provider key discovery refreshes.");
  }
  let key;
  try {
    key = createPublicKey({ key: jwk, format: "jwk" });
  } catch {
    throw new OAuthError("The OAuth signing key is invalid.");
  }
  if (!verifySignature("RSA-SHA256", Buffer.from(parsed.signingInput), key, parsed.signature)) {
    throw new OAuthError("The bearer token signature is invalid.");
  }
  const now = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(parsed.claims.exp) || parsed.claims.exp <= now - 60) {
    throw new OAuthError("The bearer token is expired.");
  }
  if (Number.isFinite(parsed.claims.nbf) && parsed.claims.nbf > now + 60) {
    throw new OAuthError("The bearer token is not active yet.");
  }
  if (!issuerMatches(parsed.claims, metadata, settings)) {
    throw new OAuthError("The bearer token issuer is not trusted.");
  }
  if (!audienceMatches(parsed.claims, settings.audience)) {
    throw new OAuthError("The bearer token was not minted for this MCP API.");
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

export async function exchangeForGraphToken(verified, env = process.env) {
  if (!verified?.token) throw new OAuthError("A verified bearer token is required for Graph token exchange.");
  const settings = validateOAuthConfiguration(env);
  const cacheKey = createHash("sha256").update(verified.token).update("\0").update(settings.graphScopes).digest("hex");
  pruneOboCache();
  const cached = oboCache.get(cacheKey);
  if (cached) return cached.accessToken;
  const metadata = verified.metadata || await discovery(settings);
  const payload = await fetchJson(metadata.token_endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: settings.apiClientId,
      client_secret: settings.clientSecret,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: verified.token,
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
  const metadataUrl = String(env.ONEDRIVE_MCP_RESOURCE_METADATA_URL || "").trim()
    || `${String(env.ONEDRIVE_MCP_PUBLIC_BASE_URL || "").replace(/\/+$/, "")}/.well-known/oauth-protected-resource`;
  const fields = [];
  if (metadataUrl.startsWith("https://")) fields.push(`resource_metadata="${metadataUrl}"`);
  if (settings.apiScope) fields.push(`scope="${settings.apiScope}"`);
  if (error) fields.push(`error="${String(error).replace(/["\\]/g, "")}"`);
  if (description) fields.push(`error_description="${String(description).replace(/["\\\r\n]/g, " ").slice(0, 180)}"`);
  return `Bearer ${fields.join(", ")}`;
}

export function resetOAuthCachesForTests() {
  discoveryCache.clear();
  jwksCache.clear();
  oboCache.clear();
}
