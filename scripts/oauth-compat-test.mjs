#!/usr/bin/env node

import { createHash } from "node:crypto";
import {
  chmodSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createOAuthCompatServer, validateOAuthCompatConfiguration } from "../mcp/oauth-compat-server.mjs";
import { oauthChallenge } from "../mcp/oauth.mjs";
import { verifyFacadeAccessToken } from "../mcp/oauth-facade-token.mjs";

const CLIENT_ID = "3caa4df0-1aa6-4473-9b4a-ecdf8d73bddd";
const CLIENT_SECRET = "oauth-compat-test-secret-value";
const REDIRECT_URI = "https://chatgpt.com/connector/oauth/QTOb4VcHdCsW";
const RESOURCE = "https://onedrive-tunnel.example.test/v1/mcp/tunnel_test";
const ISSUER = "https://onedrive-oauth.example.test";
const API_SCOPE = "api://6e97d01c-edf8-43fe-bf69-bb494ae22513/access_as_user";
const SCOPES = [
  API_SCOPE,
  "offline_access"
];
const COMBINED_CONSENT_SCOPES = [
  "api://6e97d01c-edf8-43fe-bf69-bb494ae22513/.default",
  "offline_access"
];
const UPSTREAM_TOKEN = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token";
const UPSTREAM_AUTHORIZE = "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize";
const VERIFIER = "v".repeat(43);
const CHALLENGE = createHash("sha256").update(VERIFIER, "ascii").digest("base64url");
const ALTERNATE_VERIFIER = "w".repeat(43);
const ALTERNATE_CHALLENGE = createHash("sha256")
  .update(ALTERNATE_VERIFIER, "ascii")
  .digest("base64url");
const STATE = "state_value_0123456789";
const CODE = "mock-authorization-code";
const REFRESH_TOKEN = "mock-refresh-token";
const PROVIDER_REFRESH_TOKEN = "mock-provider-refresh-token";
const MOCK_PROVIDER_ACCESS_TOKEN = [
  Buffer.from(JSON.stringify({ alg: "RS256", kid: "mock" })).toString("base64url"),
  Buffer.from(JSON.stringify({
    aud: "6e97d01c-edf8-43fe-bf69-bb494ae22513",
    iss: "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0",
    iat: 1_700_000_000,
    nbf: 1_700_000_000,
    exp: 4_102_444_800,
    scp: "access_as_user",
    azp: CLIENT_ID,
    tid: "9188040d-6c67-4c5b-b112-36a304b66dad",
    oid: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
    ver: "2.0"
  })).toString("base64url"),
  "mock-signature"
].join(".");
const DUMMY_FACADE_CODE = `odac_${"A".repeat(43)}`;
const DUMMY_REFRESH_HANDLE = `odrh_${"A".repeat(43)}`;
const PROVIDER_CALLBACK_URI = `${ISSUER}/callback`;
const TEST_DIRECTORY = mkdtempSync(join(tmpdir(), "onedrive-oauth-compat-test-"));
const PUBLIC_REFRESH_STORE_FILE = join(TEST_DIRECTORY, "public-refresh-store.json");
const PUBLIC_COMPAT_REFRESH_STORE_FILE = join(
  TEST_DIRECTORY,
  "public-compat-refresh-store.json"
);
const CIMD_REFRESH_STORE_FILE = join(TEST_DIRECTORY, "cimd-refresh-store.json");
const CIMD_CLIENT_ID =
  "https://chatgpt.com/oauth/onedrive-work-pkce/client.json";
const PUBLIC_REFRESH_STORE_KEY_FILE = join(TEST_DIRECTORY, "public-refresh-store.key");
writeFileSync(
  PUBLIC_REFRESH_STORE_KEY_FILE,
  `${Buffer.alloc(32, 0x5a).toString("base64")}\n`,
  { mode: 0o600 }
);
chmodSync(PUBLIC_REFRESH_STORE_KEY_FILE, 0o600);
const testEnv = {
  ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER: ISSUER,
  ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE: RESOURCE,
  ONEDRIVE_OAUTH_COMPAT_CLIENT_ID: CLIENT_ID,
  ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET: CLIENT_SECRET,
  ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_KEY_FILE: PUBLIC_REFRESH_STORE_KEY_FILE,
  ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI: REDIRECT_URI,
  ONEDRIVE_OAUTH_COMPAT_SCOPES: SCOPES.join(" "),
  ONEDRIVE_OAUTH_COMPAT_UPSTREAM_AUTHORIZE_URL: UPSTREAM_AUTHORIZE,
  ONEDRIVE_OAUTH_COMPAT_UPSTREAM_TOKEN_URL: UPSTREAM_TOKEN,
  ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "true",
  ONEDRIVE_OAUTH_COMPAT_HOST: "127.0.0.1",
  ONEDRIVE_OAUTH_COMPAT_PORT: "3010"
};

let passed = 0;
let failed = 0;

function assert(condition, message, details = undefined) {
  if (!condition) {
    const error = new Error(message);
    error.details = details;
    throw error;
  }
}

function test(name, callback) {
  return Promise.resolve()
    .then(callback)
    .then(() => {
      passed += 1;
      process.stdout.write(`PASS ${name}\n`);
    })
    .catch((error) => {
      failed += 1;
      process.stderr.write(`FAIL ${name}: ${error.message}\n`);
      if (error.details) process.stderr.write(`${JSON.stringify(error.details)}\n`);
    });
}

function expectConfigurationError(overrides, expectedText) {
  let thrown;
  try {
    validateOAuthCompatConfiguration({ ...testEnv, ...overrides });
  } catch (error) {
    thrown = error;
  }
  assert(thrown, `Expected configuration failure containing ${expectedText}.`);
  assert(thrown.message.includes(expectedText), "Configuration failed for the wrong reason.", thrown.message);
}

async function listen(server) {
  await new Promise((resolvePromise, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolvePromise);
  });
  return `http://127.0.0.1:${server.address().port}`;
}

async function close(server) {
  await new Promise((resolvePromise) => server.close(resolvePromise));
}

function form(values) {
  return new URLSearchParams(values).toString();
}

function authorizeUrl(base, overrides = {}, additions = []) {
  const url = new URL("/authorize", base);
  const values = {
    resource: RESOURCE,
    client_id: CLIENT_ID,
    redirect_uri: REDIRECT_URI,
    response_type: "code",
    response_mode: "query",
    scope: [...SCOPES].reverse().join(" "),
    state: STATE,
    code_challenge: CHALLENGE,
    code_challenge_method: "S256",
    nonce: "nonce-value",
    prompt: "select_account",
    login_hint: "user@example.test",
    ...overrides
  };
  for (const [name, value] of Object.entries(values)) {
    if (value !== null) url.searchParams.append(name, value);
  }
  for (const [name, value] of additions) url.searchParams.append(name, value);
  return url;
}

function codeToken(overrides = {}) {
  return {
    grant_type: "authorization_code",
    resource: RESOURCE,
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    redirect_uri: REDIRECT_URI,
    code: CODE,
    code_verifier: VERIFIER,
    scope: SCOPES.join(" "),
    ...overrides
  };
}

function refreshToken(overrides = {}) {
  return {
    grant_type: "refresh_token",
    resource: RESOURCE,
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    refresh_token: REFRESH_TOKEN,
    scope: [...SCOPES].reverse().join(" "),
    ...overrides
  };
}

function publicCodeToken(overrides = {}) {
  const values = codeToken();
  delete values.client_secret;
  values.code = DUMMY_FACADE_CODE;
  return { ...values, ...overrides };
}

function publicRefreshToken(overrides = {}) {
  const values = refreshToken();
  delete values.client_secret;
  values.refresh_token = DUMMY_REFRESH_HANDLE;
  return { ...values, ...overrides };
}

async function requestToken(base, values, options = {}) {
  return fetch(`${base}/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      ...(options.headers || {})
    },
    body: typeof values === "string" ? values : form(values),
    ...options,
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      ...(options.headers || {})
    }
  });
}

async function beginPublicAuthorization(base, overrides = {}) {
  const response = await fetch(authorizeUrl(base, overrides), { redirect: "manual" });
  assert(response.status === 302, "Public authorization did not reach the provider.", response.status);
  const location = new URL(response.headers.get("location"));
  return {
    location,
    upstreamState: location.searchParams.get("state")
  };
}

async function finishPublicAuthorization(
  base,
  upstreamState,
  { providerCode = CODE, error = null, errorDescription = null, origin = null } = {}
) {
  const callback = new URL("/callback", base);
  callback.searchParams.set("state", upstreamState);
  if (error) {
    callback.searchParams.set("error", error);
    if (errorDescription) callback.searchParams.set("error_description", errorDescription);
  } else {
    callback.searchParams.set("code", providerCode);
  }
  return fetch(callback, {
    redirect: "manual",
    headers: origin ? { Origin: origin } : {}
  });
}

async function mintPublicFacadeCode(base, overrides = {}) {
  const authorization = await beginPublicAuthorization(base, overrides);
  const callbackResponse = await finishPublicAuthorization(
    base,
    authorization.upstreamState
  );
  assert(callbackResponse.status === 302, "Provider callback did not return to ChatGPT.", callbackResponse.status);
  const outerLocation = new URL(callbackResponse.headers.get("location"));
  const facadeCode = outerLocation.searchParams.get("code");
  assert(/^odac_[A-Za-z0-9_-]{43}$/.test(facadeCode || ""), "Facade code format is invalid.");
  assert(facadeCode !== CODE, "Provider authorization code was exposed.");
  return {
    facadeCode,
    outerLocation,
    upstreamState: authorization.upstreamState,
    providerLocation: authorization.location
  };
}

const upstreamRequests = [];
let upstreamMode = "success";
let upstreamInspection = null;
const mockFetch = async (url, options) => {
  assert(url === UPSTREAM_TOKEN, "Adapter called an unexpected upstream URL.", url);
  assert(options.method === "POST", "Adapter used an unexpected upstream method.");
  assert(options.redirect === "error", "Adapter did not disable upstream redirects.");
  const parameters = new URLSearchParams(options.body);
  upstreamRequests.push(Object.fromEntries(parameters));
  assert(!parameters.has("resource"), "Adapter forwarded the MCP resource parameter to Entra.");
  if (upstreamInspection) {
    const inspect = upstreamInspection;
    upstreamInspection = null;
    inspect(parameters);
  }
  if (upstreamMode === "error") {
    return new Response(JSON.stringify({
      error: "invalid_grant",
      error_description: "The authorization grant is invalid.",
      correlation_id: "safe-test-correlation"
    }), {
      status: 400,
      headers: { "Content-Type": "application/json", "Retry-After": "3" }
    });
  }
  if (upstreamMode === "invalid-json") {
    return new Response("not-json", { status: 502, headers: { "Content-Type": "text/plain" } });
  }
  if (upstreamMode === "oversized") {
    return new Response(JSON.stringify({ padding: "x".repeat(300_000) }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });
  }
  if (upstreamMode === "secret-echo") {
    return new Response(JSON.stringify({
      error: "invalid_client",
      error_description: `Provider reflected ${CLIENT_SECRET}`
    }), {
      status: 400,
      headers: { "Content-Type": "application/json" }
    });
  }
  if (upstreamMode === "refresh-token-echo") {
    return new Response(JSON.stringify({
      token_type: "Bearer",
      access_token: MOCK_PROVIDER_ACCESS_TOKEN,
      refresh_token: PROVIDER_REFRESH_TOKEN,
      provider_note: PROVIDER_REFRESH_TOKEN,
      expires_in: 3600,
      scope: SCOPES.join(" ")
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });
  }
  if (upstreamMode === "provider-scope-only") {
    return new Response(JSON.stringify({
      token_type: "Bearer",
      access_token: MOCK_PROVIDER_ACCESS_TOKEN,
      refresh_token: PROVIDER_REFRESH_TOKEN,
      expires_in: 3600,
      scope: SCOPES[0]
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });
  }
  if (upstreamMode === "network") throw new Error("mock network failure with sensitive detail");
  const isRefresh = parameters.get("grant_type") === "refresh_token";
  return new Response(JSON.stringify({
    token_type: "Bearer",
    access_token: MOCK_PROVIDER_ACCESS_TOKEN,
    refresh_token: PROVIDER_REFRESH_TOKEN,
    expires_in: 3600,
    scope: SCOPES.join(" ")
  }), { status: 200, headers: { "Content-Type": "application/json" } });
};

const server = createOAuthCompatServer(testEnv, {
  fetchImpl: mockFetch,
  rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
  diagnostics: () => {}
});
const base = await listen(server);
const combinedConsentEnv = {
  ...testEnv,
  ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES: COMBINED_CONSENT_SCOPES.join(" ")
};
const combinedConsentServer = createOAuthCompatServer(combinedConsentEnv, {
  fetchImpl: mockFetch,
  rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
  diagnostics: () => {}
});
const combinedConsentBase = await listen(combinedConsentServer);
const publicTestEnv = {
  ...testEnv,
  ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
  ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false",
  ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: PUBLIC_REFRESH_STORE_FILE,
  ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE: PUBLIC_REFRESH_STORE_KEY_FILE,
  ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI: PROVIDER_CALLBACK_URI
};
const publicServer = createOAuthCompatServer(publicTestEnv, {
  fetchImpl: mockFetch,
  rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
  diagnostics: () => {}
});
const publicBase = await listen(publicServer);
const publicCompatTestEnv = {
  ...publicTestEnv,
  ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE: "true",
  ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: PUBLIC_COMPAT_REFRESH_STORE_FILE
};
const publicCompatServer = createOAuthCompatServer(publicCompatTestEnv, {
  fetchImpl: mockFetch,
  rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
  diagnostics: () => {}
});
const publicCompatBase = await listen(publicCompatServer);
const cimdTestEnv = {
  ...publicTestEnv,
  ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD: "true",
  ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR: "true",
  ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: CIMD_REFRESH_STORE_FILE
};
const cimdFetch = async (url, options) => {
  if (url === CIMD_CLIENT_ID) {
    return new Response(JSON.stringify({
      client_id: CIMD_CLIENT_ID,
      redirect_uris: [REDIRECT_URI],
      token_endpoint_auth_methods_supported: ["none", "private_key_jwt"],
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"]
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });
  }
  return mockFetch(url, options);
};
const cimdServer = createOAuthCompatServer(cimdTestEnv, {
  fetchImpl: cimdFetch,
  rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
  diagnostics: () => {}
});
const cimdBase = await listen(cimdServer);

try {
  await test("configuration accepts exact production-shaped allowlist", () => {
    const settings = validateOAuthCompatConfiguration(testEnv);
    assert(settings.clientId === CLIENT_ID);
    assert(settings.protectedResource === RESOURCE);
    assert(settings.redirectUri === REDIRECT_URI);
    assert(settings.authorizeUrl === UPSTREAM_AUTHORIZE);
    assert(settings.tokenUrl === UPSTREAM_TOKEN);
    assert(settings.outerTokenAuthMethod === "client_secret_post");
    assert(settings.allowConfidentialNoPkce === true);
    assert(new Set(settings.scopes).size === SCOPES.length);
    assert(settings.upstreamScopes === settings.scopes);
  });

  await test("configuration accepts only the same API /.default combined-consent translation", () => {
    const settings = validateOAuthCompatConfiguration(combinedConsentEnv);
    assert(JSON.stringify(settings.scopes) === JSON.stringify(SCOPES));
    assert(
      JSON.stringify(settings.upstreamScopes) === JSON.stringify(COMBINED_CONSENT_SCOPES)
    );
  });

  await test("configuration accepts explicit outer public PKCE mode", () => {
    const settings = validateOAuthCompatConfiguration(publicTestEnv);
    assert(settings.outerTokenAuthMethod === "none");
    assert(settings.allowConfidentialNoPkce === false);
    assert(settings.allowPublicNoPkce === false);
    assert(settings.refreshStoreFile === PUBLIC_REFRESH_STORE_FILE);
    assert(settings.refreshStoreEncryptionKey.length === 32);
    assert(settings.providerCallbackUri === PROVIDER_CALLBACK_URI);
  });

  await test("configuration enables CIMD only for a public OAuth client", () => {
    const settings = validateOAuthCompatConfiguration(cimdTestEnv);
    assert(settings.enableCimd === true);
    expectConfigurationError(
      { ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD: "true" },
      "only for a public OAuth client"
    );
  });

  await test("configuration enables DCR only for a public OAuth client", () => {
    const settings = validateOAuthCompatConfiguration(cimdTestEnv);
    assert(settings.enableDcr === true);
    expectConfigurationError(
      { ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR: "true" },
      "only for a public OAuth client"
    );
  });

  await test("configuration defaults confidential non-PKCE compatibility off", () => {
    const env = { ...testEnv };
    delete env.ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE;
    const settings = validateOAuthCompatConfiguration(env);
    assert(settings.allowConfidentialNoPkce === false);
  });

  await test("configuration parses explicit false for confidential non-PKCE compatibility", () => {
    const settings = validateOAuthCompatConfiguration({
      ...testEnv,
      ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false"
    });
    assert(settings.allowConfidentialNoPkce === false);
  });

  await test("configuration accepts explicit public non-PKCE compatibility only in public mode", () => {
    const settings = validateOAuthCompatConfiguration(publicCompatTestEnv);
    assert(settings.outerTokenAuthMethod === "none");
    assert(settings.allowPublicNoPkce === true);
  });

  await test("public HTTP server applies bounded connection lifetimes", () => {
    assert(server.headersTimeout === 10_000, "Header timeout is not bounded.", server.headersTimeout);
    assert(server.requestTimeout === 15_000, "Request timeout is not bounded.", server.requestTimeout);
    assert(server.keepAliveTimeout === 5_000, "Keep-alive timeout is not bounded.", server.keepAliveTimeout);
    assert(server.maxRequestsPerSocket === 100, "Per-socket request count is not bounded.", server.maxRequestsPerSocket);
    assert(server.maxConnections === 256, "Connection count is not bounded.", server.maxConnections);
  });

  for (const [name, overrides, expected] of [
    ["rejects missing issuer", { ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER: "" }, "required"],
    ["rejects HTTP issuer", { ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER: "http://oauth.example.test" }, "must use HTTPS"],
    ["rejects issuer path", { ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER: `${ISSUER}/oauth` }, "origin root"],
    ["rejects placeholder resource", { ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE: "https://REPLACE_WITH_TUNNEL.example" }, "placeholder"],
    ["rejects bad client ID", { ONEDRIVE_OAUTH_COMPAT_CLIENT_ID: "not-a-guid" }, "UUID"],
    ["rejects non-ChatGPT callback", { ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI: "https://evil.example/callback" }, "exact https://chatgpt.com"],
    ["rejects callback query", { ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI: `${REDIRECT_URI}?x=1` }, "query"],
    ["rejects duplicate configured scopes", { ONEDRIVE_OAUTH_COMPAT_SCOPES: "openid openid" }, "unique"],
    ["rejects OIDC scopes", { ONEDRIVE_OAUTH_COMPAT_SCOPES: `${SCOPES.join(" ")} openid profile` }, "must not include OIDC"],
    ["requires refresh scope", { ONEDRIVE_OAUTH_COMPAT_SCOPES: SCOPES.filter((scope) => scope !== "offline_access").join(" ") }, "must include offline_access"],
    [
      "rejects upstream scope escalation to another resource",
      {
        ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES:
          "api://11111111-1111-4111-8111-111111111111/.default offline_access"
      },
      "same API's /.default"
    ],
    [
      "rejects arbitrary upstream delegated scope substitution",
      {
        ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES:
          "api://6e97d01c-edf8-43fe-bf69-bb494ae22513/admin offline_access"
      },
      "same API's /.default"
    ],
    ["rejects alternate authorize endpoint", { ONEDRIVE_OAUTH_COMPAT_UPSTREAM_AUTHORIZE_URL: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize" }, "must equal"],
    ["rejects alternate token endpoint", { ONEDRIVE_OAUTH_COMPAT_UPSTREAM_TOKEN_URL: "https://login.microsoftonline.com/common/oauth2/v2.0/token" }, "must equal"],
    ["rejects placeholder secret", { ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET: "REPLACE_WITH_SECRET" }, "placeholder"],
    ["rejects loose true boolean", { ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "1" }, "exactly"],
    ["rejects uppercase boolean", { ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "TRUE" }, "exactly"],
    ["rejects padded boolean", { ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: " true " }, "exactly"],
    ["rejects empty boolean", { ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "" }, "exactly"],
    ["rejects loose public compatibility boolean", { ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE: "1" }, "exactly"],
    ["rejects public compatibility in confidential mode", { ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE: "true" }, "must be false"],
    ["rejects loose outer auth mode", { ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "public" }, "exactly"],
    ["rejects uppercase outer auth mode", { ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "NONE" }, "exactly"],
    ["rejects empty outer auth mode", { ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "" }, "exactly"],
    [
      "rejects non-PKCE exception in public mode",
      {
        ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
        ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "true"
      },
      "must be false"
    ],
    [
      "requires a public refresh store",
      {
        ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
        ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false"
      },
      "REFRESH_STORE_FILE is required"
    ],
    [
      "rejects a relative public refresh store",
      {
        ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
        ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false",
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: "relative-store.json"
      },
      "normalized absolute"
    ],
    [
      "requires a public refresh-store key",
      {
        ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
        ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false",
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: PUBLIC_REFRESH_STORE_FILE
      },
      "REFRESH_STORE_KEY_FILE is required"
    ],
    [
      "requires the provider callback",
      {
        ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
        ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false",
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: PUBLIC_REFRESH_STORE_FILE,
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE: PUBLIC_REFRESH_STORE_KEY_FILE
      },
      "PROVIDER_CALLBACK_URI is required"
    ],
    [
      "rejects a mismatched provider callback",
      {
        ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD: "none",
        ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE: "false",
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: PUBLIC_REFRESH_STORE_FILE,
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE: PUBLIC_REFRESH_STORE_KEY_FILE,
        ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI: `${ISSUER}/wrong`
      },
      "must exactly equal"
    ]
  ]) {
    await test(`configuration ${name}`, () => expectConfigurationError(overrides, expected));
  }

  await test("health endpoint is public and non-sensitive", async () => {
    const response = await fetch(`${base}/healthz`);
    const body = await response.json();
    assert(response.status === 200, "Health request failed.", response.status);
    assert(body.ok && body.service === "onedrive-oauth-compat", "Health payload is invalid.", body);
    const serialized = JSON.stringify(body);
    assert(!serialized.includes(CLIENT_ID) && !serialized.includes(CLIENT_SECRET), "Health leaked configuration.");
    assert(response.headers.get("cache-control") === "no-store", "Health is cacheable.");
    assert(response.headers.get("x-content-type-options") === "nosniff", "nosniff is missing.");
  });

  await test("health exposes only a redacted MCP diagnostic projection", async () => {
    const diagnosticServer = createOAuthCompatServer(testEnv, {
      fetchImpl: async (url, options) => {
        if (url === "http://127.0.0.1:3001/healthz") {
          return new Response(JSON.stringify({
            ok: true,
            authMode: "oauth",
            lastAuthFailure: {
              at: "2026-07-27T14:40:00.000Z",
              authMode: "oauth_error",
              code: "invalid_token",
              status: 401,
              message: "sensitive bearer detail"
            },
            lastToolFailure: {
              at: "2026-07-27T14:41:00.000Z",
              tool: "search",
              code: "graph_error",
              graphStatus: 403,
              message: "sensitive Graph detail"
            },
            lastToolCall: {
              at: "2026-07-27T14:41:00.000Z",
              tool: "search",
              isError: true
            }
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        return mockFetch(url, options);
      },
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: () => {}
    });
    const diagnosticBase = await listen(diagnosticServer);
    try {
      const response = await fetch(`${diagnosticBase}/healthz`);
      const body = await response.json();
      assert(body.mcp?.lastAuthFailure?.status === 401, "MCP auth status is missing.", body);
      assert(body.mcp?.lastToolFailure?.graphStatus === 403, "MCP Graph status is missing.", body);
      assert(body.mcp?.lastToolCall?.tool === "search", "MCP tool diagnostic is missing.", body);
      const serialized = JSON.stringify(body);
      assert(!serialized.includes("sensitive"), "MCP health projection leaked diagnostic text.", body);
    } finally {
      await close(diagnosticServer);
    }
  });

  await test("health reports only redacted diagnostics for the latest OAuth request", async () => {
    const authorize = await fetch(
      authorizeUrl(base, { scope: `${SCOPES.join(" ")} unexpected` }),
      { redirect: "manual" }
    );
    assert(authorize.status === 302, "Trusted authorize error did not use the callback.");
    const response = await fetch(`${base}/healthz`);
    const body = await response.json();
    assert(
      body.lastOAuthRequest?.route === "/authorize"
        && body.lastOAuthRequest?.status === 302
        && body.lastOAuthRequest?.code === "invalid_scope",
      "Health did not expose the latest redacted OAuth diagnostic.",
      body
    );
    const serialized = JSON.stringify(body);
    assert(
      !serialized.includes(CLIENT_ID)
        && !serialized.includes(CLIENT_SECRET)
        && !serialized.includes("authorization_code"),
      "Health leaked OAuth configuration or credentials.",
      body
    );
  });

  await test("cheap shared-loopback routes have no persistent shared-IP lockout", async () => {
    const cheapServer = createOAuthCompatServer(testEnv, {
      fetchImpl: mockFetch,
      rateLimits: {
        health: 1,
        metadata: 1,
        other: 1,
        authorize: 500,
        callback: 500,
        token: 500
      },
      diagnostics: () => {}
    });
    const cheapBase = await listen(cheapServer);
    try {
      for (let attempt = 0; attempt < 4; attempt += 1) {
        const health = await fetch(`${cheapBase}/healthz`);
        const metadata = await fetch(
          `${cheapBase}/.well-known/oauth-authorization-server`
        );
        const missing = await fetch(`${cheapBase}/missing`);
        assert(
          health.status === 200
            && metadata.status === 200
            && missing.status === 404,
          "Cheap route acquired a shared caller lockout.",
          { health: health.status, metadata: metadata.status, missing: missing.status }
        );
      }
    } finally {
      await close(cheapServer);
    }
  });

  for (const path of [
    "/.well-known/oauth-authorization-server",
    "/.well-known/openid-configuration"
  ]) {
    await test(`metadata ${path} advertises strict OAuth capabilities`, async () => {
      const response = await fetch(`${base}${path}`, {
        headers: { Origin: "https://chatgpt.com" }
      });
      const body = await response.json();
      assert(response.status === 200, "Metadata request failed.", response.status);
      assert(body.issuer === ISSUER, "Metadata issuer mismatch.", body);
      assert(body.authorization_endpoint === `${ISSUER}/authorize`);
      assert(body.token_endpoint === `${ISSUER}/token`);
      assert(body.jwks_uri === `${ISSUER}/jwks.json`);
      assert(JSON.stringify(body.response_types_supported) === '["code"]');
      assert(JSON.stringify(body.code_challenge_methods_supported) === '["S256"]');
      assert(JSON.stringify(body.token_endpoint_auth_methods_supported) === '["client_secret_post"]');
      assert(
        JSON.stringify(body.scopes_supported) === JSON.stringify(SCOPES),
        "Authorization metadata scope set drifted.",
        body.scopes_supported
      );
      assert(
        !Object.hasOwn(body, "allowConfidentialNoPkce")
          && !Object.hasOwn(body, "allow_confidential_no_pkce"),
        "Confidential non-PKCE compatibility leaked into public metadata.",
        body
      );
      assert(response.headers.get("access-control-allow-origin") === "https://chatgpt.com");
    });
  }

  await test("JWKS publishes the facade ES256 verification key", async () => {
    const response = await fetch(`${base}/jwks.json`);
    const body = await response.json();
    const key = body.keys?.[0];
    assert(response.status === 200, "JWKS request failed.", response.status);
    assert(
      body.keys?.length === 1
        && key.kty === "EC"
        && key.crv === "P-256"
        && key.alg === "ES256"
        && key.use === "sig"
        && /^[A-Za-z0-9_-]+$/.test(key.kid || "")
        && /^[A-Za-z0-9_-]+$/.test(key.x || "")
        && /^[A-Za-z0-9_-]+$/.test(key.y || "")
        && !Object.hasOwn(key, "d"),
      "JWKS did not expose exactly one public ES256 facade key.",
      body
    );
  });

  await test("public-mode metadata advertises token auth none and S256 only", async () => {
    const response = await fetch(`${publicBase}/.well-known/oauth-authorization-server`);
    const body = await response.json();
    assert(response.status === 200, "Public-mode metadata request failed.", response.status);
    assert(JSON.stringify(body.token_endpoint_auth_methods_supported) === '["none"]', body);
    assert(JSON.stringify(body.code_challenge_methods_supported) === '["S256"]', body);
    const serialized = JSON.stringify(body);
    assert(!serialized.includes(CLIENT_SECRET), "Public metadata leaked the upstream secret.", body);
    assert(!serialized.includes(PUBLIC_REFRESH_STORE_FILE), "Public metadata leaked the store path.", body);
    assert(!serialized.includes(PUBLIC_REFRESH_STORE_KEY_FILE), "Public metadata leaked the key path.", body);
    assert(!serialized.includes(PROVIDER_CALLBACK_URI), "Public metadata leaked the provider callback.", body);
  });

  await test("CIMD metadata and a verified ChatGPT client document enable public PKCE", async () => {
    const metadataResponse = await fetch(
      `${cimdBase}/.well-known/oauth-authorization-server`
    );
    const metadata = await metadataResponse.json();
    assert(
      metadata.client_id_metadata_document_supported === true,
      "CIMD support was not advertised.",
      metadata
    );
    const authorization = await fetch(
      authorizeUrl(cimdBase, { client_id: CIMD_CLIENT_ID }),
      { redirect: "manual" }
    );
    assert(authorization.status === 302, "Verified CIMD authorization failed.");
    const provider = new URL(authorization.headers.get("location"));
    assert(
      provider.origin + provider.pathname === UPSTREAM_AUTHORIZE,
      "Verified CIMD authorization did not reach Microsoft.",
      provider.toString()
    );
    assert(
      provider.searchParams.get("client_id") === CLIENT_ID,
      "CIMD client identity leaked into the private Microsoft client leg.",
      provider.toString()
    );
  });

  await test("DCR registers only signed public PKCE clients", async () => {
    const metadataResponse = await fetch(
      `${cimdBase}/.well-known/oauth-authorization-server`
    );
    const metadata = await metadataResponse.json();
    assert(
      metadata.registration_endpoint === `${ISSUER}/register`,
      "DCR endpoint was not advertised.",
      metadata
    );
    const response = await fetch(`${cimdBase}/register`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        client_name: "ChatGPT",
        redirect_uris: [REDIRECT_URI],
        token_endpoint_auth_method: "none",
        grant_types: ["authorization_code", "refresh_token"],
        response_types: ["code"],
        scope: SCOPES.join(" ")
      })
    });
    const body = await response.json();
    assert(response.status === 201, "DCR registration failed.", body);
    assert(
      /^oddc_[A-Za-z0-9_-]{22}\.[A-Za-z0-9_-]{43}$/.test(body.client_id || ""),
      "DCR returned an invalid signed client ID.",
      body
    );
    assert(body.token_endpoint_auth_method === "none", "DCR returned a client secret method.", body);
    assert(!Object.hasOwn(body, "client_secret"), "DCR exposed a client secret.", body);
    const authorization = await fetch(
      authorizeUrl(cimdBase, { client_id: body.client_id }),
      { redirect: "manual" }
    );
    assert(
      authorization.status === 302,
      "The signed DCR client was not accepted for PKCE authorization."
    );
    const tamperedClientId =
      `${body.client_id.slice(0, -1)}${body.client_id.endsWith("A") ? "B" : "A"}`;
    const tamperedAuthorization = await fetch(
      authorizeUrl(cimdBase, { client_id: tamperedClientId }),
      { redirect: "manual" }
    );
    assert(
      tamperedAuthorization.status === 400,
      "A tampered DCR client ID was accepted."
    );
  });

  await test("DCR rejects untrusted redirect metadata", async () => {
    const response = await fetch(`${cimdBase}/register`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        redirect_uris: ["https://attacker.example.test/callback"],
        token_endpoint_auth_method: "none",
        response_types: ["code"]
      })
    });
    assert(response.status === 400, "DCR accepted an untrusted redirect URI.");
  });

  await test("CIMD rejects non-ChatGPT client metadata URLs", async () => {
    const authorization = await fetch(
      authorizeUrl(cimdBase, {
        client_id: "https://attacker.example.test/oauth/client.json"
      }),
      { redirect: "manual" }
    );
    assert(authorization.status === 400, "Untrusted CIMD client was accepted.");
  });

  await test("public no-PKCE compatibility metadata does not advertise outer PKCE", async () => {
    const response = await fetch(`${publicCompatBase}/.well-known/oauth-authorization-server`);
    const body = await response.json();
    assert(response.status === 200, "Public compatibility metadata request failed.", response.status);
    assert(JSON.stringify(body.token_endpoint_auth_methods_supported) === '["none"]', body);
    assert(
      !Object.hasOwn(body, "code_challenge_methods_supported"),
      "Public no-PKCE compatibility metadata falsely advertised outer PKCE.",
      body
    );
  });

  await test("OpenID discovery aliases OAuth metadata without advertising ID tokens", async () => {
    const response = await fetch(`${base}/.well-known/openid-configuration`);
    const body = await response.json();
    assert(response.status === 200 && body.issuer === ISSUER, body);
    assert(
      !Object.hasOwn(body, "id_token_signing_alg_values_supported")
        && !Object.hasOwn(body, "claims_supported"),
      "OAuth compatibility service incorrectly claims to issue OIDC ID tokens.",
      body
    );
  });

  await test("authorize strips resource, canonicalizes scope, and preserves other allowed parameters", async () => {
    const input = authorizeUrl(base, {
      claims: JSON.stringify({ id_token: { email: null } }),
      max_age: "0",
      ui_locales: "en-US"
    });
    const response = await fetch(input, { redirect: "manual" });
    assert(response.status === 302, "Authorize did not redirect.", response.status);
    const location = new URL(response.headers.get("location"));
    assert(`${location.origin}${location.pathname}` === UPSTREAM_AUTHORIZE, "Wrong authorize target.", location.toString());
    assert(!location.searchParams.has("resource"), "Resource was not stripped.");
    for (const [name, value] of input.searchParams) {
      if (name === "scope") {
        assert(
          location.searchParams.get(name) === SCOPES.join(" "),
          "Authorize did not canonicalize the configured upstream scope set."
        );
      } else if (name !== "resource") {
        assert(location.searchParams.get(name) === value, `Authorize changed ${name}.`);
      }
    }
    assert(response.headers.get("cache-control") === "no-store");
  });

  await test("authorize accepts confidential flow without PKCE and preserves its absence", async () => {
    const input = authorizeUrl(base, {
      code_challenge: null,
      code_challenge_method: null
    });
    const response = await fetch(input, { redirect: "manual" });
    assert(response.status === 302, "Confidential non-PKCE authorization did not redirect.", response.status);
    const location = new URL(response.headers.get("location"));
    assert(`${location.origin}${location.pathname}` === UPSTREAM_AUTHORIZE, "Wrong authorize target.", location.toString());
    assert(!location.searchParams.has("code_challenge"), "Adapter added a PKCE challenge.");
    assert(!location.searchParams.has("code_challenge_method"), "Adapter added a PKCE method.");
  });

  await test("authorize accepts and preserves complete S256 PKCE", async () => {
    const response = await fetch(authorizeUrl(base), { redirect: "manual" });
    assert(response.status === 302, "S256 PKCE authorization did not redirect.", response.status);
    const location = new URL(response.headers.get("location"));
    assert(location.searchParams.get("code_challenge") === CHALLENGE, "PKCE challenge was not preserved.");
    assert(location.searchParams.get("code_challenge_method") === "S256", "PKCE method was not preserved.");
  });

  await test("public mode accepts and preserves complete S256 PKCE", async () => {
    const response = await fetch(authorizeUrl(publicBase), { redirect: "manual" });
    assert(response.status === 302, "Public-mode S256 authorization did not redirect.", response.status);
    const location = new URL(response.headers.get("location"));
    assert(`${location.origin}${location.pathname}` === UPSTREAM_AUTHORIZE, "Wrong authorize target.", location.toString());
    assert(location.searchParams.get("code_challenge") === CHALLENGE, "Public-mode challenge was not preserved.");
    assert(location.searchParams.get("code_challenge_method") === "S256", "Public-mode method was not preserved.");
    assert(location.searchParams.get("redirect_uri") === PROVIDER_CALLBACK_URI, "Provider callback was not substituted.");
    assert(/^odst_[A-Za-z0-9_-]{43}$/.test(location.searchParams.get("state") || ""), "Upstream state is invalid.");
    assert(location.searchParams.get("state") !== STATE, "Outer state was exposed upstream.");
  });

  for (const [name, overrides] of [
    ["missing PKCE pair", { code_challenge: null, code_challenge_method: null }],
    ["challenge without method", { code_challenge_method: null }],
    ["method without challenge", { code_challenge: null }],
    ["non-S256 method", { code_challenge_method: "plain" }],
    ["invalid challenge", { code_challenge: "short" }]
  ]) {
    await test(`public mode rejects ${name}`, async () => {
      const response = await fetch(authorizeUrl(publicBase, overrides), { redirect: "manual" });
      assert(response.status === 302, "Trusted public-mode error did not use the callback.", response.status);
      const location = new URL(response.headers.get("location"));
      assert(
        `${location.origin}${location.pathname}` === REDIRECT_URI
          && location.searchParams.get("error") === "invalid_request",
        "Public-mode PKCE failure was not rejected safely.",
        location.toString()
      );
    });
  }

  await test("public compatibility mode protects the provider leg with generated S256 PKCE", async () => {
    upstreamMode = "success";
    const authorization = await beginPublicAuthorization(publicCompatBase, {
      code_challenge: null,
      code_challenge_method: null
    });
    const generatedChallenge = authorization.location.searchParams.get("code_challenge");
    assert(
      /^[A-Za-z0-9_-]{43}$/.test(generatedChallenge || ""),
      "Compatibility authorization did not generate a provider PKCE challenge."
    );
    assert(
      authorization.location.searchParams.get("code_challenge_method") === "S256",
      "Compatibility authorization did not force provider S256."
    );
    assert(
      authorization.location.searchParams.get("redirect_uri") === PROVIDER_CALLBACK_URI,
      "Compatibility authorization used the wrong provider callback."
    );

    const callbackResponse = await finishPublicAuthorization(
      publicCompatBase,
      authorization.upstreamState
    );
    assert(callbackResponse.status === 302, "Compatibility callback did not return to ChatGPT.");
    const outerLocation = new URL(callbackResponse.headers.get("location"));
    const facadeCode = outerLocation.searchParams.get("code");
    assert(
      /^odac_[A-Za-z0-9_-]{43}$/.test(facadeCode || ""),
      "Compatibility callback did not issue a facade code."
    );

    const tokenValues = publicCodeToken({ code: facadeCode });
    delete tokenValues.code_verifier;
    const response = await requestToken(publicCompatBase, tokenValues);
    const body = await response.json();
    assert(response.status === 200, "Compatibility token exchange failed.", body);
    assert(
      /^odrh_[A-Za-z0-9_-]{43}$/.test(body.refresh_token || ""),
      "Compatibility token exchange did not issue an opaque refresh handle."
    );
    assert(
      body.scope === SCOPES.join(" "),
      "Compatibility token response did not preserve the outer scope contract.",
      body.scope
    );
    const upstream = upstreamRequests.at(-1);
    assert(
      /^[A-Za-z0-9._~-]{43,128}$/.test(upstream.code_verifier || ""),
      "Generated provider verifier was not restored for token exchange."
    );
    assert(
      createHash("sha256")
        .update(upstream.code_verifier, "ascii")
        .digest("base64url") === generatedChallenge,
      "Generated provider PKCE verifier did not match its authorization challenge."
    );
    assert(upstream.code === CODE, "Compatibility mapping did not restore the provider code.");
    assert(upstream.client_secret === CLIENT_SECRET, "Compatibility exchange omitted provider proof.");
  });

  await test("MCP challenge advertises only resource scopes while the facade requests refresh consent", async () => {
    const challenge = oauthChallenge({
      env: {
        ONEDRIVE_MCP_AUTH_MODE: "oauth",
        ONEDRIVE_MCP_OAUTH_API_SCOPE: API_SCOPE,
        ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET: "challenge-test-client-secret",
        ONEDRIVE_MCP_RESOURCE_METADATA_URL:
          "https://onedrive-tunnel.example.test/.well-known/oauth-protected-resource/v1/mcp/tunnel_test"
      }
    });
    const challengeScope = challenge.match(/\bscope="([^"]+)"/)?.[1] || "";
    assert(challengeScope === API_SCOPE, "MCP challenge advertised a non-resource scope.", challenge);
    const response = await fetch(authorizeUrl(base, { scope: SCOPES.join(" ") }), { redirect: "manual" });
    assert(response.status === 302, "Facade authorization did not redirect.", response.status);
    const location = new URL(response.headers.get("location"));
    assert(
      `${location.origin}${location.pathname}` === UPSTREAM_AUTHORIZE,
      "Challenge-derived authorization did not reach Microsoft.",
      location.toString()
    );
  });

  await test("combined consent preserves the outer scope contract and requests API /.default upstream", async () => {
    const response = await fetch(authorizeUrl(combinedConsentBase), {
      redirect: "manual"
    });
    assert(response.status === 302, "Combined-consent authorization did not redirect.");
    const location = new URL(response.headers.get("location"));
    assert(
      location.searchParams.get("scope") === COMBINED_CONSENT_SCOPES.join(" "),
      "Combined-consent authorization forwarded the wrong upstream scopes.",
      location.toString()
    );
    const tokenResponse = await requestToken(combinedConsentBase, codeToken());
    assert(tokenResponse.status === 200, "Combined-consent token exchange failed.", await tokenResponse.text());
    const upstream = upstreamRequests.at(-1);
    assert(
      upstream.scope === COMBINED_CONSENT_SCOPES.join(" "),
      "Combined-consent token exchange forwarded the wrong upstream scopes.",
      upstream
    );
    const refreshResponse = await requestToken(
      combinedConsentBase,
      refreshToken()
    );
    assert(refreshResponse.status === 200, "Combined-consent refresh failed.", await refreshResponse.text());
    assert(
      upstreamRequests.at(-1).scope === COMBINED_CONSENT_SCOPES.join(" "),
      "Combined-consent refresh forwarded the wrong upstream scopes.",
      upstreamRequests.at(-1)
    );
  });

  await test("authorize accepts and preserves bounded opaque state", async () => {
    for (const state of ["x", "opaque+state=with.padding", "s".repeat(512)]) {
      const response = await fetch(authorizeUrl(base, { state }), { redirect: "manual" });
      assert(response.status === 302, "Bounded opaque state was rejected.", { stateLength: state.length });
      const location = new URL(response.headers.get("location"));
      assert(location.searchParams.get("state") === state, "OAuth state was not preserved exactly.");
    }
    const oversized = await fetch(
      authorizeUrl(base, { state: "s".repeat(513) }),
      { redirect: "manual" }
    );
    const oversizedLocation = new URL(oversized.headers.get("location"));
    assert(
      oversized.status === 302
        && oversizedLocation.searchParams.get("error") === "invalid_request"
        && !oversizedLocation.searchParams.has("state"),
      "Oversized OAuth state was not rejected safely.",
      oversizedLocation.toString()
    );
  });

  await test("authorization-code token exchange validates and strips only resource", async () => {
    upstreamMode = "success";
    const response = await requestToken(base, codeToken());
    const body = await response.json();
    assert(response.status === 200, "Token exchange failed.", body);
    const [facadeHeaderValue, facadeClaimsValue] = body.access_token.split(".");
    const facadeHeader = JSON.parse(Buffer.from(facadeHeaderValue, "base64url").toString("utf8"));
    const facadeClaims = JSON.parse(Buffer.from(facadeClaimsValue, "base64url").toString("utf8"));
    assert(
      facadeHeader.alg === "ES256"
        && facadeHeader.typ === "at+jwt"
        && facadeClaims.aud === RESOURCE
        && facadeClaims.azp === CLIENT_ID
        && facadeClaims.client_id === CLIENT_ID
        && facadeClaims.scp === SCOPES.join(" ")
        && facadeClaims.scope === SCOPES.join(" "),
      "Facade JWT did not expose the resource-routing claims required by ChatGPT Work.",
      { facadeHeader, facadeClaims }
    );
    assert(
      !JSON.stringify({ facadeHeader, facadeClaims }).includes(MOCK_PROVIDER_ACCESS_TOKEN),
      "Facade JWT exposed the Microsoft access token in readable claims."
    );
    const verifiedFacade = verifyFacadeAccessToken(body.access_token, {
      issuer: ISSUER,
      audience: RESOURCE,
      clientId: CLIENT_ID,
      requiredScope: API_SCOPE,
      keyFile: PUBLIC_REFRESH_STORE_KEY_FILE
    });
    assert(
      verifiedFacade.upstreamAccessToken === MOCK_PROVIDER_ACCESS_TOKEN,
      "Facade token did not preserve the Microsoft access token securely."
    );
    const upstream = upstreamRequests.at(-1);
    assert(!Object.hasOwn(upstream, "resource"), "Resource reached provider.", upstream);
    for (const [name, value] of Object.entries(codeToken())) {
      if (name !== "resource") assert(upstream[name] === value, `Token proxy changed ${name}.`);
    }
    assert(response.headers.get("cache-control") === "no-store");
  });

  await test("authorization-code token exchange accepts omitted verifier and preserves its absence", async () => {
    upstreamMode = "success";
    const values = codeToken();
    delete values.code_verifier;
    const response = await requestToken(base, values);
    const body = await response.json();
    assert(response.status === 200, "Confidential non-PKCE token exchange failed.", body);
    const upstream = upstreamRequests.at(-1);
    assert(!Object.hasOwn(upstream, "code_verifier"), "Adapter added a PKCE verifier.", upstream);
  });

  await test("captured ChatGPT confidential flow succeeds without an outer PKCE pair", async () => {
    upstreamMode = "success";
    const authorization = authorizeUrl(base, {
      code_challenge: null,
      code_challenge_method: null
    });
    const authorizeResponse = await fetch(authorization, { redirect: "manual" });
    assert(
      authorizeResponse.status === 302,
      "Captured ChatGPT authorization shape did not reach Microsoft.",
      authorizeResponse.status
    );
    const authorizeLocation = new URL(authorizeResponse.headers.get("location"));
    assert(
      `${authorizeLocation.origin}${authorizeLocation.pathname}` === UPSTREAM_AUTHORIZE,
      "Captured ChatGPT authorization shape used the wrong provider.",
      authorizeLocation.toString()
    );
    assert(
      authorizeLocation.searchParams.get("redirect_uri") === REDIRECT_URI,
      "Captured ChatGPT authorization shape changed the exact callback.",
      authorizeLocation.toString()
    );
    assert(
      !authorizeLocation.searchParams.has("code_challenge")
        && !authorizeLocation.searchParams.has("code_challenge_method"),
      "Captured ChatGPT authorization shape unexpectedly gained PKCE parameters.",
      authorizeLocation.toString()
    );

    const tokenValues = codeToken();
    delete tokenValues.code_verifier;
    delete tokenValues.scope;
    const tokenResponse = await requestToken(base, tokenValues);
    const tokenBody = await tokenResponse.json();
    assert(tokenResponse.status === 200, "Captured ChatGPT token shape failed.", tokenBody);
    assert(
      tokenBody.scope === API_SCOPE,
      "Captured ChatGPT token response advertised a scope that was not requested.",
      tokenBody.scope
    );
    const upstream = upstreamRequests.at(-1);
    assert(upstream.client_secret === CLIENT_SECRET, "Confidential client proof was not forwarded.");
    assert(!Object.hasOwn(upstream, "resource"), "MCP resource leaked to Microsoft.", upstream);
    assert(!Object.hasOwn(upstream, "code_verifier"), "Adapter invented an outer verifier.", upstream);
  });

  await test("confidential action-scoped reconnect accepts only the resource scope", async () => {
    upstreamMode = "success";
    const authorizeResponse = await fetch(
      authorizeUrl(base, { scope: API_SCOPE, code_challenge: null, code_challenge_method: null }),
      { redirect: "manual" }
    );
    assert(
      authorizeResponse.status === 302,
      "Action-scoped authorization did not reach Microsoft.",
      authorizeResponse.status
    );
    const authorizeLocation = new URL(authorizeResponse.headers.get("location"));
    assert(
      authorizeLocation.searchParams.get("scope") === SCOPES.join(" "),
      "Action-scoped authorization did not retain upstream refresh consent.",
      authorizeLocation.toString()
    );
    const tokenValues = codeToken({ scope: API_SCOPE });
    delete tokenValues.code_verifier;
    const tokenResponse = await requestToken(base, tokenValues);
    const tokenBody = await tokenResponse.json();
    assert(tokenResponse.status === 200, "Action-scoped token exchange failed.", tokenBody);
    assert(
      tokenBody.scope === API_SCOPE,
      "Action-scoped token response did not preserve the requested outer scope.",
      tokenBody.scope
    );
    assert(
      tokenBody.refresh_token === PROVIDER_REFRESH_TOKEN,
      "Action-scoped token response did not preserve refresh capability."
    );
  });

  await test("confidential token response preserves the exact outer scope contract", async () => {
    upstreamMode = "provider-scope-only";
    const response = await requestToken(base, codeToken());
    const body = await response.json();
    assert(response.status === 200, "Confidential token exchange failed.", body);
    assert(
      body.scope === SCOPES.join(" "),
      "Confidential token response did not preserve the outer scope contract.",
      body.scope
    );
    assert(
      body.refresh_token === PROVIDER_REFRESH_TOKEN,
      "Confidential token response did not preserve the provider refresh token."
    );
    upstreamMode = "success";
  });

  let firstPublicRefreshHandle;
  let rotatedPublicRefreshHandle;
  let consumedFacadeCode;

  await test("public callback maps provider code and preserves only the outer state", async () => {
    const minted = await mintPublicFacadeCode(publicBase);
    assert(
      `${minted.outerLocation.origin}${minted.outerLocation.pathname}` === REDIRECT_URI,
      "Facade callback used the wrong ChatGPT redirect.",
      minted.outerLocation.toString()
    );
    assert(minted.outerLocation.searchParams.get("state") === STATE, "Outer state was not restored.");
    assert(!minted.outerLocation.toString().includes(CODE), "Provider code leaked through the callback.");

    const replay = await finishPublicAuthorization(publicBase, minted.upstreamState);
    assert(replay.status === 400, "Consumed provider state was accepted twice.", replay.status);
  });

  await test("public callback relays only sanitized provider OAuth errors", async () => {
    const authorization = await beginPublicAuthorization(publicBase);
    const response = await finishPublicAuthorization(
      publicBase,
      authorization.upstreamState,
      {
        providerCode: null,
        error: "access_denied",
        errorDescription: `Denied\u0000 ${CLIENT_SECRET}`,
        origin: "https://login.microsoftonline.com"
      }
    );
    assert(response.status === 302, "Provider error did not return to ChatGPT.", response.status);
    const location = new URL(response.headers.get("location"));
    assert(location.searchParams.get("error") === "access_denied", "Provider OAuth error was not preserved.");
    assert(location.searchParams.get("state") === STATE, "Outer state was not restored on error.");
    assert(!location.searchParams.has("error_description"), "Secret-bearing provider description was relayed.");
    assert(!location.toString().includes(CLIENT_SECRET), "Provider error leaked the configured secret.");
    assert(!location.toString().includes(authorization.upstreamState), "Upstream state leaked to ChatGPT.");
  });

  await test("public callback preserves form_post without exposing provider code or state", async () => {
    const authorization = await beginPublicAuthorization(publicBase, {
      response_mode: "form_post"
    });
    assert(
      authorization.location.searchParams.get("response_mode") === "query",
      "Provider callback was not forced to bounded query mode."
    );
    const rejectedPost = await fetch(new URL("/callback", publicBase), {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Origin: "https://login.microsoftonline.com"
      },
      body: `padding=${"x".repeat(40_000)}`,
      redirect: "manual"
    });
    assert(
      rejectedPost.status === 405 && rejectedPost.headers.get("allow") === "GET",
      "Callback POST body was accepted.",
      rejectedPost.status
    );
    const response = await finishPublicAuthorization(
      publicBase,
      authorization.upstreamState,
      { origin: "https://login.microsoftonline.com" }
    );
    const body = await response.text();
    assert(response.status === 200, "form_post callback failed.", response.status);
    assert(body.includes(`action="${REDIRECT_URI}"`), "form_post used the wrong ChatGPT callback.");
    assert(body.includes(`name="state" value="${STATE}"`), "form_post omitted outer state.");
    assert(/name="code" value="odac_[A-Za-z0-9_-]{43}"/.test(body), "form_post omitted facade code.");
    assert(!body.includes(CODE), "form_post exposed the provider code.");
    assert(!body.includes(authorization.upstreamState), "form_post exposed upstream state.");
  });

  await test("public callback rejects unknown state without redirecting", async () => {
    const response = await finishPublicAuthorization(
      publicBase,
      `odst_${"A".repeat(43)}`
    );
    assert(response.status === 400, "Unknown provider state was redirected.", response.status);
    assert(!response.headers.has("location"), "Unknown provider state received a trusted callback.");
  });

  await test("public-mode authorization-code exchange consumes mapped code and returns only an opaque refresh handle", async () => {
    upstreamMode = "success";
    const minted = await mintPublicFacadeCode(publicBase);
    const outer = publicCodeToken({ code: minted.facadeCode });
    consumedFacadeCode = minted.facadeCode;
    assert(!Object.hasOwn(outer, "client_secret"), "Public request fixture contains a client secret.");
    const response = await requestToken(publicBase, outer);
    const body = await response.json();
    assert(response.status === 200, "Public-mode code exchange failed.", body);
    assert(/^odrh_[A-Za-z0-9_-]{43}$/.test(body.refresh_token || ""), "Outer refresh handle is invalid.");
    assert(body.refresh_token !== PROVIDER_REFRESH_TOKEN, "Provider refresh token was exposed.");
    assert(body.scope === SCOPES.join(" "), "Outer token scope was not canonicalized.", body.scope);
    firstPublicRefreshHandle = body.refresh_token;
    const upstream = upstreamRequests.at(-1);
    assert(upstream.client_secret === CLIENT_SECRET, "Configured Entra secret was not injected upstream.");
    assert(upstream.code_verifier === VERIFIER, "PKCE verifier was not forwarded upstream.");
    assert(upstream.code === CODE, "Mapped provider code was not restored upstream.");
    assert(upstream.redirect_uri === PROVIDER_CALLBACK_URI, "Provider callback was not used for token exchange.");
    assert(!Object.hasOwn(upstream, "resource"), "Resource reached Microsoft.", upstream);

    const persisted = readFileSync(PUBLIC_REFRESH_STORE_FILE, "utf8");
    assert(!persisted.includes(PROVIDER_REFRESH_TOKEN), "Provider refresh token was stored in plaintext.");
    assert(!persisted.includes(firstPublicRefreshHandle), "Outer refresh handle was stored in plaintext.");
    assert(!persisted.includes(CODE), "Provider authorization code was stored in plaintext.");
    assert((statSync(PUBLIC_REFRESH_STORE_FILE).mode & 0o077) === 0, "Refresh store is not mode 0600.");
  });

  await test("public-mode refresh atomically rotates its handle and forces exact scope", async () => {
    upstreamMode = "success";
    const outer = publicRefreshToken({ refresh_token: firstPublicRefreshHandle });
    delete outer.scope;
    let observedConsumedBeforeUpstream = false;
    upstreamInspection = () => {
      const persisted = JSON.parse(readFileSync(PUBLIC_REFRESH_STORE_FILE, "utf8"));
      observedConsumedBeforeUpstream = persisted.entries.some(
        (entry) => entry.status === "consumed"
      );
    };
    assert(!Object.hasOwn(outer, "client_secret"), "Public refresh fixture contains a client secret.");
    assert(!Object.hasOwn(outer, "code_verifier"), "Public refresh fixture unexpectedly contains a verifier.");
    let response;
    try {
      response = await requestToken(publicBase, outer);
    } finally {
      upstreamInspection = null;
    }
    const body = await response.json();
    assert(response.status === 200, "Public-mode refresh exchange failed.", body);
    assert(/^odrh_[A-Za-z0-9_-]{43}$/.test(body.refresh_token || ""), "Rotated refresh handle is invalid.");
    assert(body.refresh_token !== firstPublicRefreshHandle, "Refresh handle did not rotate.");
    assert(body.refresh_token !== PROVIDER_REFRESH_TOKEN, "Provider refresh token was exposed on rotation.");
    rotatedPublicRefreshHandle = body.refresh_token;
    const upstream = upstreamRequests.at(-1);
    assert(upstream.client_secret === CLIENT_SECRET, "Configured Entra secret was not injected on refresh.");
    assert(upstream.refresh_token === PROVIDER_REFRESH_TOKEN, "Encrypted provider refresh token was not restored.");
    assert(upstream.scope === SCOPES.join(" "), "Exact configured scope was not forced upstream.");
    assert(!Object.hasOwn(upstream, "resource"), "Resource reached Microsoft.", upstream);
    assert(observedConsumedBeforeUpstream, "Refresh handle was not durably consumed before upstream I/O.");
  });

  await test("refresh-handle replay revokes every active descendant in its family", async () => {
    const beforeReplay = upstreamRequests.length;
    const replay = await requestToken(
      publicBase,
      publicRefreshToken({ refresh_token: firstPublicRefreshHandle })
    );
    assert(replay.status === 400, "Consumed refresh handle replay was accepted.", await replay.text());
    assert(upstreamRequests.length === beforeReplay, "Replay reached Microsoft.");

    const descendant = await requestToken(
      publicBase,
      publicRefreshToken({ refresh_token: rotatedPublicRefreshHandle })
    );
    assert(descendant.status === 400, "Descendant survived family replay revocation.", await descendant.text());
    assert(upstreamRequests.length === beforeReplay, "Revoked descendant reached Microsoft.");
  });

  await test("facade authorization codes are one-time and reject direct provider-code injection", async () => {
    const before = upstreamRequests.length;
    const replay = await requestToken(
      publicBase,
      publicCodeToken({ code: consumedFacadeCode })
    );
    assert(replay.status === 400, "Facade authorization code replay was accepted.", await replay.text());
    const direct = await requestToken(
      publicBase,
      publicCodeToken({ code: CODE })
    );
    assert(direct.status === 400, "Direct provider code injection was accepted.", await direct.text());
    assert(upstreamRequests.length === before, "Rejected authorization code reached Microsoft.");
  });

  await test("facade authorization codes cannot be cross-swapped with another PKCE verifier", async () => {
    const alternate = await mintPublicFacadeCode(publicBase, {
      code_challenge: ALTERNATE_CHALLENGE
    });
    const before = upstreamRequests.length;
    const swapped = await requestToken(
      publicBase,
      publicCodeToken({ code: alternate.facadeCode, code_verifier: VERIFIER })
    );
    assert(swapped.status === 400, "Cross-swapped PKCE verifier was accepted.", await swapped.text());
    const retry = await requestToken(
      publicBase,
      publicCodeToken({
        code: alternate.facadeCode,
        code_verifier: ALTERNATE_VERIFIER
      })
    );
    assert(retry.status === 400, "Failed PKCE attempt did not consume the facade code.", await retry.text());
    assert(upstreamRequests.length === before, "Cross-swapped facade code reached Microsoft.");
  });

  await test("authorization transactions, facade codes, and refresh handles survive secure reloads", async () => {
    upstreamMode = "success";
    const persistentEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "persistent-public-store.json"
      )
    };
    const serverOptions = {
      fetchImpl: mockFetch,
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: () => {}
    };
    let activeServer;
    try {
      activeServer = createOAuthCompatServer(persistentEnv, serverOptions);
      const authorizeBase = await listen(activeServer);
      const authorization = await beginPublicAuthorization(authorizeBase);
      await close(activeServer);
      activeServer = null;

      activeServer = createOAuthCompatServer(persistentEnv, serverOptions);
      const callbackBase = await listen(activeServer);
      const callbackResponse = await finishPublicAuthorization(
        callbackBase,
        authorization.upstreamState
      );
      assert(callbackResponse.status === 302, "Reloaded callback transaction failed.");
      const facadeCode = new URL(
        callbackResponse.headers.get("location")
      ).searchParams.get("code");
      await close(activeServer);
      activeServer = null;

      activeServer = createOAuthCompatServer(persistentEnv, serverOptions);
      const codeBase = await listen(activeServer);
      const codeResponse = await requestToken(
        codeBase,
        publicCodeToken({ code: facadeCode })
      );
      const codeBody = await codeResponse.json();
      assert(codeResponse.status === 200, "Reloaded facade code exchange failed.", codeBody);
      const refreshHandle = codeBody.refresh_token;
      await close(activeServer);
      activeServer = null;

      const rotatedClientSecret = "rotated-oauth-compat-test-secret-value";
      activeServer = createOAuthCompatServer({
        ...persistentEnv,
        ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET: rotatedClientSecret
      }, serverOptions);
      const refreshBase = await listen(activeServer);
      const refreshResponse = await requestToken(
        refreshBase,
        publicRefreshToken({ refresh_token: refreshHandle })
      );
      const refreshBody = await refreshResponse.json();
      assert(refreshResponse.status === 200, "Reloaded refresh handle failed.", refreshBody);
      assert(
        refreshBody.refresh_token !== refreshHandle
          && /^odrh_[A-Za-z0-9_-]{43}$/.test(refreshBody.refresh_token || ""),
        "Reloaded refresh handle did not rotate."
      );
      assert(
        upstreamRequests.at(-1).client_secret === rotatedClientSecret,
        "Stable vault key did not preserve the session across OAuth secret rotation."
      );
      await close(activeServer);
      activeServer = null;
    } finally {
      if (activeServer) await close(activeServer);
    }
  });

  await test("authorization transactions, facade codes, and refresh handles expire at bounded TTLs", async () => {
    upstreamMode = "success";
    let currentTime = 2_000_000_000_000;
    const ttlEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "ttl-public-store.json"
      )
    };
    const ttlServer = createOAuthCompatServer(ttlEnv, {
      fetchImpl: mockFetch,
      now: () => currentTime,
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: () => {},
      refreshStoreLimits: {
        authorizationTtlMs: 10,
        codeTtlMs: 10,
        refreshTtlMs: 10
      }
    });
    const ttlBase = await listen(ttlServer);
    try {
      const expiringAuthorization = await beginPublicAuthorization(ttlBase);
      currentTime += 11;
      const expiredCallback = await finishPublicAuthorization(
        ttlBase,
        expiringAuthorization.upstreamState
      );
      assert(expiredCallback.status === 400, "Expired authorization transaction was accepted.");

      const expiringCode = await mintPublicFacadeCode(ttlBase);
      currentTime += 11;
      const beforeCode = upstreamRequests.length;
      const expiredCode = await requestToken(
        ttlBase,
        publicCodeToken({ code: expiringCode.facadeCode })
      );
      assert(expiredCode.status === 400, "Expired facade code was accepted.", await expiredCode.text());
      assert(upstreamRequests.length === beforeCode, "Expired facade code reached Microsoft.");

      const liveCode = await mintPublicFacadeCode(ttlBase);
      const codeResponse = await requestToken(
        ttlBase,
        publicCodeToken({ code: liveCode.facadeCode })
      );
      const codeBody = await codeResponse.json();
      assert(codeResponse.status === 200, "Live code failed in TTL test.", codeBody);
      currentTime += 11;
      const beforeRefresh = upstreamRequests.length;
      const expiredRefresh = await requestToken(
        ttlBase,
        publicRefreshToken({ refresh_token: codeBody.refresh_token })
      );
      assert(expiredRefresh.status === 400, "Expired refresh handle was accepted.", await expiredRefresh.text());
      assert(upstreamRequests.length === beforeRefresh, "Expired refresh handle reached Microsoft.");
    } finally {
      await close(ttlServer);
    }
  });

  await test("refresh store fails closed on malformed, oversized, or unauthenticated ciphertext", async () => {
    const corruptPath = join(TEST_DIRECTORY, "corrupt-public-store.json");
    writeFileSync(corruptPath, "{not-json", { mode: 0o600 });
    chmodSync(corruptPath, 0o600);
    let malformedFailure;
    try {
      createOAuthCompatServer({
        ...publicTestEnv,
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: corruptPath
      }, { fetchImpl: mockFetch, diagnostics: () => {} });
    } catch (error) {
      malformedFailure = error;
    }
    assert(malformedFailure, "Malformed refresh store was accepted.");
    assert(!malformedFailure.message.includes(corruptPath), "Corruption error leaked the store path.");

    const symlinkPath = join(TEST_DIRECTORY, "symlink-public-store.json");
    symlinkSync(corruptPath, symlinkPath);
    let symlinkFailure;
    try {
      createOAuthCompatServer({
        ...publicTestEnv,
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: symlinkPath
      }, { fetchImpl: mockFetch, diagnostics: () => {} });
    } catch (error) {
      symlinkFailure = error;
    }
    assert(symlinkFailure, "Symlink refresh store was accepted.");

    const oversizedPath = join(TEST_DIRECTORY, "oversized-public-store.json");
    writeFileSync(oversizedPath, "x".repeat(129), { mode: 0o600 });
    chmodSync(oversizedPath, 0o600);
    let oversizedFailure;
    try {
      createOAuthCompatServer({
        ...publicTestEnv,
        ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: oversizedPath
      }, {
        fetchImpl: mockFetch,
        diagnostics: () => {},
        refreshStoreLimits: { maxBytes: 128 }
      });
    } catch (error) {
      oversizedFailure = error;
    }
    assert(oversizedFailure, "Oversized refresh store was accepted.");

    const tamperedPath = join(TEST_DIRECTORY, "tampered-public-store.json");
    const tamperedEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: tamperedPath
    };
    const tamperServer = createOAuthCompatServer(tamperedEnv, {
      fetchImpl: mockFetch,
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: () => {}
    });
    const tamperBase = await listen(tamperServer);
    const minted = await mintPublicFacadeCode(tamperBase);
    const tokenResponse = await requestToken(
      tamperBase,
      publicCodeToken({ code: minted.facadeCode })
    );
    assert(tokenResponse.status === 200, "Could not seed encrypted store for tamper test.");
    await close(tamperServer);
    const tampered = JSON.parse(readFileSync(tamperedPath, "utf8"));
    const activeEntry = tampered.entries.find((entry) => entry.status === "active");
    activeEntry.tag = `${activeEntry.tag[0] === "A" ? "B" : "A"}${activeEntry.tag.slice(1)}`;
    writeFileSync(tamperedPath, `${JSON.stringify(tampered)}\n`, { mode: 0o600 });
    chmodSync(tamperedPath, 0o600);
    let tamperedFailure;
    try {
      createOAuthCompatServer(tamperedEnv, {
        fetchImpl: mockFetch,
        diagnostics: () => {}
      });
    } catch (error) {
      tamperedFailure = error;
    }
    assert(tamperedFailure, "Tampered encrypted refresh token was accepted.");
  });

  await test("refresh-store entry bounds fail closed without exposing a provider token", async () => {
    upstreamMode = "success";
    const boundedEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "bounded-entry-store.json"
      )
    };
    const boundedServer = createOAuthCompatServer(boundedEnv, {
      fetchImpl: mockFetch,
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: () => {},
      refreshStoreLimits: {
        maxEntries: 1,
        maxFamilies: 1
      }
    });
    const boundedBase = await listen(boundedServer);
    try {
      const first = await mintPublicFacadeCode(boundedBase);
      const firstResponse = await requestToken(
        boundedBase,
        publicCodeToken({ code: first.facadeCode })
      );
      assert(firstResponse.status === 200, "First bounded refresh family failed.");

      const second = await mintPublicFacadeCode(boundedBase);
      const secondResponse = await requestToken(
        boundedBase,
        publicCodeToken({ code: second.facadeCode })
      );
      const secondText = await secondResponse.text();
      assert(secondResponse.status === 503, "Refresh-store entry bound did not fail closed.");
      assert(!secondText.includes(PROVIDER_REFRESH_TOKEN), "Bound failure exposed provider refresh token.");
    } finally {
      await close(boundedServer);
    }
  });

  await test("authorize pair and route budgets resist distinct max-state floods and refill", async () => {
    let currentTime = 2_100_000_000_000;
    const floodStore = join(TEST_DIRECTORY, "authorize-route-flood-store.json");
    const floodServer = createOAuthCompatServer({
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: floodStore
    }, {
      fetchImpl: mockFetch,
      now: () => currentTime,
      rateLimits: { authorize: 1, callback: 500, token: 500 },
      rateKeyCapacity: 64,
      diagnostics: () => {}
    });
    const floodBase = await listen(floodServer);
    const maxState = (index) => {
      const prefix = `state-${String(index).padStart(6, "0")}-`;
      return `${prefix}${"s".repeat(512 - prefix.length)}`;
    };
    const challenge = (index) =>
      createHash("sha256")
        .update(`authorize-flood-${index}`, "utf8")
        .digest("base64url");
    const authorizeFloodRequest = async (stateIndex, challengeIndex = stateIndex) => {
      const response = await fetch(
        authorizeUrl(floodBase, {
          state: maxState(stateIndex),
          code_challenge: challenge(challengeIndex)
        }),
        { redirect: "manual" }
      );
      const location = new URL(response.headers.get("location"));
      return {
        response,
        location,
        provider:
          `${location.origin}${location.pathname}` === UPSTREAM_AUTHORIZE
      };
    };
    try {
      let accepted = 0;
      const first = await authorizeFloodRequest(0);
      assert(first.provider, "First max-state authorization was not admitted.");
      accepted += 1;

      const repeatedPair = await authorizeFloodRequest(0);
      assert(
        !repeatedPair.provider
          && repeatedPair.location.searchParams.get("error") === "temporarily_unavailable",
        "Repeated state/challenge pair was not throttled.",
        repeatedPair.location.toString()
      );
      const differentState = await authorizeFloodRequest(1, 0);
      const differentChallenge = await authorizeFloodRequest(0, 1);
      assert(
        differentState.provider && differentChallenge.provider,
        "Authorization throttle did not bind both state and PKCE challenge."
      );
      accepted += 2;

      for (let index = 2; index < 2082; index += 1) {
        const result = await authorizeFloodRequest(index);
        if (result.provider) accepted += 1;
      }
      assert(accepted === 32, "Distinct-key flood exceeded the route burst.", accepted);
      let stored = JSON.parse(readFileSync(floodStore, "utf8"));
      assert(
        stored.authorizations.length === 32,
        "Distinct-key flood filled unexpected authorization records.",
        stored.authorizations.length
      );

      const stillExhausted = await authorizeFloodRequest(3000);
      assert(!stillExhausted.provider, "Route burst recovered without refill time.");

      for (let step = 1; step <= 119; step += 1) {
        currentTime += 5_000;
        const recovered = await authorizeFloodRequest(3000 + step);
        assert(
          recovered.provider,
          "Continuously refilling authorize bucket did not recover in five seconds.",
          { step, status: recovered.response.status }
        );
      }
      stored = JSON.parse(readFileSync(floodStore, "utf8"));
      assert(
        stored.authorizations.length === 151,
        "Max-state transactions exceeded their projected live envelope.",
        stored.authorizations.length
      );
      assert(
        Buffer.byteLength(JSON.stringify({
          version: stored.version,
          authorizations: stored.authorizations,
          codes: stored.codes
        }), "utf8") < 640 * 1024,
        "Max-state authorization envelope exceeded the ephemeral partition."
      );
      assert(
        Buffer.byteLength(JSON.stringify(stored), "utf8") < 1024 * 1024,
        "Max-state authorization envelope exceeded the total store."
      );

      const noStickyWindow = await authorizeFloodRequest(5000);
      assert(!noStickyWindow.provider, "Exhausted route admitted without another refill.");
      currentTime += 5_000;
      const afterExpiry = await authorizeFloodRequest(5001);
      assert(
        afterExpiry.provider,
        "Legitimate authorization did not recover when the next token refilled."
      );
      stored = JSON.parse(readFileSync(floodStore, "utf8"));
      assert(
        stored.authorizations.length < 152,
        "Expired transactions were not reclaimed below the admission ceiling.",
        stored.authorizations.length
      );
    } finally {
      await close(floodServer);
    }
  });

  await test("max callback-code flood cannot consume the refresh partition", async () => {
    upstreamMode = "success";
    let currentTime = 2_200_000_000_000;
    const isolationStore = join(TEST_DIRECTORY, "callback-refresh-isolation-store.json");
    const isolationServer = createOAuthCompatServer({
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: isolationStore
    }, {
      fetchImpl: mockFetch,
      now: () => currentTime,
      rateLimits: { authorize: 500, callback: 500, token: 500 },
      diagnostics: () => {}
    });
    const isolationBase = await listen(isolationServer);
    try {
      const legitimateCode = await mintPublicFacadeCode(isolationBase);
      const initialToken = await requestToken(
        isolationBase,
        publicCodeToken({ code: legitimateCode.facadeCode })
      );
      const initialBody = await initialToken.json();
      assert(initialToken.status === 200, "Could not establish refresh family.", initialBody);
      const liveRefreshHandle = initialBody.refresh_token;

      const attackStates = [];
      for (let index = 0; index < 64; index += 1) {
        const state = `attack-state-${index}`;
        const codeChallenge = createHash("sha256")
          .update(`callback-attack-${index}`, "utf8")
          .digest("base64url");
        const response = await fetch(
          authorizeUrl(isolationBase, { state, code_challenge: codeChallenge }),
          { redirect: "manual" }
        );
        const location = new URL(response.headers.get("location"));
        if (`${location.origin}${location.pathname}` === UPSTREAM_AUTHORIZE) {
          attackStates.push(location.searchParams.get("state"));
        }
      }
      assert(attackStates.length === 31, "Authorize route admitted the wrong attack burst.", attackStates.length);

      const oversized = await finishPublicAuthorization(
        isolationBase,
        attackStates[0],
        { providerCode: "\\".repeat(3 * 1024 + 1) }
      );
      assert(oversized.status === 400, "Oversized provider code was accepted.");

      const maximumProviderCode = "\\".repeat(3 * 1024);
      let admittedCallbacks = 0;
      let stateIndex = 0;
      while (stateIndex < attackStates.length && admittedCallbacks < 15) {
        const response = await finishPublicAuthorization(
          isolationBase,
          attackStates[stateIndex],
          { providerCode: maximumProviderCode }
        );
        if (response.status === 302) admittedCallbacks += 1;
        stateIndex += 1;
      }
      assert(admittedCallbacks === 15, "Callback route admitted the wrong initial burst.", admittedCallbacks);

      for (let step = 1; step <= 14; step += 1) {
        currentTime += 20_000;
        const response = await finishPublicAuthorization(
          isolationBase,
          attackStates[stateIndex],
          { providerCode: maximumProviderCode }
        );
        assert(response.status === 302, "Callback route did not refill continuously.", { step, status: response.status });
        admittedCallbacks += 1;
        stateIndex += 1;
      }
      const throttledCallback = await finishPublicAuthorization(
        isolationBase,
        attackStates[stateIndex],
        { providerCode: maximumProviderCode }
      );
      assert(throttledCallback.status === 429, "Callback admission ceiling did not throttle.");
      assert(admittedCallbacks === 29, "Unexpected callback-code attack envelope.", admittedCallbacks);

      const beforeRefresh = upstreamRequests.length;
      const rotated = await requestToken(
        isolationBase,
        publicRefreshToken({ refresh_token: liveRefreshHandle })
      );
      const rotatedBody = await rotated.json();
      assert(rotated.status === 200, "Ephemeral pressure blocked refresh rotation.", rotatedBody);
      assert(
        /^odrh_[A-Za-z0-9_-]{43}$/.test(rotatedBody.refresh_token || "")
          && rotatedBody.refresh_token !== liveRefreshHandle,
        "Refresh handle did not rotate under ephemeral pressure."
      );
      assert(upstreamRequests.length === beforeRefresh + 1, "Callback flood reached Microsoft token exchange.");

      const stored = JSON.parse(readFileSync(isolationStore, "utf8"));
      const refreshBytes = Buffer.byteLength(JSON.stringify({
        version: stored.version,
        families: stored.families,
        entries: stored.entries
      }), "utf8");
      const ephemeralBytes = Buffer.byteLength(JSON.stringify({
        version: stored.version,
        authorizations: stored.authorizations,
        codes: stored.codes
      }), "utf8");
      assert(refreshBytes < 320 * 1024, "Refresh partition exceeded its reserved budget.", refreshBytes);
      assert(ephemeralBytes < 640 * 1024, "Ephemeral partition exceeded its independent budget.", ephemeralBytes);
    } finally {
      await close(isolationServer);
    }
  });

  await test("unsafe transaction, code, and ephemeral-byte envelopes fail at startup", () => {
    for (const [name, refreshStoreLimits, expected] of [
      ["authorization count", { maxAuthorizations: 1 }, "transaction count"],
      ["authorization-code count", { maxCodes: 1 }, "authorization-code count"],
      ["ephemeral bytes", { maxEphemeralBytes: 200 * 1024 }, "ephemeral refresh-store byte"]
    ]) {
      let failure;
      try {
        createOAuthCompatServer({
          ...publicTestEnv,
          ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
            TEST_DIRECTORY,
            `unsafe-${name.replaceAll(" ", "-")}-store.json`
          )
        }, {
          fetchImpl: mockFetch,
          diagnostics: () => {},
          refreshStoreLimits
        });
      } catch (error) {
        failure = error;
      }
      assert(failure, `Unsafe ${name} envelope was accepted.`);
      assert(
        failure.message.includes(expected),
        `Unsafe ${name} envelope failed for the wrong reason.`,
        failure.message
      );
    }
  });

  await test("public token limits use validated opaque keys with bounded LRU eviction", async () => {
    const rateEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "public-opaque-key-rate-store.json"
      )
    };
    const rateDiagnostics = [];
    const rateServer = createOAuthCompatServer(rateEnv, {
      fetchImpl: mockFetch,
      rateLimits: { token: 1 },
      rateKeyCapacity: 2,
      diagnostics: (entry) => rateDiagnostics.push(entry)
    });
    const rateBase = await listen(rateServer);
    try {
      const before = upstreamRequests.length;
      const malformed = await fetch(`${rateBase}/token`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: "{}"
      });
      assert(malformed.status === 415, "Malformed public token body had wrong status.");
      const first = await requestToken(rateBase, publicCodeToken());
      const repeated = await requestToken(rateBase, publicCodeToken());
      assert(first.status === 400, "Malformed body poisoned the opaque code bucket.", await first.text());
      assert(repeated.status === 429, "Repeated opaque code was not throttled.", await repeated.text());

      const distinctCodes = Array.from({ length: 8 }, (_, index) =>
        `odac_${createHash("sha256")
          .update(`rate-capacity-${index}`, "utf8")
          .digest("base64url")}`
      );
      for (const code of distinctCodes) {
        const response = await requestToken(
          rateBase,
          publicCodeToken({ code })
        );
        assert(
          response.status === 400,
          "A random distinct key caused a global rate-limit failure.",
          { codeStatus: response.status }
        );
      }
      const lastRepeated = await requestToken(
        rateBase,
        publicCodeToken({ code: distinctCodes.at(-1) })
      );
      assert(lastRepeated.status === 429, "LRU did not retain the most recent key.");
      assert(upstreamRequests.length === before, "Rejected opaque-key requests reached Microsoft.");
      const serializedDiagnostics = JSON.stringify(rateDiagnostics);
      for (const code of [DUMMY_FACADE_CODE, ...distinctCodes]) {
        assert(!serializedDiagnostics.includes(code), "Diagnostics exposed a rate-limit key.");
      }
    } finally {
      await close(rateServer);
    }
  });

  await test("public token request bodies remain strictly bounded", async () => {
    const bodyEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "public-body-bound-store.json"
      )
    };
    const bodyServer = createOAuthCompatServer(bodyEnv, {
      fetchImpl: mockFetch,
      rateLimits: { token: 500 },
      diagnostics: () => {}
    });
    const bodyBase = await listen(bodyServer);
    try {
      const response = await fetch(`${bodyBase}/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: `padding=${"x".repeat(40_000)}`
      });
      assert(response.status === 413, "Oversized public token body was accepted.");
    } finally {
      await close(bodyServer);
    }
  });

  await test("token pre-parse concurrency rejects overload and releases its slot", async () => {
    const concurrencyEnv = {
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "public-token-concurrency-store.json"
      )
    };
    let releaseProvider = () => {};
    let markProviderEntered;
    const providerEntered = new Promise((resolvePromise) => {
      markProviderEntered = resolvePromise;
    });
    let providerCalls = 0;
    const providerGate = new Promise((resolvePromise) => {
      releaseProvider = resolvePromise;
    });
    const blockingFetch = async (...arguments_) => {
      providerCalls += 1;
      if (providerCalls === 1) {
        markProviderEntered();
        await providerGate;
      }
      return mockFetch(...arguments_);
    };
    const concurrencyServer = createOAuthCompatServer(concurrencyEnv, {
      fetchImpl: blockingFetch,
      rateLimits: { authorize: 500, callback: 500, token: 500 },
      tokenConcurrency: 1,
      diagnostics: () => {}
    });
    const concurrencyBase = await listen(concurrencyServer);
    try {
      const firstCode = await mintPublicFacadeCode(concurrencyBase);
      const secondCode = await mintPublicFacadeCode(concurrencyBase, {
        state: "second-concurrency-state",
        code_challenge: ALTERNATE_CHALLENGE
      });
      const firstPromise = requestToken(
        concurrencyBase,
        publicCodeToken({ code: firstCode.facadeCode })
      );
      await providerEntered;
      const overloaded = await requestToken(
        concurrencyBase,
        publicCodeToken({
          code: secondCode.facadeCode,
          code_verifier: ALTERNATE_VERIFIER
        })
      );
      assert(overloaded.status === 503, "Concurrent token overload was not rejected.");
      assert(overloaded.headers.get("retry-after") === "1", "Overload response lacks Retry-After.");
      assert(providerCalls === 1, "Overloaded token request reached Microsoft.");

      releaseProvider();
      const first = await firstPromise;
      assert(first.status === 200, "First concurrent token request failed.", await first.text());
      const retried = await requestToken(
        concurrencyBase,
        publicCodeToken({
          code: secondCode.facadeCode,
          code_verifier: ALTERNATE_VERIFIER
        })
      );
      assert(retried.status === 200, "Released token slot was not reusable.", await retried.text());
      assert(providerCalls === 2, "Retried token request did not reach Microsoft exactly once.");
    } finally {
      releaseProvider();
      await close(concurrencyServer);
    }
  });

  await test("public mode never reflects an upstream secret echo", async () => {
    upstreamMode = "secret-echo";
    const minted = await mintPublicFacadeCode(publicBase);
    const response = await requestToken(
      publicBase,
      publicCodeToken({ code: minted.facadeCode })
    );
    const text = await response.text();
    assert(response.status === 502, "Secret-bearing provider response was not rejected.", response.status);
    assert(JSON.parse(text).error === "server_error", "Secret-bearing provider response was not generic.", text);
    assert(!text.includes(CLIENT_SECRET), "Configured Entra secret was reflected to the public client.");
    upstreamMode = "success";
  });

  await test("public mode never exposes a provider refresh token from any response field", async () => {
    upstreamMode = "refresh-token-echo";
    const minted = await mintPublicFacadeCode(publicBase);
    const response = await requestToken(
      publicBase,
      publicCodeToken({ code: minted.facadeCode })
    );
    const text = await response.text();
    assert(response.status === 502, "Provider refresh-token echo was not rejected.", response.status);
    assert(!text.includes(PROVIDER_REFRESH_TOKEN), "Provider refresh token escaped the facade.");
    upstreamMode = "success";
  });

  const publicTokenFailures = [
    ["authorization-code client secret", publicCodeToken({ client_secret: CLIENT_SECRET }), {}, 400],
    ["empty authorization-code client secret", publicCodeToken({ client_secret: "" }), {}, 400],
    ["missing verifier", (() => { const value = publicCodeToken(); delete value.code_verifier; return value; })(), {}, 400],
    ["invalid verifier", publicCodeToken({ code_verifier: "short" }), {}, 400],
    ["wrong client", publicCodeToken({ client_id: "11111111-1111-4111-8111-111111111111" }), {}, 400],
    ["missing client", (() => { const value = publicCodeToken(); delete value.client_id; return value; })(), {}, 400],
    ["wrong callback", publicCodeToken({ redirect_uri: `${REDIRECT_URI}x` }), {}, 400],
    ["wrong resource", publicCodeToken({ resource: `${RESOURCE}/wrong` }), {}, 400],
    ["broadened scope", publicCodeToken({ scope: `${SCOPES.join(" ")} Files.Read` }), {}, 400],
    ["basic authentication", publicCodeToken(), { headers: { Authorization: "Basic dGVzdDp0ZXN0" } }, 401],
    ["refresh client secret", publicRefreshToken({ client_secret: CLIENT_SECRET }), {}, 400],
    ["empty refresh client secret", publicRefreshToken({ client_secret: "" }), {}, 400],
    ["wrong refresh client", publicRefreshToken({ client_id: "11111111-1111-4111-8111-111111111111" }), {}, 400],
    ["wrong refresh resource", publicRefreshToken({ resource: `${RESOURCE}/wrong` }), {}, 400],
    ["broadened refresh scope", publicRefreshToken({ scope: `${SCOPES.join(" ")} Files.Read` }), {}, 400],
    ["missing refresh token", (() => { const value = publicRefreshToken(); delete value.refresh_token; return value; })(), {}, 400]
  ];
  for (const [name, values, options, status] of publicTokenFailures) {
    await test(`public mode rejects ${name}`, async () => {
      const before = upstreamRequests.length;
      const response = await requestToken(publicBase, values, options);
      const text = await response.text();
      assert(response.status === status, "Unsafe public token request had wrong status.", {
        status: response.status,
        text
      });
      assert(upstreamRequests.length === before, "Unsafe public token request reached Microsoft.");
      assert(!text.includes(CLIENT_SECRET), "Public token error leaked the configured Entra secret.");
    });
  }

  await test("default-off server rejects non-PKCE flows while accepting complete S256 PKCE", async () => {
    upstreamMode = "success";
    const env = { ...testEnv };
    delete env.ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE;
    const defaultOffServer = createOAuthCompatServer(env, {
      fetchImpl: mockFetch,
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: () => {}
    });
    const defaultOffBase = await listen(defaultOffServer);
    try {
      const noPkceAuthorize = await fetch(authorizeUrl(defaultOffBase, {
        code_challenge: null,
        code_challenge_method: null
      }), { redirect: "manual" });
      assert(noPkceAuthorize.status === 302, "Trusted missing-PKCE error did not use the callback.");
      const noPkceLocation = new URL(noPkceAuthorize.headers.get("location"));
      assert(
        `${noPkceLocation.origin}${noPkceLocation.pathname}` === REDIRECT_URI
          && noPkceLocation.searchParams.get("error") === "invalid_request",
        "Default-off authorization did not reject missing PKCE safely.",
        noPkceLocation.toString()
      );

      const validAuthorize = await fetch(authorizeUrl(defaultOffBase), { redirect: "manual" });
      assert(validAuthorize.status === 302, "Default-off server rejected complete S256 PKCE.");
      const validLocation = new URL(validAuthorize.headers.get("location"));
      assert(
        `${validLocation.origin}${validLocation.pathname}` === UPSTREAM_AUTHORIZE,
        "Complete S256 PKCE did not reach Microsoft.",
        validLocation.toString()
      );

      const withoutVerifier = codeToken();
      delete withoutVerifier.code_verifier;
      const beforeRejectedToken = upstreamRequests.length;
      const rejectedToken = await requestToken(defaultOffBase, withoutVerifier);
      assert(rejectedToken.status === 400, "Default-off token endpoint accepted an omitted verifier.");
      assert(
        upstreamRequests.length === beforeRejectedToken,
        "Default-off omitted verifier reached Microsoft."
      );

      const validToken = await requestToken(defaultOffBase, codeToken());
      assert(validToken.status === 200, "Default-off server rejected a valid PKCE verifier.", await validToken.text());
    } finally {
      await close(defaultOffServer);
    }
  });

  await test("refresh token exchange validates and strips only resource", async () => {
    upstreamMode = "success";
    const response = await requestToken(base, refreshToken());
    const body = await response.json();
    assert(response.status === 200, "Refresh failed.", body);
    const verifiedFacade = verifyFacadeAccessToken(body.access_token, {
      issuer: ISSUER,
      audience: RESOURCE,
      clientId: CLIENT_ID,
      requiredScope: API_SCOPE,
      keyFile: PUBLIC_REFRESH_STORE_KEY_FILE
    });
    assert(
      verifiedFacade.upstreamAccessToken === MOCK_PROVIDER_ACCESS_TOKEN,
      "Refresh facade token did not preserve the Microsoft access token securely."
    );
    const upstream = upstreamRequests.at(-1);
    assert(upstream.grant_type === "refresh_token");
    assert(!Object.hasOwn(upstream, "resource"));
  });

  await test("token scope may be omitted but may not be broadened", async () => {
    upstreamMode = "success";
    const values = codeToken({ scope: undefined });
    delete values.scope;
    const response = await requestToken(base, values);
    assert(response.status === 200, "Omitted token scope should be accepted.", await response.text());
  });

  await test("provider OAuth JSON error and status are preserved safely", async () => {
    upstreamMode = "error";
    const response = await requestToken(base, codeToken({ code: "provider-error-code" }));
    const body = await response.json();
    assert(response.status === 400, "Provider status was not preserved.", response.status);
    assert(body.error === "invalid_grant", "Provider OAuth error was not preserved.", body);
    assert(body.correlation_id === "safe-test-correlation", "Provider JSON was not preserved.", body);
    assert(response.headers.get("retry-after") === "3", "Safe Retry-After was not preserved.");
    upstreamMode = "success";
  });

  const untrustedAuthorizeFailures = [
    ["wrong client", { client_id: "11111111-1111-4111-8111-111111111111" }, []],
    ["missing client", { client_id: null }, []],
    ["wrong callback", { redirect_uri: `${REDIRECT_URI}x` }, []],
    ["missing callback", { redirect_uri: null }, []],
    ["duplicate client ID", {}, [["client_id", CLIENT_ID]]],
    ["duplicate callback", {}, [["redirect_uri", REDIRECT_URI]]]
  ];
  for (const [name, overrides, additions] of untrustedAuthorizeFailures) {
    await test(`authorize directly rejects untrusted ${name}`, async () => {
      const response = await fetch(authorizeUrl(base, overrides, additions), { redirect: "manual" });
      const body = await response.json();
      assert(
        response.status === 400 && body.error === "invalid_request",
        "Untrusted authorize target was redirected.",
        { status: response.status, body }
      );
      assert(!response.headers.has("location"), "Untrusted authorize target received a callback redirect.");
    });
  }

  const trustedAuthorizeFailures = [
    ["wrong resource", { resource: `${RESOURCE}/wrong` }, [], "invalid_request", true],
    ["implicit response", { response_type: "token" }, [], "invalid_request", true],
    ["plain PKCE", { code_challenge_method: "plain" }, [], "invalid_request", true],
    ["challenge without PKCE method", { code_challenge_method: null }, [], "invalid_request", true],
    ["PKCE method without challenge", { code_challenge: null }, [], "invalid_request", true],
    ["short PKCE", { code_challenge: "short" }, [], "invalid_request", true],
    ["empty state", { state: "" }, [], "invalid_request", false],
    ["oversized state", { state: "s".repeat(9000) }, [], "invalid_request", false],
    ["extra scope", { scope: `${SCOPES.join(" ")} Files.Read` }, [], "invalid_scope", true],
    ["missing scope", { scope: null }, [], "invalid_scope", true],
    ["unknown parameter", {}, [["client_secret", CLIENT_SECRET]], "invalid_request", true],
    ["duplicate state", {}, [["state", "second-state"]], "invalid_request", false],
    ["fragment response mode", { response_mode: "fragment" }, [], "invalid_request", true],
    ["invalid claims JSON", { claims: "not-json" }, [], "invalid_request", true],
    ["prompt none combination", { prompt: "none login" }, [], "invalid_request", true]
  ];
  for (const [name, overrides, additions, expectedError, expectState] of trustedAuthorizeFailures) {
    await test(`authorize returns trusted ${name} error to ChatGPT`, async () => {
      const response = await fetch(authorizeUrl(base, overrides, additions), { redirect: "manual" });
      assert(response.status === 302, "Trusted authorize error did not use the callback.", response.status);
      const location = new URL(response.headers.get("location"));
      assert(
        `${location.origin}${location.pathname}` === REDIRECT_URI,
        "Trusted authorize error used the wrong callback.",
        location.toString()
      );
      assert(location.searchParams.get("error") === expectedError, "Wrong callback OAuth error.", location.toString());
      assert(
        expectState ? location.searchParams.get("state") === STATE : !location.searchParams.has("state"),
        "Callback state handling was incorrect.",
        location.toString()
      );
      assert(!location.toString().includes(CLIENT_SECRET), "Authorize callback leaked a secret.");
    });
  }

  await test("authorize returns form_post errors safely with exact state", async () => {
    const state = "opaque+state=with.padding";
    const response = await fetch(authorizeUrl(base, {
      response_mode: "form_post",
      scope: `${SCOPES.join(" ")} Files.Read`,
      state
    }), { redirect: "manual" });
    const body = await response.text();
    assert(response.status === 200, "form_post authorization error had the wrong status.", response.status);
    assert(
      response.headers.get("content-type")?.startsWith("text/html"),
      "form_post authorization error was not HTML."
    );
    assert(body.includes(`action="${REDIRECT_URI}"`), "form_post used the wrong callback.", body);
    assert(body.includes('name="error" value="invalid_scope"'), "form_post omitted the OAuth error.", body);
    assert(body.includes(`name="state" value="${state}"`), "form_post did not preserve state.", body);
    assert(
      /script-src 'nonce-[A-Za-z0-9_-]+'/.test(response.headers.get("content-security-policy") || ""),
      "form_post script is not nonce-bound."
    );
    assert(!body.includes(CLIENT_SECRET), "form_post error leaked a secret.");
  });

  const tokenFailures = [
    ["wrong resource", codeToken({ resource: `${RESOURCE}/wrong` }), {}, 400],
    ["wrong client", codeToken({ client_id: "11111111-1111-4111-8111-111111111111" }), {}, 400],
    ["wrong client secret", codeToken({ client_secret: "wrong-secret" }), {}, 401],
    ["missing callback", (() => { const value = codeToken(); delete value.redirect_uri; return value; })(), {}, 400],
    ["wrong callback", codeToken({ redirect_uri: `${REDIRECT_URI}x` }), {}, 400],
    ["missing code", (() => { const value = codeToken(); delete value.code; return value; })(), {}, 400],
    ["bad verifier", codeToken({ code_verifier: "short" }), {}, 400],
    ["broadened scope", codeToken({ scope: `${SCOPES.join(" ")} Files.Read` }), {}, 400],
    ["unsupported grant", codeToken({ grant_type: "client_credentials" }), {}, 400],
    ["basic authentication", codeToken(), { headers: { Authorization: "Basic dGVzdDp0ZXN0" } }, 401],
    ["code field on refresh", { ...refreshToken(), code: CODE }, {}, 400],
    ["callback on refresh", { ...refreshToken(), redirect_uri: REDIRECT_URI }, {}, 400],
    ["missing refresh token", (() => { const value = refreshToken(); delete value.refresh_token; return value; })(), {}, 400]
  ];
  for (const [name, values, options, status] of tokenFailures) {
    await test(`token rejects ${name}`, async () => {
      const before = upstreamRequests.length;
      const response = await requestToken(base, values, options);
      const text = await response.text();
      assert(response.status === status, "Unsafe token request had wrong status.", { status: response.status, text });
      assert(upstreamRequests.length === before, "Unsafe token request reached provider.");
      assert(!text.includes(CLIENT_SECRET) && !text.includes(CODE) && !text.includes(REFRESH_TOKEN), "Token error leaked sensitive input.");
    });
  }

  await test("token rejects duplicate form parameters", async () => {
    const before = upstreamRequests.length;
    const body = `${form(codeToken())}&client_id=${encodeURIComponent(CLIENT_ID)}`;
    const response = await requestToken(base, body);
    assert(response.status === 400, "Duplicate token parameter was accepted.", await response.text());
    assert(upstreamRequests.length === before);
  });

  await test("token rejects non-form content", async () => {
    const response = await fetch(`${base}/token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}"
    });
    assert(response.status === 415, "JSON token request was accepted.", await response.text());
  });

  await test("token bounds request body size", async () => {
    const response = await fetch(`${base}/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: `padding=${"x".repeat(40_000)}`
    });
    assert(response.status === 413, "Oversized token request was accepted.", await response.text());
  });

  await test("provider invalid JSON becomes a generic bounded error", async () => {
    upstreamMode = "invalid-json";
    const response = await requestToken(base, codeToken());
    const text = await response.text();
    assert(response.status === 502 && JSON.parse(text).error === "server_error", "Invalid provider JSON leaked.", text);
    assert(!text.includes("not-json"));
    upstreamMode = "success";
  });

  await test("provider oversized JSON becomes a generic bounded error", async () => {
    upstreamMode = "oversized";
    const response = await requestToken(base, codeToken());
    const text = await response.text();
    assert(response.status === 502 && JSON.parse(text).error === "server_error", "Oversized provider JSON leaked.", response.status);
    assert(text.length < 256, "Oversized provider body was reflected.");
    upstreamMode = "success";
  });

  await test("provider network details are not leaked", async () => {
    upstreamMode = "network";
    const response = await requestToken(base, codeToken());
    const text = await response.text();
    assert(response.status === 502 && JSON.parse(text).error === "server_error");
    assert(!text.includes("sensitive detail"));
    upstreamMode = "success";
  });

  await test("untrusted browser origin is rejected without CORS reflection", async () => {
    const response = await fetch(`${base}/.well-known/oauth-authorization-server`, {
      headers: { Origin: "https://evil.example" }
    });
    assert(response.status === 403, "Untrusted origin was accepted.", await response.text());
    assert(!response.headers.has("access-control-allow-origin"), "Untrusted origin was reflected.");
  });

  await test("trusted CORS preflight is narrow", async () => {
    const response = await fetch(`${base}/token`, {
      method: "OPTIONS",
      headers: { Origin: "https://chatgpt.com", "Access-Control-Request-Method": "POST" }
    });
    assert(response.status === 204);
    assert(response.headers.get("access-control-allow-origin") === "https://chatgpt.com");
    assert(!response.headers.get("access-control-allow-headers").toLowerCase().includes("authorization"));
  });

  await test("wrong endpoint methods fail closed", async () => {
    const authorize = await fetch(authorizeUrl(base), { method: "POST", redirect: "manual" });
    const token = await fetch(`${base}/token`);
    const metadata = await fetch(`${base}/.well-known/oauth-authorization-server`, { method: "POST" });
    assert(authorize.status === 405 && token.status === 405 && metadata.status === 405);
  });

  await test("metadata rejects query strings", async () => {
    const response = await fetch(`${base}/.well-known/oauth-authorization-server?unsafe=1`);
    assert(response.status === 400, "Metadata query was accepted.", await response.text());
  });

  await test("unknown paths disclose no configuration", async () => {
    const response = await fetch(`${base}/missing`);
    const text = await response.text();
    assert(response.status === 404);
    assert(!text.includes(CLIENT_ID) && !text.includes(CLIENT_SECRET) && !text.includes(RESOURCE));
  });

  await test("structured diagnostics contain only redacted request shape", async () => {
    const diagnosticEntries = [];
    const diagnosticServer = createOAuthCompatServer(testEnv, {
      fetchImpl: mockFetch,
      rateLimits: { health: 500, metadata: 500, authorize: 500, token: 500, other: 500 },
      diagnostics: (entry) => diagnosticEntries.push(entry)
    });
    const diagnosticBase = await listen(diagnosticServer);
    const sensitiveState = "diagnostic-sensitive-state";
    const sensitiveQueryValue = "diagnostic-sensitive-query-value";
    const sensitiveCode = "diagnostic-sensitive-code";
    try {
      await fetch(authorizeUrl(diagnosticBase, { state: sensitiveState }), { redirect: "manual" });
      await fetch(
        authorizeUrl(diagnosticBase, { state: sensitiveState }, [["unexpected", sensitiveQueryValue]]),
        { redirect: "manual" }
      );
      await requestToken(diagnosticBase, codeToken({ code: sensitiveCode }));
      assert(diagnosticEntries.length === 3, "Unexpected diagnostic entry count.", diagnosticEntries);
      for (const entry of diagnosticEntries) {
        const expectedKeys = [
          "code", "parameterNames", "reason", "route", "status",
          ...(entry.route === "/token" && entry.status === 200 ? ["tokenResponse"] : [])
        ].sort();
        assert(
          JSON.stringify(Object.keys(entry).sort())
            === JSON.stringify(expectedKeys),
          "Diagnostics exposed an unexpected field.",
          entry
        );
      }
      const tokenDiagnostic = diagnosticEntries.find(
        (entry) => entry.route === "/token" && entry.status === 200
      )?.tokenResponse;
      assert(
        tokenDiagnostic?.tokenType === "Bearer"
          && tokenDiagnostic?.expiresIn === 3600
          && tokenDiagnostic?.hasRefreshToken === true
          && tokenDiagnostic?.accessTokenFormat === "jwt"
          && tokenDiagnostic?.claims?.aud === "6e97d01c-edf8-43fe-bf69-bb494ae22513",
        "Token response diagnostic is incomplete or unsafe.",
        tokenDiagnostic
      );
      const serialized = JSON.stringify(diagnosticEntries);
      for (const secretValue of [
        CLIENT_SECRET,
        RESOURCE,
        REDIRECT_URI,
        sensitiveState,
        sensitiveQueryValue,
        sensitiveCode
      ]) {
        assert(!serialized.includes(secretValue), "Diagnostics leaked an OAuth parameter value.", serialized);
      }
      assert(
        diagnosticEntries.some((entry) =>
          entry.route === "/authorize"
          && entry.status === 302
          && entry.code === "redirect"
          && entry.reason === "microsoft_authorization_redirect"
          && entry.parameterNames.includes("state")
        ),
        "Successful authorization diagnostic is missing.",
        diagnosticEntries
      );
      assert(
        diagnosticEntries.some((entry) =>
          entry.route === "/authorize"
          && entry.status === 302
          && entry.code === "invalid_request"
          && entry.reason === "unsupported_parameter"
          && entry.parameterNames.includes("<unsupported>")
        ),
        "Rejected authorization diagnostic is missing.",
        diagnosticEntries
      );
      assert(
        diagnosticEntries.some((entry) =>
          entry.route === "/token"
          && entry.status === 200
          && entry.code === "ok"
          && entry.reason === "microsoft_token_response"
          && entry.parameterNames.includes("client_secret")
          && entry.parameterNames.includes("code")
        ),
        "Token diagnostic is missing.",
        diagnosticEntries
      );
    } finally {
      await close(diagnosticServer);
    }
  });

  await test("diagnostics suppress healthy probes and sample repeated throttles", async () => {
    let currentTime = 2_300_000_000_000;
    const sampledEntries = [];
    const sampledServer = createOAuthCompatServer({
      ...publicTestEnv,
      ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE: join(
        TEST_DIRECTORY,
        "sampled-diagnostics-store.json"
      )
    }, {
      fetchImpl: mockFetch,
      now: () => currentTime,
      rateLimits: { authorize: 1, callback: 500, token: 1 },
      diagnostics: (entry) => sampledEntries.push(entry)
    });
    const sampledBase = await listen(sampledServer);
    try {
      await fetch(`${sampledBase}/healthz`);
      await fetch(`${sampledBase}/healthz`);
      await fetch(authorizeUrl(sampledBase), { redirect: "manual" });
      await fetch(authorizeUrl(sampledBase), { redirect: "manual" });
      await fetch(authorizeUrl(sampledBase), { redirect: "manual" });
      await requestToken(sampledBase, publicCodeToken());
      await requestToken(sampledBase, publicCodeToken());
      await requestToken(sampledBase, publicCodeToken());
      assert(
        !sampledEntries.some((entry) => entry.route === "/healthz"),
        "Successful health probes were logged.",
        sampledEntries
      );
      assert(
        sampledEntries.filter((entry) => entry.status === 429).length === 1,
        "Repeated 429 diagnostics were not sampled.",
        sampledEntries
      );
      assert(
        sampledEntries.filter((entry) =>
          entry.route === "/authorize"
          && entry.code === "temporarily_unavailable"
        ).length === 1,
        "OAuth redirect-form throttle diagnostics were not sampled.",
        sampledEntries
      );

      currentTime += 60_001;
      await requestToken(sampledBase, publicCodeToken());
      await requestToken(sampledBase, publicCodeToken());
      assert(
        sampledEntries.filter((entry) => entry.status === 429).length === 2,
        "429 diagnostic sampling did not recover after its window.",
        sampledEntries
      );
      assert(
        !JSON.stringify(sampledEntries).includes(DUMMY_FACADE_CODE),
        "Sampled diagnostics exposed an opaque authorization code."
      );
    } finally {
      await close(sampledServer);
    }
  });

  await test("token endpoint rate limit is bounded", async () => {
    const rateServer = createOAuthCompatServer(testEnv, {
      fetchImpl: mockFetch,
      rateLimits: { token: 1 },
      diagnostics: () => {}
    });
    const rateBase = await listen(rateServer);
    try {
      const invalid = await requestToken(rateBase, codeToken({ client_secret: "invalid-client-secret" }));
      const first = await requestToken(rateBase, codeToken());
      const second = await requestToken(rateBase, codeToken());
      assert(invalid.status === 401, "Invalid client request had the wrong status.", await invalid.text());
      assert(first.status === 200, "First bounded token request failed.", await first.text());
      assert(second.status === 429, "Token rate limit did not fire.", await second.text());
      assert(Number(second.headers.get("retry-after")) >= 1, "Rate response lacks Retry-After.");
    } finally {
      await close(rateServer);
    }
  });
} finally {
  await Promise.all([
    close(server),
    close(combinedConsentServer),
    close(publicServer),
    close(publicCompatServer),
    close(cimdServer)
  ]);
  rmSync(TEST_DIRECTORY, { recursive: true, force: true });
}

process.stdout.write(`OAuth compatibility tests: ${passed} passed, ${failed} failed\n`);
if (failed) process.exitCode = 1;
