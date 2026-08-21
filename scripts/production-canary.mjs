#!/usr/bin/env node

const allowedFlags = new Set(["base-url", "samples", "p95-ms", "expect-tools", "expect-version-prefix", "self-check"]);

function parseArgs(argv = []) {
  const result = {};
  for (const raw of argv) {
    if (!raw.startsWith("--")) throw new Error(`Unexpected positional argument: ${raw}`);
    const [key, ...rest] = raw.slice(2).split("=");
    if (!allowedFlags.has(key)) throw new Error(`Unknown flag: --${key}`);
    if (Object.hasOwn(result, key)) throw new Error(`Duplicate flag: --${key}`);
    result[key] = rest.length ? rest.join("=") : true;
  }
  return result;
}

function boundedInteger(value, name, fallback, minimum, maximum) {
  if (value === undefined) return fallback;
  const number = Number(value);
  if (!Number.isInteger(number) || number < minimum || number > maximum) {
    throw new Error(`--${name} must be an integer from ${minimum} to ${maximum}.`);
  }
  return number;
}

function percentile(values, probability) {
  if (!values.length) return null;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * probability) - 1))];
}

function summarizeLatency(values) {
  return {
    samples: values.length,
    minMs: Math.min(...values),
    p50Ms: percentile(values, 0.5),
    p95Ms: percentile(values, 0.95),
    maxMs: Math.max(...values)
  };
}

function validatedBaseUrl(value) {
  const url = new URL(String(value || ""));
  if (url.username || url.password || url.search || url.hash) throw new Error("--base-url must not contain credentials, query, or fragment.");
  if (url.protocol !== "https:" && !(url.protocol === "http:" && new Set(["127.0.0.1", "localhost", "::1"]).has(url.hostname))) {
    throw new Error("--base-url must use HTTPS except for an explicit loopback test endpoint.");
  }
  url.pathname = url.pathname.replace(/\/+$/, "");
  return url;
}

function safeJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

async function request(baseUrl, path, init = {}) {
  const startedAt = performance.now();
  const response = await fetch(new URL(path, `${baseUrl.origin}${baseUrl.pathname || "/"}`), {
    redirect: "error",
    signal: AbortSignal.timeout(20_000),
    ...init
  });
  const text = await response.text();
  return {
    status: response.status,
    durationMs: Math.round((performance.now() - startedAt) * 10) / 10,
    authenticate: response.headers.get("www-authenticate"),
    allow: response.headers.get("allow"),
    body: safeJson(text)
  };
}

function assert(condition, message, details = undefined) {
  if (condition) return;
  const error = new Error(message);
  error.details = details;
  throw error;
}

function initializeRequest() {
  return {
    method: "POST",
    headers: { accept: "application/json, text/event-stream", "content-type": "application/json" },
    body: JSON.stringify({
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: { protocolVersion: "2025-11-25", capabilities: {}, clientInfo: { name: "onedrive-production-canary", version: "1" } }
    })
  };
}

function toolsListRequest() {
  return {
    method: "POST",
    headers: { accept: "application/json, text/event-stream", "content-type": "application/json" },
    body: JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} })
  };
}

async function runSelfCheck() {
  const checks = {
    percentileP50: percentile([9, 1, 5, 3, 7], 0.5) === 5,
    percentileP95: percentile([9, 1, 5, 3, 7], 0.95) === 9,
    localhostAllowed: validatedBaseUrl("http://127.0.0.1:3001").origin === "http://127.0.0.1:3001",
    credentialsRejected: false,
    invalidSamplesRejected: false
  };
  try { validatedBaseUrl("https://user:secret@example.test"); } catch { checks.credentialsRejected = true; }
  try { boundedInteger("0", "samples", 5, 1, 50); } catch { checks.invalidSamplesRejected = true; }
  const ok = Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, checks }, null, 2));
  return ok;
}

const args = parseArgs(process.argv.slice(2));
if (args["self-check"] !== undefined) process.exit(await runSelfCheck() ? 0 : 1);

const baseUrl = validatedBaseUrl(args["base-url"]);
const samples = boundedInteger(args.samples, "samples", 5, 1, 50);
const p95BudgetMs = boundedInteger(args["p95-ms"], "p95-ms", 2_000, 100, 60_000);
const expectedTools = boundedInteger(args["expect-tools"], "expect-tools", 21, 1, 500);
const expectedVersionPrefix = String(args["expect-version-prefix"] || "").trim();

try {
  const healthSamples = [];
  let latestHealth;
  for (let index = 0; index < samples; index += 1) {
    latestHealth = await request(baseUrl, "/healthz");
    assert(latestHealth.status === 200 && latestHealth.body?.ok === true, "Health endpoint failed.", { status: latestHealth.status });
    healthSamples.push(latestHealth.durationMs);
  }
  const authorization = await request(baseUrl, "/.well-known/oauth-authorization-server");
  const protectedResource = await request(baseUrl, "/.well-known/oauth-protected-resource/mcp");
  const initialized = await request(baseUrl, "/mcp", initializeRequest());
  const protectedTools = await request(baseUrl, "/mcp", toolsListRequest());
  const mcpGet = await request(baseUrl, "/mcp");

  assert(authorization.status === 200 && typeof authorization.body?.issuer === "string", "OAuth authorization-server metadata failed.", { status: authorization.status });
  assert(protectedResource.status === 200 && typeof protectedResource.body?.resource === "string", "OAuth protected-resource metadata failed.", { status: protectedResource.status });
  assert(initialized.status === 200 && initialized.body?.result?.serverInfo?.name === "onedrive", "MCP initialize failed.", { status: initialized.status, body: initialized.body });
  assert(protectedTools.status === 401 && /Bearer/i.test(protectedTools.authenticate || ""), "Unauthenticated tools/list was not rejected with a Bearer challenge.", { status: protectedTools.status });
  assert(mcpGet.status === 405, "MCP GET must be refused.", { status: mcpGet.status, allow: mcpGet.allow });

  const latency = summarizeLatency(healthSamples);
  assert(latency.p95Ms <= p95BudgetMs, "Health latency exceeded the p95 budget.", { latency, p95BudgetMs });
  const observability = latestHealth.body?.mcp?.observability || latestHealth.body?.observability;
  assert(observability?.release?.toolCount === expectedTools, "Deployed focused tool count does not match the release expectation.", observability?.release);
  assert(observability?.release?.profile === "chatgpt", "Hosted release is not using the ChatGPT profile.", observability?.release);
  if (expectedVersionPrefix) {
    assert(String(observability?.release?.version || "").startsWith(expectedVersionPrefix), "Deployed server version does not match the expected release prefix.", observability?.release);
  }

  console.log(JSON.stringify({
    ok: true,
    endpoint: baseUrl.origin,
    release: observability.release,
    healthLatency: latency,
    p95BudgetMs,
    checks: {
      health: true,
      authorizationServerMetadata: true,
      protectedResourceMetadata: true,
      mcpInitialize: true,
      unauthenticatedToolsRejected: true,
      mcpGetRejected: true
    },
    counters: {
      uptimeSeconds: observability.uptimeSeconds,
      toolCalls: observability.toolCalls,
      toolErrors: observability.toolErrors,
      graphThrottles: observability.graphThrottles,
      oauthFailures: observability.oauthFailures,
      rollingLatency: observability.latency
    }
  }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message, details: error.details }, null, 2));
  process.exitCode = 1;
}
