#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptsRoot = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(scriptsRoot, "..");
const testRoot = await mkdtemp(join(tmpdir(), "onedrive-hosted-profile-boundary-"));
const missingCredentialPath = join(testRoot, "missing-tunnel.env");
const missingTunnelClient = join(testRoot, "missing-tunnel-client");
const missingPythonPath = join(testRoot, "private-runtime", "missing-python");

const forbiddenFocusedFields = new Set([
  "localPath",
  "manifestPath",
  "stagingPath",
  "transactionRoot",
  "backupPath",
  "inputPath",
  "outputPath",
  "pythonPath",
  "helperPath",
  "storagePath",
  "cachePath",
  "auditPath",
  "downloadPath",
  "updatePath"
]);
const concreteHostPathPattern = /(?:\/(?:Users|home|tmp|var|private|etc|opt|usr|srv|run|app|mnt)(?:\/|$)|~\/\.codex(?:\/|$)|file:\/\/|[A-Za-z]:\\)/u;

function launcherEnvironment() {
  return {
    ...process.env,
    HOME: testRoot,
    ONEDRIVE_TOOL_PROFILE: "full",
    ONEDRIVE_TUNNEL_ENV_FILE: missingCredentialPath,
    ONEDRIVE_TUNNEL_CLIENT: missingTunnelClient,
    ONEDRIVE_MCP_HTTP_PORT: "0",
    ONEDRIVE_OAUTH_COMPAT_ENABLED: "false"
  };
}

function assertLauncherRejectsFull(scriptName, expectedMessage) {
  const result = spawnSync(process.execPath, [join(scriptsRoot, scriptName)], {
    cwd: pluginRoot,
    env: launcherEnvironment(),
    encoding: "utf8",
    timeout: 5_000
  });
  assert.ifError(result.error);
  assert.equal(result.status, 1, `${scriptName} must exit unsuccessfully for the full profile.`);
  const output = `${result.stdout || ""}\n${result.stderr || ""}`;
  assert.match(output, expectedMessage, `${scriptName} must explain the hosted profile boundary.`);
  assert.ok(!output.includes(missingCredentialPath), `${scriptName} must reject the profile before reading tunnel credentials.`);
  assert.doesNotMatch(output, /Could not read tunnel credential|before becoming healthy|listening on http/u, `${scriptName} must reject before credential, child-server, or network setup.`);
}

function pathFindings(value, path = "$", findings = []) {
  if (Array.isArray(value)) {
    value.forEach((entry, index) => pathFindings(entry, `${path}[${index}]`, findings));
    return findings;
  }
  if (!value || typeof value !== "object") {
    if (typeof value === "string" && concreteHostPathPattern.test(value)) {
      findings.push({ type: "literal", path, value });
    }
    return findings;
  }
  for (const [key, entry] of Object.entries(value)) {
    const entryPath = `${path}.${key}`;
    if (forbiddenFocusedFields.has(key)) findings.push({ type: "field", path: entryPath });
    pathFindings(entry, entryPath, findings);
  }
  return findings;
}

function parseJsonLines(output) {
  return String(output || "")
    .trim()
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

let shutdownOneDriveServer = null;
try {
  assertLauncherRejectsFull(
    "run-chatgpt-tunnel.mjs",
    /ChatGPT tunnel requires ONEDRIVE_TOOL_PROFILE=chatgpt/u
  );
  assertLauncherRejectsFull(
    "run-chatgpt-oauth-tunnel.mjs",
    /OAuth HTTP tunnel requires ONEDRIVE_TOOL_PROFILE=chatgpt/u
  );

  process.env.HOME = testRoot;
  process.env.ONEDRIVE_TOOL_PROFILE = "chatgpt";
  process.env.ONEDRIVE_MCP_AUTH_MODE = "noauth";
  process.env.ONEDRIVE_STORAGE_ROOT = join(testRoot, "storage");
  process.env.ONEDRIVE_CACHE_ROOT = join(testRoot, "cache");
  process.env.ONEDRIVE_TOKEN_STORE = "encrypted-file";
  process.env.ONEDRIVE_TOKEN_FILE = join(testRoot, "missing-token.enc");
  process.env.ONEDRIVE_OFFICE_PYTHON = missingPythonPath;
  delete process.env.ONEDRIVE_TEST_ACCESS_TOKEN;

  const httpServerModule = await import("../mcp/http-server.mjs");
  const oneDriveServerModule = await import("../mcp/server.mjs");
  shutdownOneDriveServer = oneDriveServerModule.shutdownOneDriveServer;

  assert.throws(
    () => httpServerModule.createOneDriveHttpServer({ ...process.env, ONEDRIVE_TOOL_PROFILE: "full" }),
    /Streamable HTTP requires ONEDRIVE_TOOL_PROFILE=chatgpt/u,
    "Streamable HTTP must reject the local-only full profile."
  );
  const unsetProfileEnvironment = { ...process.env };
  delete unsetProfileEnvironment.ONEDRIVE_TOOL_PROFILE;
  assert.throws(
    () => httpServerModule.createOneDriveHttpServer(unsetProfileEnvironment),
    /Streamable HTTP requires ONEDRIVE_TOOL_PROFILE=chatgpt/u,
    "Streamable HTTP must reject an unset profile rather than inheriting the server module's default."
  );

  for (const host of [undefined, "", "   ", "127.0.0.1", " 127.0.0.1 ", "::1", "localhost", "LOCALHOST"]) {
    const environment = { ...process.env, ONEDRIVE_MCP_AUTH_MODE: "noauth" };
    if (host === undefined) delete environment.ONEDRIVE_MCP_HTTP_HOST;
    else environment.ONEDRIVE_MCP_HTTP_HOST = host;
    const server = httpServerModule.createOneDriveHttpServer(environment);
    assert.equal(typeof server.listen, "function", `Hosted HTTP must accept the canonical loopback host ${JSON.stringify(host)}.`);
  }
  const refusedHostedHttpHosts = [
    "0.0.0.0", "::", "127.0.0.2", "10.0.0.1", "192.168.1.10",
    "example.com", "host.docker.internal", "localhost.example", "[::1]"
  ];
  for (const authMode of [undefined, "noauth", "oauth"]) {
    for (const host of refusedHostedHttpHosts) {
      const environment = { ...process.env, ONEDRIVE_MCP_HTTP_HOST: host };
      if (authMode === undefined) delete environment.ONEDRIVE_MCP_AUTH_MODE;
      else environment.ONEDRIVE_MCP_AUTH_MODE = authMode;
      assert.throws(
        () => httpServerModule.createOneDriveHttpServer(environment),
        /canonical loopback host/u,
        `Hosted HTTP must reject ${host} before ${authMode || "default"} authentication configuration is evaluated.`
      );
    }
  }

  const listed = await oneDriveServerModule.processMcpMessage({
    jsonrpc: "2.0",
    id: 1,
    method: "tools/list",
    params: {}
  });
  const focusedTools = listed?.result?.tools;
  assert.ok(Array.isArray(focusedTools), "Focused tools/list must return a tools array.");
  assert.equal(focusedTools.length, 19, "Hosted tools/list must expose exactly the reviewed focused surface.");
  assert.deepEqual(pathFindings(focusedTools), [], "Focused descriptors must not expose local filesystem selectors or concrete host paths.");

  const capabilities = await oneDriveServerModule.processMcpMessage({
    jsonrpc: "2.0",
    id: 2,
    method: "tools/call",
    params: { name: "onedrive_office_capabilities", arguments: {} }
  });
  assert.equal(capabilities?.result?.isError, false, "Focused Office capability discovery must remain available when the local runtime is misconfigured.");
  const capabilityText = capabilities.result.content?.find((entry) => entry?.type === "text")?.text;
  assert.equal(typeof capabilityText, "string", "Focused Office capability discovery must return text content.");
  const capabilityValue = JSON.parse(capabilityText);
  assert.equal(capabilityValue.runtime?.error, "The Office document runtime is unavailable or misconfigured.");
  assert.ok(!JSON.stringify(capabilities).includes(missingPythonPath), "Focused Office runtime failures must not disclose the configured interpreter path.");
  assert.deepEqual(pathFindings(capabilities), [], "Focused Office capability results must not expose local filesystem fields or host paths.");

  const fullInput = [
    JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list", params: {} }),
    JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/call", params: { name: "onedrive_config", arguments: {} } }),
    ""
  ].join("\n");
  const fullStdio = spawnSync(process.execPath, [join(pluginRoot, "mcp", "server.mjs")], {
    cwd: pluginRoot,
    env: {
      ...process.env,
      ONEDRIVE_TOOL_PROFILE: "full",
      ONEDRIVE_STORAGE_ROOT: join(testRoot, "full-storage"),
      ONEDRIVE_CACHE_ROOT: join(testRoot, "full-cache"),
      ONEDRIVE_TOKEN_FILE: join(testRoot, "full-missing-token.enc")
    },
    input: fullInput,
    encoding: "utf8",
    timeout: 10_000
  });
  assert.ifError(fullStdio.error);
  assert.equal(fullStdio.status, 0, fullStdio.stderr || "Trusted local full-profile stdio probe failed.");
  const fullMessages = parseJsonLines(fullStdio.stdout);
  const fullTools = fullMessages.find((message) => message.id === 1)?.result?.tools;
  const fullConfig = fullMessages.find((message) => message.id === 2)?.result;
  assert.equal(fullTools?.length, 84, "Trusted local stdio must preserve the full 84-tool contract.");
  assert.equal(fullConfig?.isError, false, "Trusted local stdio must continue executing a full-profile tool.");
  assert.ok(
    fullTools.find((tool) => tool.name === "onedrive_upload")?.inputSchema?.properties?.localPath,
    "The full local-only profile must retain its intentional local filesystem upload selector."
  );

  process.stdout.write(`${JSON.stringify({
    ok: true,
    checks: {
      fullChatgptLaunchersRejected: 2,
      streamableHttpFullAndUnsetRejected: true,
      hostedHttpLoopbackHostsAccepted: 8,
      hostedHttpNonLoopbackHostsRejected: refusedHostedHttpHosts.length * 3,
      focusedToolDescriptors: focusedTools.length,
      focusedRuntimeFailurePathsRedacted: true,
      trustedLocalStdioTools: fullTools.length
    }
  }, null, 2)}\n`);
} finally {
  if (shutdownOneDriveServer) await shutdownOneDriveServer().catch(() => null);
  await rm(testRoot, { recursive: true, force: true });
}
