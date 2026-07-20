#!/usr/bin/env node

import { spawn } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const scriptsRoot = dirname(fileURLToPath(import.meta.url));
const pluginRoot = join(scriptsRoot, "..");
const httpHost = process.env.ONEDRIVE_MCP_HTTP_HOST || "127.0.0.1";
const httpPort = process.env.ONEDRIVE_MCP_HTTP_PORT || "3001";
const toolProfile = process.env.ONEDRIVE_TOOL_PROFILE || "chatgpt";
const healthUrl = `http://${httpHost}:${httpPort}/healthz`;
let stopping = false;

if (!["full", "chatgpt"].includes(toolProfile)) {
  throw new Error("ONEDRIVE_TOOL_PROFILE must be full or chatgpt.");
}

function child(script, env = process.env) {
  return spawn(process.execPath, [script], {
    cwd: pluginRoot,
    env,
    stdio: "inherit"
  });
}

async function waitForHttpServer(processHandle) {
  const deadline = Date.now() + 20_000;
  while (Date.now() < deadline) {
    if (processHandle.exitCode !== null) {
      throw new Error(`The OneDrive MCP HTTP server exited with code ${processHandle.exitCode} before becoming healthy.`);
    }
    try {
      const response = await fetch(healthUrl, { signal: AbortSignal.timeout(1_000) });
      if (response.ok) return;
    } catch {
      // Startup races are expected until the listener is ready.
    }
    await new Promise((resolvePromise) => setTimeout(resolvePromise, 250));
  }
  throw new Error(`The OneDrive MCP HTTP server did not become healthy at ${healthUrl} within 20 seconds.`);
}

// The HTTP server owns MCP initialize/tools/list handling, so it must receive
// the focused ChatGPT profile directly. Passing the profile only to the tunnel
// client leaves the server on its much larger default `full` contract.
const httpServer = child(join(pluginRoot, "mcp", "http-server.mjs"), {
  ...process.env,
  ONEDRIVE_TOOL_PROFILE: toolProfile
});
let tunnel = null;

async function stop(signal = "SIGTERM") {
  if (stopping) return;
  stopping = true;
  if (tunnel && tunnel.exitCode === null) tunnel.kill(signal);
  if (httpServer.exitCode === null) httpServer.kill(signal);
}

for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.once(signal, async () => {
    await stop(signal);
  });
}

try {
  await waitForHttpServer(httpServer);
  tunnel = child(join(scriptsRoot, "run-chatgpt-tunnel.mjs"));
  tunnel.once("exit", async (code, signal) => {
    await stop(signal || "SIGTERM");
    process.exit(code ?? (signal ? 0 : 1));
  });
  httpServer.once("exit", async (code, signal) => {
    if (stopping) return;
    if (tunnel && tunnel.exitCode === null) tunnel.kill("SIGTERM");
    process.stderr.write(`OneDrive MCP HTTP server exited unexpectedly (${signal || code}).\n`);
    process.exit(code || 1);
  });
} catch (error) {
  await stop();
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exit(1);
}
