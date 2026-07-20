#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const scriptPath = fileURLToPath(import.meta.url);

if (process.argv.includes("--probe")) {
  const { processMcpMessage, shutdownOneDriveServer } = await import("../mcp/server.mjs");
  const initialized = await processMcpMessage({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2024-11-05" }
  });
  const listed = await processMcpMessage({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
  const tools = listed.result.tools;
  const officeTransform = tools.find((tool) => tool.name === "onedrive_office_batch_transform");
  console.log(JSON.stringify({
    profile: process.env.ONEDRIVE_TOOL_PROFILE || "full",
    count: tools.length,
    bytes: Buffer.byteLength(JSON.stringify(listed)),
    names: tools.map((tool) => tool.name),
    officeTransformBytes: Buffer.byteLength(JSON.stringify(officeTransform || {})),
    instructions: initialized.result.instructions || "",
    serverVersion: initialized.result.serverInfo?.version || ""
  }));
  await shutdownOneDriveServer();
  process.exit(0);
}

function probe(profile) {
  const result = spawnSync(process.execPath, [scriptPath, "--probe"], {
    env: { ...process.env, ONEDRIVE_TOOL_PROFILE: profile },
    encoding: "utf8",
    timeout: 10_000
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(result.stderr || `Profile ${profile} probe failed.`);
  return JSON.parse(result.stdout.trim());
}

function assert(condition, message, details = undefined) {
  if (!condition) {
    const error = new Error(message);
    error.details = details;
    throw error;
  }
}

try {
  const full = probe("full");
  const chatgpt = probe("chatgpt");
  const fullNames = new Set(full.names);
  assert(full.count === 84, "Full profile must preserve the 84-tool contract.", full);
  assert(chatgpt.count === 26, "ChatGPT profile must expose the reviewed 26-tool surface.", chatgpt);
  assert(chatgpt.names.every((name) => fullNames.has(name)), "ChatGPT profile must be a subset of the full contract.", chatgpt.names);
  assert(chatgpt.bytes <= 40 * 1024, "ChatGPT tools/list payload must stay at or below 40 KiB.", chatgpt);
  assert(chatgpt.bytes < full.bytes * 0.15, "ChatGPT tools/list payload must remain at least 85% smaller than full.", { full, chatgpt });
  assert(chatgpt.officeTransformBytes <= 4096, "ChatGPT Office transform descriptor must remain compact.", chatgpt);
  assert(chatgpt.instructions.length > 0 && chatgpt.instructions.length <= 512, "Server instructions must be present and concise.", chatgpt.instructions);
  assert(chatgpt.serverVersion !== full.serverVersion && chatgpt.serverVersion.includes(".chatgpt."), "ChatGPT metadata must use a contract-specific server version to invalidate stale app caches.", { full: full.serverVersion, chatgpt: chatgpt.serverVersion });

  const invalid = spawnSync(process.execPath, [scriptPath, "--probe"], {
    env: { ...process.env, ONEDRIVE_TOOL_PROFILE: "invalid" },
    encoding: "utf8",
    timeout: 10_000
  });
  assert(invalid.status !== 0, "Invalid tool profiles must fail closed.");

  console.log(JSON.stringify({
    ok: true,
    full: { count: full.count, bytes: full.bytes, officeTransformBytes: full.officeTransformBytes, serverVersion: full.serverVersion },
    chatgpt: { count: chatgpt.count, bytes: chatgpt.bytes, officeTransformBytes: chatgpt.officeTransformBytes, serverVersion: chatgpt.serverVersion },
    reductionPercent: Number(((1 - chatgpt.bytes / full.bytes) * 100).toFixed(1))
  }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message, details: error.details }, null, 2));
  process.exit(1);
}
