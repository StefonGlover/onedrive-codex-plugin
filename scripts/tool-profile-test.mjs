#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const scriptPath = fileURLToPath(import.meta.url);

if (process.argv.includes("--probe")) {
  const { boundChatgptToolPayload, processMcpMessage, shutdownOneDriveServer } = await import("../mcp/server.mjs");
  const initialized = await processMcpMessage({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2024-11-05" }
  });
  const listed = await processMcpMessage({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
  const tools = listed.result.tools;
  const officeTransform = tools.find((tool) => tool.name === "onedrive_office_batch_transform");
  const compatibility = Object.fromEntries(
    ["search", "fetch", "onedrive_upload_file", "onedrive_permanent_delete"].map((name) => [name, tools.find((tool) => tool.name === name) || null])
  );
  const oversized = boundChatgptToolPayload({ rows: [{ value: "x".repeat(11 * 1024 * 1024) }] });
  console.log(JSON.stringify({
    profile: process.env.ONEDRIVE_TOOL_PROFILE || "full",
    count: tools.length,
    bytes: Buffer.byteLength(JSON.stringify(listed)),
    names: tools.map((tool) => tool.name),
    metadata: tools.map((tool) => ({
      name: tool.name,
      title: tool.title,
      description: tool.description,
      invoking: tool._meta?.["openai/toolInvocation/invoking"] || null,
      invoked: tool._meta?.["openai/toolInvocation/invoked"] || null
    })),
    compatibility,
    oversized: {
      truncated: oversized.truncated,
      originalBytes: oversized.originalBytes,
      boundedBytes: oversized.boundedBytes
    },
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
  const compatibilityNames = new Set(["search", "fetch", "onedrive_upload_file", "onedrive_permanent_delete"]);
  assert(full.count === 84, "Full profile must preserve the 84-tool contract.", full);
  assert(chatgpt.count === 19, "ChatGPT profile must expose the reviewed 19-tool surface.", chatgpt);
  assert(chatgpt.names.every((name) => fullNames.has(name) || compatibilityNames.has(name)), "ChatGPT profile may add only the reviewed compatibility tools.", chatgpt.names);
  assert(chatgpt.names.includes("search") && chatgpt.names.includes("fetch"), "ChatGPT profile must expose standard search and fetch tools.", chatgpt.names);
  assert(chatgpt.names.includes("onedrive_upload_file") && chatgpt.names.includes("onedrive_restore_deleted") && chatgpt.names.includes("onedrive_permanent_delete"), "ChatGPT profile must expose upload, recycle-bin restore, and guarded permanent delete.", chatgpt.names);
  assert(!chatgpt.names.includes("onedrive_find") && !chatgpt.names.includes("onedrive_get_info") && !chatgpt.names.includes("onedrive_read_text"), "ChatGPT profile must not advertise redundant slow retrieval tools.", chatgpt.names);
  assert(!chatgpt.names.includes("onedrive_preview") && !chatgpt.names.includes("onedrive_recent") && !chatgpt.names.includes("onedrive_office_search"), "ChatGPT profile must not advertise redundant retrieval helpers.", chatgpt.names);
  assert(!chatgpt.names.includes("onedrive_word_get_document") && !chatgpt.names.includes("onedrive_excel_get_workbook") && !chatgpt.names.includes("onedrive_powerpoint_get_presentation"), "ChatGPT profile must use the bounded fetch extractor instead of redundant high-volume Office reads.", chatgpt.names);
  assert(!full.names.includes("search") && !full.names.includes("fetch") && !full.names.includes("onedrive_upload_file") && !full.names.includes("onedrive_permanent_delete"), "ChatGPT compatibility tools must not change the immutable full tool contract.", full.names);
  assert(JSON.stringify(chatgpt.compatibility.search?.inputSchema) === JSON.stringify({ type: "object", required: ["query"], properties: { query: { type: "string", minLength: 1 } }, additionalProperties: false }), "search must keep the exact company-knowledge input contract.", chatgpt.compatibility.search);
  assert(JSON.stringify(chatgpt.compatibility.fetch?.inputSchema) === JSON.stringify({ type: "object", required: ["id"], properties: { id: { type: "string", minLength: 1 } }, additionalProperties: false }), "fetch must keep the exact company-knowledge input contract.", chatgpt.compatibility.fetch);
  assert(chatgpt.compatibility.search?.outputSchema?.required?.includes("results"), "search must advertise the standard results output.", chatgpt.compatibility.search);
  assert(chatgpt.compatibility.fetch?.outputSchema?.required?.includes("text"), "fetch must advertise readable text output.", chatgpt.compatibility.fetch);
  assert(chatgpt.compatibility.search?.annotations?.readOnlyHint === true && chatgpt.compatibility.fetch?.annotations?.readOnlyHint === true, "search/fetch must remain read-only.", chatgpt.compatibility);
  assert(JSON.stringify(chatgpt.compatibility.onedrive_upload_file?._meta?.["openai/fileParams"]) === JSON.stringify(["sourceFile"]), "ChatGPT upload must advertise its file parameter.", chatgpt.compatibility.onedrive_upload_file);
  assert(chatgpt.compatibility.onedrive_upload_file?.annotations?.destructiveHint === true && chatgpt.compatibility.onedrive_permanent_delete?.annotations?.destructiveHint === true, "Upload replacement and permanent delete must advertise destructive impact.", chatgpt.compatibility);
  assert(chatgpt.oversized.truncated === true && chatgpt.oversized.boundedBytes <= 1024 * 1024, "Oversized ChatGPT tool results must be bounded below the response cap.", chatgpt.oversized);
  assert(chatgpt.bytes <= 40 * 1024, "ChatGPT tools/list payload must stay at or below 40 KiB.", chatgpt);
  assert(chatgpt.bytes < full.bytes * 0.15, "ChatGPT tools/list payload must remain at least 85% smaller than full.", { full, chatgpt });
  assert(chatgpt.officeTransformBytes <= 4096, "ChatGPT Office transform descriptor must remain compact.", chatgpt);
  assert(chatgpt.instructions.length > 0 && chatgpt.instructions.length <= 512, "Server instructions must be present and concise.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("fetch each file") && !chatgpt.instructions.includes("matching structured read tool"), "ChatGPT server instructions must use the focused fetch-first Office sequence.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("call search separately"), "ChatGPT server instructions must prevent multi-target queries from hiding distinct result IDs.", chatgpt.instructions);
  assert(chatgpt.metadata.every((tool) => /^Use this when\b/u.test(tool.description || "")), "Every focused ChatGPT tool description must begin with a discriminative 'Use this when' cue.", chatgpt.metadata);
  assert(new Set(chatgpt.metadata.map((tool) => tool.description)).size === chatgpt.metadata.length, "Focused ChatGPT tool descriptions must be unique.", chatgpt.metadata);
  assert(chatgpt.metadata.every((tool) => tool.invoking && tool.invoked && tool.invoking.length <= 64 && tool.invoked.length <= 64), "Every focused ChatGPT tool must advertise bounded invocation status text.", chatgpt.metadata);
  assert(chatgpt.metadata.find((tool) => tool.name === "fetch")?.description.includes("continuation ID"), "Fetch metadata must explain progressive continuation behavior.", chatgpt.metadata);
  for (const name of ["onedrive_rename", "onedrive_move", "onedrive_copy", "onedrive_create_sharing_link", "onedrive_revoke_permission"]) {
    const description = chatgpt.metadata.find((tool) => tool.name === name)?.description || "";
    assert(description.includes("Inputs:") && description.includes("expectedId or expectedName"), `${name} must expose compact live-call input guidance for deferred ChatGPT schema loading.`, description);
  }
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
    chatgpt: { count: chatgpt.count, bytes: chatgpt.bytes, officeTransformBytes: chatgpt.officeTransformBytes, serverVersion: chatgpt.serverVersion, oversized: chatgpt.oversized },
    reductionPercent: Number(((1 - chatgpt.bytes / full.bytes) * 100).toFixed(1))
  }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message, details: error.details }, null, 2));
  process.exit(1);
}
