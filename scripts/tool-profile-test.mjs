#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const scriptPath = fileURLToPath(import.meta.url);

if (process.argv.includes("--probe")) {
  const {
    boundChatgptToolPayload,
    isPublicChatgptFileAddress,
    processMcpMessage,
    shutdownOneDriveServer,
    trustedChatgptFileUrl
  } = await import("../mcp/server.mjs");
  const rejectsChatgptFileUrl = (value) => {
    try {
      trustedChatgptFileUrl(value);
      return false;
    } catch {
      return true;
    }
  };
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
    [
      "search",
      "fetch",
      "onedrive_open_files",
      "onedrive_preview_actions",
      "onedrive_upload_file",
      "onedrive_export_file",
      "onedrive_write_text",
      "onedrive_create_sharing_link",
      "onedrive_invite_permission",
      "onedrive_create_folder",
      "onedrive_copy",
      "onedrive_delete",
      "onedrive_revoke_permission",
      "onedrive_office_batch_transform"
    ].map((name) => [name, tools.find((tool) => tool.name === name) || null])
  );
  const oversized = boundChatgptToolPayload({ rows: [{ value: "x".repeat(11 * 1024 * 1024) }] });
  const probeOutput = JSON.stringify({
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
    security: {
      topLevel: tools.map((tool) => tool.securitySchemes),
      mirrored: tools.map((tool) => tool._meta?.securitySchemes)
    },
    annotations: tools.map((tool) => tool.annotations),
    compatibility,
    oversized: {
      truncated: oversized.truncated,
      originalBytes: oversized.originalBytes,
      boundedBytes: oversized.boundedBytes
    },
    officeTransformBytes: Buffer.byteLength(JSON.stringify(officeTransform || {})),
    instructions: initialized.result.instructions || "",
    serverVersion: initialized.result.serverInfo?.version || "",
    chatgptFileUrlPolicy: {
      acceptsHostAgnosticHttps: trustedChatgptFileUrl("https://future-files.example.test/signed-download?token=redacted").hostname === "future-files.example.test",
      rejectsHttp: rejectsChatgptFileUrl("http://future-files.example.test/file"),
      rejectsCredentials: rejectsChatgptFileUrl("https://user:password@future-files.example.test/file"),
      rejectsNonStandardPort: rejectsChatgptFileUrl("https://future-files.example.test:8443/file"),
      rejectsIpLiteral: rejectsChatgptFileUrl("https://127.0.0.1/file"),
      acceptsPublicIpv4: isPublicChatgptFileAddress("8.8.8.8"),
      rejectsLoopbackIpv4: !isPublicChatgptFileAddress("127.0.0.1"),
      rejectsPrivateIpv4: !isPublicChatgptFileAddress("10.0.0.1"),
      acceptsPublicIpv6: isPublicChatgptFileAddress("2606:4700:4700::1111"),
      rejectsLoopbackIpv6: !isPublicChatgptFileAddress("::1"),
      rejectsMappedIpv4: !isPublicChatgptFileAddress("::ffff:127.0.0.1")
    }
  });
  await shutdownOneDriveServer();
  await new Promise((resolve, reject) => {
    process.stdout.write(`${probeOutput}\n`, (error) => error ? reject(error) : resolve());
  });
  process.exit(0);
}

function probe(profile, authMode = "noauth") {
  const env = {
    ...process.env,
    ONEDRIVE_TOOL_PROFILE: profile,
    ONEDRIVE_MCP_AUTH_MODE: authMode
  };
  if (authMode === "oauth") {
    env.ONEDRIVE_MCP_OAUTH_API_CLIENT_ID = "6e97d01c-edf8-43fe-bf69-bb494ae22513";
    env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET = "tool-profile-test-secret";
    delete env.ONEDRIVE_MCP_OAUTH_API_CLIENT_SECRET_FILE;
  }
  const result = spawnSync(process.execPath, [scriptPath, "--probe"], {
    env,
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
  const chatgptOauth = probe("chatgpt", "oauth");
  const maxChatgptToolListBytes = 38 * 1024;
  const expectedNoauthSchemes = [{ type: "noauth" }];
  const expectedOauthSchemes = [{
    type: "oauth2",
    scopes: [
      "api://6e97d01c-edf8-43fe-bf69-bb494ae22513/access_as_user"
    ]
  }];
  const fullNames = new Set(full.names);
  const compatibilityNames = new Set(["search", "fetch", "onedrive_open_files", "onedrive_preview_actions", "onedrive_upload_file", "onedrive_export_file", "onedrive_permanent_delete"]);
  assert(full.count === 84, "Full profile must preserve the 84-tool contract.", full);
  assert(chatgpt.count === 21, "ChatGPT profile must expose the reviewed 21-tool surface.", chatgpt);
  assert(chatgptOauth.count === 21 && JSON.stringify(chatgptOauth.names) === JSON.stringify(chatgpt.names), "OAuth must expose the same reviewed ChatGPT tool surface.", chatgptOauth);
  assert(chatgpt.security.topLevel.every((schemes) => JSON.stringify(schemes) === JSON.stringify(expectedNoauthSchemes)), "Every noauth ChatGPT tool must advertise the standard top-level security scheme.", chatgpt.security);
  assert(chatgpt.security.mirrored.every((schemes) => JSON.stringify(schemes) === JSON.stringify(expectedNoauthSchemes)), "Every noauth ChatGPT tool must retain the compatibility security-scheme mirror.", chatgpt.security);
  assert(chatgptOauth.security.topLevel.every((schemes) => JSON.stringify(schemes) === JSON.stringify(expectedOauthSchemes)), "Every OAuth ChatGPT tool must advertise the exact standard top-level security scheme.", chatgptOauth.security);
  assert(chatgptOauth.security.mirrored.every((schemes) => JSON.stringify(schemes) === JSON.stringify(expectedOauthSchemes)), "Every OAuth ChatGPT tool must retain the compatibility security-scheme mirror.", chatgptOauth.security);
  assert(JSON.stringify(chatgptOauth.annotations) === JSON.stringify(chatgpt.annotations), "OAuth must preserve every reviewed ChatGPT impact annotation.", { noauth: chatgpt.annotations, oauth: chatgptOauth.annotations });
  assert(chatgpt.names.every((name) => fullNames.has(name) || compatibilityNames.has(name)), "ChatGPT profile may add only the reviewed compatibility tools.", chatgpt.names);
  assert(chatgpt.names.includes("search") && chatgpt.names.includes("fetch"), "ChatGPT profile must expose standard search and fetch tools.", chatgpt.names);
  assert(chatgpt.names.includes("onedrive_upload_file") && chatgpt.names.includes("onedrive_export_file") && chatgpt.names.includes("onedrive_restore_deleted") && !chatgpt.names.includes("onedrive_permanent_delete"), "ChatGPT profile must expose upload, remote export, and recycle-bin restore while hiding irreversible permanent deletion.", chatgpt.names);
  assert(!chatgpt.names.includes("onedrive_find") && !chatgpt.names.includes("onedrive_get_info") && !chatgpt.names.includes("onedrive_read_text"), "ChatGPT profile must not advertise redundant slow retrieval tools.", chatgpt.names);
  assert(!chatgpt.names.includes("onedrive_preview") && !chatgpt.names.includes("onedrive_recent") && !chatgpt.names.includes("onedrive_office_search"), "ChatGPT profile must not advertise redundant retrieval helpers.", chatgpt.names);
  assert(!chatgpt.names.includes("onedrive_word_get_document") && !chatgpt.names.includes("onedrive_excel_get_workbook") && !chatgpt.names.includes("onedrive_powerpoint_get_presentation"), "ChatGPT profile must use the bounded fetch extractor instead of redundant high-volume Office reads.", chatgpt.names);
  assert(!full.names.includes("search") && !full.names.includes("fetch") && !full.names.includes("onedrive_open_files") && !full.names.includes("onedrive_preview_actions") && !full.names.includes("onedrive_upload_file") && !full.names.includes("onedrive_export_file") && !full.names.includes("onedrive_permanent_delete"), "ChatGPT compatibility tools must not change the immutable full tool contract.", full.names);
  assert(JSON.stringify(chatgpt.compatibility.search?.inputSchema) === JSON.stringify({ type: "object", required: ["query"], properties: { query: { type: "string", minLength: 1 } }, additionalProperties: false }), "search must keep the exact company-knowledge input contract.", chatgpt.compatibility.search);
  assert(JSON.stringify(chatgpt.compatibility.fetch?.inputSchema) === JSON.stringify({ type: "object", required: ["id"], properties: { id: { type: "string", minLength: 1 } }, additionalProperties: false }), "fetch must keep the exact company-knowledge input contract.", chatgpt.compatibility.fetch);
  assert(chatgpt.compatibility.search?.outputSchema?.required?.includes("results"), "search must advertise the standard results output.", chatgpt.compatibility.search);
  assert(chatgpt.compatibility.fetch?.outputSchema?.required?.includes("text"), "fetch must advertise readable text output.", chatgpt.compatibility.fetch);
  const searchResultSchema = chatgpt.compatibility.search?.outputSchema?.properties?.results?.items?.properties || {};
  const fetchResultSchema = chatgpt.compatibility.fetch?.outputSchema?.properties || {};
  for (const field of ["name", "path", "parent", "type", "webUrl"]) {
    assert(field in searchResultSchema, `search must advertise optional ${field} identity metadata.`, searchResultSchema);
    assert(field in fetchResultSchema, `fetch must advertise optional ${field} identity metadata.`, fetchResultSchema);
  }
  assert("itemId" in fetchResultSchema, "fetch must expose the stable source item ID when id is a progressive continuation.", fetchResultSchema);
  assert(chatgpt.compatibility.search?.annotations?.readOnlyHint === true && chatgpt.compatibility.fetch?.annotations?.readOnlyHint === true, "search/fetch must remain read-only.", chatgpt.compatibility);
  assert(chatgpt.compatibility.onedrive_open_files?.annotations?.readOnlyHint === true && chatgpt.compatibility.onedrive_preview_actions?.annotations?.readOnlyHint === true, "Combined reads and action previews must be advertised as read-only.", chatgpt.compatibility);
  assert(chatgpt.compatibility.onedrive_preview_actions?.annotations?.openWorldHint === false && chatgpt.compatibility.onedrive_preview_actions?.annotations?.destructiveHint === false, "Action previews must not be classified as publishing or destructive.", chatgpt.compatibility.onedrive_preview_actions);
  assert(chatgpt.compatibility.onedrive_create_sharing_link?.annotations?.openWorldHint === true, "Live sharing-link creation must advertise open-world impact.", chatgpt.compatibility.onedrive_create_sharing_link);
  assert(JSON.stringify(chatgpt.compatibility.onedrive_upload_file?._meta?.["openai/fileParams"]) === JSON.stringify(["sourceFile"]), "ChatGPT upload must advertise its file parameter.", chatgpt.compatibility.onedrive_upload_file);
  assert(chatgpt.compatibility.onedrive_upload_file?.annotations?.destructiveHint === true && chatgpt.compatibility.onedrive_export_file?.annotations?.destructiveHint === true, "Upload replacement and remote export must advertise destructive impact.", chatgpt.compatibility);
  assert(chatgpt.oversized.truncated === true && chatgpt.oversized.boundedBytes <= 1024 * 1024, "Oversized ChatGPT tool results must be bounded below the response cap.", chatgpt.oversized);
  assert(chatgpt.bytes <= maxChatgptToolListBytes, "ChatGPT tools/list payload must stay at or below the 38 KiB discovery budget.", chatgpt);
  assert(chatgptOauth.bytes <= maxChatgptToolListBytes, "OAuth ChatGPT tools/list payload must stay at or below the 38 KiB discovery budget.", chatgptOauth);
  assert(chatgpt.bytes < full.bytes * 0.15, "ChatGPT tools/list payload must remain at least 85% smaller than full.", { full, chatgpt });
  assert(chatgpt.officeTransformBytes <= 4096, "ChatGPT Office transform descriptor must remain compact.", chatgpt);
  assert(chatgpt.instructions.length > 0 && chatgpt.instructions.length <= 1400, "Server instructions must be present and bounded.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("onedrive_open_files once") && !chatgpt.instructions.includes("matching structured read tool"), "ChatGPT server instructions must use the combined exact-file read path.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("onedrive_preview_actions") && chatgpt.instructions.includes("identity-free access counts"), "ChatGPT server instructions must route previews through the read-only batch path.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("multiple related documents") && chatgpt.instructions.includes("whole natural-language request once"), "ChatGPT instructions must keep multi-document discovery in one search call.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("Create folders directly") && chatgpt.instructions.includes("no dryRun, confirmed, or previewToken fields"), "ChatGPT instructions must match the direct create-folder schema.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("execute and verify one mutation at a time") && chatgpt.instructions.includes("obtain a fresh preview"), "ChatGPT instructions must serialize dependent mutations and refresh their guards.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("not authentication credentials"), "ChatGPT instructions must classify returned OneDrive identifiers accurately.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("prefer user-visible paths") && chatgpt.instructions.includes("asynchronous acceptance immediately"), "ChatGPT instructions must prefer path selectors and avoid inline copy polling.", chatgpt.instructions);
  assert(chatgpt.instructions.includes("onedrive_export_file") && chatgpt.instructions.includes("PDF or text copy saved in OneDrive"), "ChatGPT instructions must route remote document exports through the guarded export tool.", chatgpt.instructions);
  assert(Object.values(chatgpt.chatgptFileUrlPolicy || {}).every(Boolean), "ChatGPT file URLs must remain host-agnostic while rejecting unsafe URL and network forms.", chatgpt.chatgptFileUrlPolicy);
  const createFolderSchema = chatgpt.compatibility.onedrive_create_folder?.inputSchema;
  assert(createFolderSchema && !["dryRun", "confirmed", "previewToken"].some((field) => field in (createFolderSchema.properties || {})), "Create-folder must not advertise preview-only arguments.", createFolderSchema);
  assert(createFolderSchema?.properties?.parentPath?.description?.includes("Prefer this whenever known") && createFolderSchema?.properties?.parentItemId?.description?.includes("Use only when no parent path is available"), "Create-folder must steer ChatGPT to a user-visible parent path.", createFolderSchema);
  assert(chatgpt.metadata.find((tool) => tool.name === "onedrive_create_folder")?.description.includes("do not send dryRun, confirmed, or previewToken"), "Create-folder metadata must explicitly describe its direct conflict-safe call.", chatgpt.metadata);
  const copySchema = chatgpt.compatibility.onedrive_copy?.inputSchema?.properties || {};
  assert(!("waitForCompletion" in copySchema) && !("timeoutSeconds" in copySchema), "Focused ChatGPT copy must return asynchronous acceptance instead of inline polling.", copySchema);
  assert(copySchema.path?.description?.includes("Prefer this user-visible selector") && copySchema.destinationParentPath?.description?.includes("Prefer this whenever known"), "Focused ChatGPT copy must prefer source and destination paths.", copySchema);
  const deleteSchema = chatgpt.compatibility.onedrive_delete?.inputSchema?.properties || {};
  assert(deleteSchema.path?.description?.includes("Prefer this user-visible selector") && deleteSchema.itemId?.description?.includes("Use only when no path is available"), "Recycle-bin moves must prefer known paths over opaque IDs.", deleteSchema);
  const focusedMetadata = new Map(chatgpt.metadata.map((tool) => [tool.name, tool.description || ""]));
  assert(focusedMetadata.get("onedrive_write_text")?.includes("required content field") && focusedMetadata.get("onedrive_write_text")?.includes("never a text field"), "Text writes must name the exact content argument.", focusedMetadata.get("onedrive_write_text"));
  assert(focusedMetadata.get("onedrive_permissions")?.includes("Send only itemId or path") && focusedMetadata.get("onedrive_permissions")?.includes("never send expectedId"), "Permission reads must reject mutation-only argument copying in metadata.", focusedMetadata.get("onedrive_permissions"));
  assert(focusedMetadata.get("onedrive_create_sharing_link")?.includes("Never use for folder creation or access inspection"), "Sharing-link metadata must exclude adjacent read and folder-create intents.", focusedMetadata.get("onedrive_create_sharing_link"));
  assert(focusedMetadata.get("onedrive_delete")?.includes("Never use for read-only inspection") && focusedMetadata.get("onedrive_delete")?.includes("already recycled item"), "Recycle-bin metadata must exclude read-only and already-deleted probes.", focusedMetadata.get("onedrive_delete"));
  assert(focusedMetadata.get("onedrive_restore_deleted")?.includes("Never call again after restoration"), "Restore metadata must prevent duplicate previews after a successful restore.", focusedMetadata.get("onedrive_restore_deleted"));
  const revokeSchema = chatgpt.compatibility.onedrive_revoke_permission?.inputSchema?.properties || {};
  assert(revokeSchema.itemId?.description?.includes("not an authentication credential") && revokeSchema.permissionId?.description?.includes("not an auth credential") && revokeSchema.previewToken?.description?.includes("not an auth credential"), "Revoke identifiers and preview proof must be accurately classified for the host safety layer.", revokeSchema);
  const writeSchema = chatgpt.compatibility.onedrive_write_text?.inputSchema?.properties || {};
  assert(writeSchema.remotePath?.description?.includes("including filename") && writeSchema.content?.description === "Full UTF-8 file body." && writeSchema.previewToken?.description?.includes("not an auth credential"), "Focused text writes must retain destination, body, and preview-proof argument guidance.", writeSchema);
  const uploadSchema = chatgpt.compatibility.onedrive_upload_file?.inputSchema?.properties || {};
  assert(uploadSchema.remotePath?.description?.includes("including filename") && uploadSchema.confirmed?.description?.includes("explicit user confirmation") && uploadSchema.previewToken?.description?.includes("not an auth credential"), "Focused uploads must retain destination and guarded-live-call argument guidance.", uploadSchema);
  const exportSchema = chatgpt.compatibility.onedrive_export_file?.inputSchema?.properties || {};
  assert(exportSchema.itemId?.description?.includes("Stable source ID") && exportSchema.remotePath?.description?.includes(".pdf or .txt") && exportSchema.expectedETag?.description?.includes("matching preview") && exportSchema.previewToken?.description?.includes("not an auth credential"), "Focused exports must retain stable source, destination, revision, and preview-proof guidance.", exportSchema);
  assert(!["path", "conflictBehavior", "destinationExpectedId", "destinationExpectedName"].some((field) => field in exportSchema), "Focused exports must remain fail-on-conflict and omit replacement-only fields.", exportSchema);
  assert(chatgpt.compatibility.onedrive_export_file?.annotations?.openWorldHint === true && chatgpt.compatibility.onedrive_export_file?.annotations?.destructiveHint === true, "Remote exports must advertise external write and replacement impact.", chatgpt.compatibility.onedrive_export_file);
  const inviteSchema = chatgpt.compatibility.onedrive_invite_permission?.inputSchema?.properties || {};
  assert(inviteSchema.recipients?.description?.includes("exactly one of email, alias, or objectId") && inviteSchema.expectedName?.description?.includes("provide this or expectedId") && inviteSchema.previewToken?.description?.includes("not an auth credential"), "Focused permission invitations must retain recipient and guarded-live-call argument guidance.", inviteSchema);
  const officeItemSchema = chatgpt.compatibility.onedrive_office_batch_transform?.inputSchema?.properties?.items?.items?.properties || {};
  assert(officeItemSchema.path?.description?.includes("relative to OneDrive root") && officeItemSchema.operations?.description?.includes("onedrive_office_capabilities") && officeItemSchema.expectedId?.description?.includes("expectedName"), "Focused Office batches must retain selector, capabilities-handoff, and identity guidance.", officeItemSchema);
  assert(chatgpt.metadata.every((tool) => /^Use this when\b/u.test(tool.description || "")), "Every focused ChatGPT tool description must begin with a discriminative 'Use this when' cue.", chatgpt.metadata);
  assert(new Set(chatgpt.metadata.map((tool) => tool.description)).size === chatgpt.metadata.length, "Focused ChatGPT tool descriptions must be unique.", chatgpt.metadata);
  assert(chatgpt.metadata.every((tool) => tool.invoking && tool.invoked && tool.invoking.length <= 64 && tool.invoked.length <= 64), "Every focused ChatGPT tool must advertise bounded invocation status text.", chatgpt.metadata);
  assert(chatgpt.metadata.find((tool) => tool.name === "fetch")?.description.includes("continuation ID"), "Fetch metadata must explain progressive continuation behavior.", chatgpt.metadata);
  for (const name of ["onedrive_rename", "onedrive_move", "onedrive_copy", "onedrive_create_sharing_link", "onedrive_revoke_permission"]) {
    const description = chatgpt.metadata.find((tool) => tool.name === name)?.description || "";
    assert(description.includes("Inputs:") && description.includes("expectedId or expectedName") && description.includes("previewToken"), `${name} must expose compact live-call input guidance for deferred ChatGPT schema loading.`, description);
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
    chatgptOauth: { count: chatgptOauth.count, bytes: chatgptOauth.bytes, serverVersion: chatgptOauth.serverVersion },
    reductionPercent: Number(((1 - chatgpt.bytes / full.bytes) * 100).toFixed(1))
  }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message, details: error.details }, null, 2));
  process.exit(1);
}
