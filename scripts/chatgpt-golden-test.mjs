#!/usr/bin/env node

process.env.ONEDRIVE_TOOL_PROFILE = "chatgpt";

const { processMcpMessage, shutdownOneDriveServer } = await import("../mcp/server.mjs");

function assert(condition, message, details = undefined) {
  if (condition) return;
  const error = new Error(message);
  error.details = details;
  throw error;
}

const goldenPrompts = [
  { prompt: "Find my 2026 family budget workbook", tool: "search", cues: ["find", "keywords", "call search separately"] },
  { prompt: "Read the budget workbook you found", tool: "fetch", cues: ["read", "returned by search"] },
  { prompt: "Show the files directly inside Documents/Taxes", tool: "onedrive_list", cues: ["direct children", "known"] },
  { prompt: "What structured Excel edits are supported?", tool: "onedrive_office_capabilities", cues: ["supported structured operations"] },
  { prompt: "Update cells in these two Excel workbooks", tool: "onedrive_office_batch_transform", cues: ["structured edits", "preview"] },
  { prompt: "Upload this attached PDF to OneDrive", tool: "onedrive_upload_file", cues: ["chatgpt-provided file", "upload"] },
  { prompt: "Create a new markdown file with this full content", tool: "onedrive_write_text", cues: ["create or fully replace", "text"] },
  { prompt: "Change only one line in this existing text file", tool: "onedrive_patch_text", cues: ["targeted", "preserving"] },
  { prompt: "Create a folder named Receipts under Documents", tool: "onedrive_create_folder", cues: ["create a new folder"] },
  { prompt: "Rename this file without moving it", tool: "onedrive_rename", cues: ["change the name", "without changing", "newname", "expectedid or expectedname", "previewtoken"] },
  { prompt: "Move this workbook into Archive", tool: "onedrive_move", cues: ["different parent folder", "destinationparentitemid", "expectedid or expectedname", "previewtoken"] },
  { prompt: "Copy this document and keep the original", tool: "onedrive_copy", cues: ["leaving the source in place", "waitforcompletion", "expectedid or expectedname", "previewtoken"] },
  { prompt: "Create a view-only sharing link", tool: "onedrive_create_sharing_link", cues: ["shareable", "type and scope", "previewtoken"] },
  { prompt: "Give these named people edit access", tool: "onedrive_invite_permission", cues: ["specific named recipients"] },
  { prompt: "Remove this sharing permission", tool: "onedrive_revoke_permission", cues: ["revoke", "permissionid", "previewtoken"] },
  { prompt: "Who currently has access to this folder?", tool: "onedrive_permissions", cues: ["inspect who can access"] },
  { prompt: "Move this file to the recycle bin", tool: "onedrive_delete", cues: ["recycle bin", "permanent deletion"] },
  { prompt: "Restore this item from the recycle bin", tool: "onedrive_restore_deleted", cues: ["restore", "recycle-bin"] },
  { prompt: "Permanently delete this item and skip the recycle bin", tool: "onedrive_permanent_delete", cues: ["irreversibly", "without the recycle bin"] }
];

const ambiguityPairs = [
  ["search", "onedrive_list"],
  ["onedrive_write_text", "onedrive_patch_text"],
  ["onedrive_move", "onedrive_copy"],
  ["onedrive_create_sharing_link", "onedrive_invite_permission"],
  ["onedrive_permissions", "onedrive_revoke_permission"],
  ["onedrive_delete", "onedrive_permanent_delete"]
];

try {
  const initialized = await processMcpMessage({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: { protocolVersion: "2025-11-25" }
  });
  const listed = await processMcpMessage({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
  const tools = listed.result.tools || [];
  const byName = new Map(tools.map((tool) => [tool.name, tool]));
  assert(tools.length === goldenPrompts.length, "Golden prompt coverage must match the complete focused ChatGPT tool surface.", { tools: tools.map((tool) => tool.name), prompts: goldenPrompts.map((entry) => entry.tool) });
  assert(initialized.result.instructions.includes("fetch each file"), "ChatGPT instructions must describe the fetch-first Office edit sequence.", initialized.result.instructions);
  assert(!initialized.result.instructions.includes("matching structured read tool"), "ChatGPT instructions must not reference tools absent from the focused profile.", initialized.result.instructions);

  for (const entry of goldenPrompts) {
    const tool = byName.get(entry.tool);
    assert(tool, `Golden prompt targets a missing tool: ${entry.tool}`, entry);
    const description = String(tool.description || "").toLowerCase();
    assert(description.startsWith("use this when"), `Tool ${entry.tool} is missing the required selection cue.`, tool);
    for (const cue of entry.cues) {
      assert(description.includes(cue), `Tool ${entry.tool} does not encode the golden-prompt cue '${cue}'.`, { prompt: entry.prompt, description: tool.description });
    }
    for (const field of ["openai/toolInvocation/invoking", "openai/toolInvocation/invoked"]) {
      const status = tool._meta?.[field];
      assert(typeof status === "string" && status.length > 0 && status.length <= 64, `Tool ${entry.tool} has invalid ${field} status text.`, status);
    }
  }

  for (const [leftName, rightName] of ambiguityPairs) {
    const left = byName.get(leftName);
    const right = byName.get(rightName);
    assert(left && right, "Ambiguity-pair tool is missing.", { leftName, rightName });
    assert(left.description !== right.description, "Ambiguous tools must not share descriptions.", { leftName, rightName });
  }

  console.log(JSON.stringify({
    ok: true,
    toolCount: tools.length,
    goldenPromptCount: goldenPrompts.length,
    ambiguityPairCount: ambiguityPairs.length,
    serverInstructionBytes: Buffer.byteLength(initialized.result.instructions, "utf8")
  }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message, details: error.details }, null, 2));
  process.exitCode = 1;
} finally {
  await shutdownOneDriveServer();
}
