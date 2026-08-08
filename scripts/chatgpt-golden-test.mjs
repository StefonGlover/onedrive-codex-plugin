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
  { prompt: "Read the budget workbook you found", tool: "fetch", cues: ["read", "returned by onedrive_read_actions"] },
  { prompt: "Open 2026 Family Budgeting.xlsx and Annual Report.pdf", tool: "onedrive_open_files", cues: ["exact filenames", "one read-only call"] },
  { prompt: "Preview renaming this workbook, copying it, and creating a view link", tool: "onedrive_preview_actions", cues: ["preview", "read-only batch", "onedrive_commit_actions"] },
  { prompt: "List the root and search for insurance while checking this folder's permissions", tool: "onedrive_read_actions", cues: ["folder listings", "descriptive searches", "permission inspections", "concurrently"] },
  { prompt: "I approve all three previewed actions; apply them", tool: "onedrive_commit_actions", cues: ["approves one or more actions", "stops on the first error", "partial completion"] },
  { prompt: "What structured Excel edits are supported?", tool: "onedrive_office_capabilities", cues: ["supported structured operations"] },
  { prompt: "Update cells in these two Excel workbooks", tool: "onedrive_office_batch_transform", cues: ["structured edits", "preview"] },
  { prompt: "Upload this attached PDF to OneDrive", tool: "onedrive_upload_file", cues: ["chatgpt-provided file", "upload"] },
  { prompt: "Export this Word document as a PDF beside the source in OneDrive", tool: "onedrive_export_file", cues: ["converted to pdf or plain text", "saved back in onedrive", "preview"] },
  { prompt: "Create a new markdown file with this full content", tool: "onedrive_write_text", cues: ["create or fully replace", "required content field", "never a text field"] },
  { prompt: "Change only one line in this existing text file", tool: "onedrive_patch_text", cues: ["targeted", "preserving", "expectedetag", "previewtoken"] },
  { prompt: "Create a folder named Receipts under Documents", tool: "onedrive_create_folder", cues: ["direct conflict-safe create", "do not send dryrun"] },
  { prompt: "Give these named people edit access", tool: "onedrive_invite_permission", cues: ["specific named recipients"] },
  { prompt: "Move this file to the recycle bin", tool: "onedrive_delete", cues: ["only for that intent", "active onedrive item", "never use for read-only inspection", "permanent deletion"] },
  { prompt: "Restore this item from the recycle bin", tool: "onedrive_restore_deleted", cues: ["only for that intent", "while it is in the onedrive recycle bin", "never call again after restoration"] }
];

const ambiguityPairs = [
  ["onedrive_preview_actions", "onedrive_commit_actions"],
  ["onedrive_read_actions", "onedrive_open_files"],
  ["onedrive_upload_file", "onedrive_export_file"],
  ["onedrive_write_text", "onedrive_patch_text"],
  ["onedrive_commit_actions", "onedrive_invite_permission"],
  ["onedrive_create_folder", "onedrive_commit_actions"],
  ["onedrive_commit_actions", "onedrive_read_actions"],
  ["onedrive_delete", "onedrive_restore_deleted"]
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
  assert(initialized.result.instructions.includes("onedrive_open_files once"), "ChatGPT instructions must describe the combined exact-file read sequence.", initialized.result.instructions);
  assert(initialized.result.instructions.includes("onedrive_preview_actions once") && initialized.result.instructions.includes("onedrive_commit_actions"), "ChatGPT instructions must describe the guarded batch preview/commit sequence.", initialized.result.instructions);
  assert(initialized.result.instructions.includes("dependency order") && initialized.result.instructions.includes("proofs can become stale"), "ChatGPT instructions must prevent stale guards in dependent mutation sequences.", initialized.result.instructions);
  assert(initialized.result.instructions.includes("Create folders directly"), "ChatGPT instructions must match the create-folder contract.", initialized.result.instructions);
  assert(initialized.result.instructions.includes("Prefer user-visible paths") && initialized.result.instructions.includes("verified stable results"), "ChatGPT instructions must avoid false credential routing and redundant mutation readbacks.", initialized.result.instructions);
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
