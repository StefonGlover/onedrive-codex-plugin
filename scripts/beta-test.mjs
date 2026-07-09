#!/usr/bin/env node

import { spawn } from "node:child_process";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const serverPath = join(pluginRoot, "mcp", "server.mjs");
const workspace = process.cwd();
const cliArgs = Object.fromEntries(process.argv.slice(2).map((arg) => {
  const [key, ...rest] = arg.replace(/^--/, "").split("=");
  return [key, rest.length ? rest.join("=") : true];
}));
const keepWork = Boolean(cliArgs["keep-work"]);
const doctorOnly = Boolean(cliArgs["doctor-only"]);
const cleanupStale = Boolean(cliArgs["cleanup-stale"]);
const cleanupConfirmed = Boolean(cliArgs.confirmed || cliArgs.delete);
const cleanupStaleDays = Number(cliArgs["stale-days"] || 1);
const tenantMatrix = cliArgs["tenant-matrix"];
const tenantMatrixLive = Boolean(cliArgs["tenant-matrix-live"]);
const unique = `codex-beta-${Date.now()}-${process.pid}`;
const outDir = join(workspace, "work", "onedrive-beta", unique);
const localUpload = join(outDir, "upload-source.txt");
const localSessionUpload = join(outDir, "upload-session-source.txt");
const localBinary = join(outDir, "binary-source.bin");
const localDownload = join(outDir, "upload-downloaded.txt");
const auditExport = join(outDir, "audit-export.jsonl");
const excelDownload = join(outDir, "downloaded-sheet.csv");
const wordDownload = join(outDir, "downloaded-doc.docx");
const powerpointDownload = join(outDir, "downloaded-deck.pptx");
const exportPdfDownload = join(outDir, "exported-doc.pdf");
const exportTextDownload = join(outDir, "exported-doc.txt");
const updateCheckout = join(outDir, "update-checkout.txt");
const updateManifest = join(outDir, "update-checkout.json");
const blockedSyncDownload = join(homedir(), "Library", "CloudStorage", "OneDrive-Personal", `${unique}-blocked-download.txt`);
const blockedSyncUpload = join(homedir(), "Library", "CloudStorage", "OneDrive-Personal", `${unique}-blocked-upload.txt`);
const folderName = `Codex OneDrive Plugin Beta Test ${unique}`;
const movedFolderName = "Moved";
const renamedTextFile = "note-renamed.txt";
const movedTextFile = "note-moved.txt";
const copyFileName = "note-copy.txt";
const content = [
  `OneDrive plugin beta test token: ${unique}`,
  "This file was created by Codex and should be deleted during cleanup.",
  ""
].join("\n");

await mkdir(outDir, { recursive: true });
await writeFile(localUpload, `Uploaded through onedrive_upload: ${unique}\n`, "utf8");
await writeFile(localSessionUpload, Buffer.alloc(400 * 1024, `session-${unique}\n`));
await writeFile(localBinary, Buffer.from([0, 1, 2, 3, 4, 0, 255, 128]));
await rm(localDownload, { force: true });
await rm(excelDownload, { force: true });
await rm(wordDownload, { force: true });
await rm(powerpointDownload, { force: true });
await rm(exportPdfDownload, { force: true });
await rm(exportTextDownload, { force: true });
await rm(updateCheckout, { force: true });
await rm(updateManifest, { force: true });

const child = spawn(process.execPath, [serverPath], {
  cwd: workspace,
  stdio: ["pipe", "pipe", "pipe"]
});

let nextId = 1;
let buffer = "";
const pending = new Map();
const stderr = [];
let childExited = false;
const childExit = new Promise((resolve) => {
  child.once("exit", () => {
    childExited = true;
    resolve();
  });
});

child.stderr.on("data", (chunk) => stderr.push(chunk.toString()));
child.stdout.on("data", (chunk) => {
  buffer += chunk.toString();
  for (;;) {
    const newline = buffer.indexOf("\n");
    if (newline === -1) break;
    const line = buffer.slice(0, newline).trim();
    buffer = buffer.slice(newline + 1);
    if (!line) continue;
    const message = JSON.parse(line);
    const waiter = pending.get(message.id);
    if (waiter) {
      pending.delete(message.id);
      clearTimeout(waiter.timeout);
      waiter.resolve(message);
    }
  }
});

function request(method, params = {}) {
  const id = nextId++;
  const payload = { jsonrpc: "2.0", id, method, params };
  const promise = new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      if (pending.has(id)) {
        pending.delete(id);
        reject(new Error(`Timed out waiting for ${method}`));
      }
    }, 120_000);
    pending.set(id, { resolve, reject, timeout });
  });
  child.stdin.write(`${JSON.stringify(payload)}\n`);
  return promise;
}

async function tool(name, args = {}) {
  const response = await request("tools/call", { name, arguments: args });
  if (response.error) throw new Error(response.error.message);
  const text = response.result?.content?.[0]?.text ?? "";
  let value = text;
  try {
    value = JSON.parse(text);
  } catch {
    // Keep string responses as-is.
  }
  return { isError: Boolean(response.result?.isError), value, raw: response };
}

function assertOk(name, result) {
  if (result.isError) {
    throw new Error(`${name} returned error: ${typeof result.value === "string" ? result.value : JSON.stringify(result.value)}`);
  }
  return result.value;
}

async function previewTokenFor(name, args = {}) {
  const previewArgs = { ...args };
  delete previewArgs.dryRun;
  delete previewArgs.confirmed;
  delete previewArgs.previewToken;
  const preview = assertOk(`${name} preview`, await tool(name, previewArgs));
  if (!preview.previewToken) throw new Error(`${name} preview did not return a previewToken.`);
  return preview.previewToken;
}

async function toolWithPreview(name, args = {}) {
  const previewToken = await previewTokenFor(name, args);
  return await tool(name, { ...args, previewToken });
}

function betaFolderLooksStale(item = {}, cutoffMs) {
  if (!item.folder || !String(item.name || "").startsWith("Codex OneDrive Plugin Beta Test codex-beta-")) return false;
  const modified = Date.parse(item.lastModifiedDateTime || item.createdDateTime || "");
  return Number.isFinite(modified) ? modified <= cutoffMs : true;
}

async function cleanupStaleBetaFolders() {
  const cutoffMs = Date.now() - Math.max(0, cleanupStaleDays) * 24 * 60 * 60 * 1000;
  const scan = assertOk("cleanup stale scan", await tool("onedrive_scan", {
    nameContains: "Codex OneDrive Plugin Beta Test codex-beta-",
    includeFiles: false,
    includeFolders: true,
    maxItems: Number(cliArgs["cleanup-max-items"] || 5000),
    maxFolders: Number(cliArgs["cleanup-max-folders"] || 1000),
    maxResults: Number(cliArgs["cleanup-max-results"] || 100),
    format: "full"
  }));
  const candidates = (scan.items || []).filter((item) => betaFolderLooksStale(item, cutoffMs));
  const deleted = [];
  if (cleanupConfirmed) {
    for (const candidate of candidates) {
      const result = assertOk("cleanup stale delete", await toolWithPreview("onedrive_delete", {
        itemId: candidate.id,
        expectedName: candidate.name,
        dryRun: false,
        confirmed: true
      }));
      deleted.push(result.deleted);
    }
  }
  const details = {
    mode: cleanupConfirmed ? "delete" : "dry-run",
    cutoff: new Date(cutoffMs).toISOString(),
    staleDays: cleanupStaleDays,
    candidateCount: candidates.length,
    deletedCount: deleted.length,
    candidates: candidates.map((item) => ({ id: item.id, name: item.name, lastModifiedDateTime: item.lastModifiedDateTime })),
    deleted: deleted.map((item) => ({ id: item.id, name: item.name }))
  };
  results.cleanup = details;
  record("cleanup stale beta folders", "pass", details);
}

async function runOneTenantMatrixEntry(tenant) {
  const childArgs = [fileURLToPath(import.meta.url)];
  if (!tenantMatrixLive) childArgs.push("--doctor-only");
  if (keepWork) childArgs.push("--keep-work");
  const entry = spawn(process.execPath, childArgs, {
    cwd: workspace,
    env: { ...process.env, ONEDRIVE_TENANT: tenant },
    stdio: ["ignore", "pipe", "pipe"]
  });
  let stdout = "";
  let stderrText = "";
  entry.stdout.on("data", (chunk) => {
    stdout += chunk.toString();
  });
  entry.stderr.on("data", (chunk) => {
    stderrText += chunk.toString();
  });
  const exitCode = await new Promise((resolve) => entry.once("exit", (code) => resolve(code)));
  let parsed = null;
  try {
    parsed = JSON.parse(stdout);
  } catch {
    // Keep raw stdout below.
  }
  return {
    tenant,
    ok: exitCode === 0 && parsed?.summary?.failCount === 0 && !parsed?.error,
    exitCode,
    summary: parsed?.summary,
    configuredTenant: parsed?.checks?.find((check) => check.name === "configured and token available")?.details?.tenant,
    error: parsed?.error,
    stdout: parsed ? undefined : stdout.trim(),
    stderr: stderrText.trim() || undefined
  };
}

async function runTenantMatrix() {
  const tenants = String(tenantMatrix === true ? "common,consumers,organizations" : tenantMatrix)
    .split(",")
    .map((tenant) => tenant.trim())
    .filter(Boolean);
  const matrix = [];
  for (const tenant of tenants) {
    matrix.push(await runOneTenantMatrixEntry(tenant));
  }
  return {
    mode: tenantMatrixLive ? "live-beta" : "doctor-only",
    tenants,
    ok: matrix.every((entry) => entry.ok),
    matrix
  };
}

const results = {
  unique,
  folderName,
  checks: [],
  cleanup: null,
  stderr: ""
};

function record(name, status, details = {}) {
  results.checks.push({ name, status, details });
}

let folder = null;

if (tenantMatrix) {
  child.kill("SIGTERM");
  await Promise.race([
    childExit,
    new Promise((resolve) => setTimeout(resolve, 2_000))
  ]);
  if (!childExited) child.kill("SIGKILL");
  const matrix = await runTenantMatrix();
  await rm(outDir, { recursive: true, force: true });
  console.log(JSON.stringify(matrix, null, 2));
  process.exitCode = matrix.ok ? 0 : 1;
  process.exit();
}

try {
  const init = await request("initialize", {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: { name: "onedrive-beta-test", version: "1.0.0" }
  });
  record("initialize", init.result?.serverInfo?.name === "onedrive" ? "pass" : "fail", init.result);

  const listed = await request("tools/list");
  const toolNames = listed.result.tools.map((entry) => entry.name).sort();
  const requiredTools = [
    "onedrive_config",
    "onedrive_auth_device_start",
    "onedrive_auth_device_poll",
    "onedrive_logout",
    "onedrive_doctor",
    "onedrive_me",
    "onedrive_drive",
    "onedrive_presets",
    "onedrive_list",
    "onedrive_list_all",
    "onedrive_scan",
    "onedrive_search",
    "onedrive_search_all",
    "onedrive_find",
    "onedrive_find_all",
    "onedrive_delta",
    "onedrive_sync_status",
    "onedrive_cache_refresh",
    "onedrive_cache_clear",
    "onedrive_get_info",
    "onedrive_read_text",
    "onedrive_preview",
    "onedrive_download",
    "onedrive_download_excel",
    "onedrive_download_word",
    "onedrive_download_powerpoint",
    "onedrive_export_pdf",
    "onedrive_export_text",
    "onedrive_upload",
    "onedrive_write_text",
    "onedrive_create_folder",
    "onedrive_rename",
    "onedrive_move",
    "onedrive_copy",
    "onedrive_create_sharing_link",
    "onedrive_invite_permission",
    "onedrive_revoke_permission",
    "onedrive_batch_revoke_permissions",
    "onedrive_permissions",
    "onedrive_batch_get_info",
    "onedrive_batch_permissions",
    "onedrive_batch_download",
    "onedrive_batch_delete",
    "onedrive_batch_move",
    "onedrive_update_file",
    "onedrive_recent",
    "onedrive_large_files",
    "onedrive_duplicates",
    "onedrive_shared_by_me",
    "onedrive_public_links",
    "onedrive_restore_deleted",
    "onedrive_audit_recent",
    "onedrive_audit_export",
    "onedrive_audit_clear",
    "onedrive_delete"
  ];
  const uniqueToolNames = new Set(toolNames);
  record("tools/list includes enhanced tools", requiredTools.every((name) => toolNames.includes(name)) && uniqueToolNames.size === toolNames.length ? "pass" : "fail", {
    toolCount: toolNames.length,
    uniqueToolCount: uniqueToolNames.size,
    missing: requiredTools.filter((name) => !toolNames.includes(name))
  });

  const config = assertOk("onedrive_config", await tool("onedrive_config", { checkToken: true }));
  record("configured and token available", config.clientIdConfigured && config.accessTokenAvailable ? "pass" : "fail", {
    tenant: config.tenant,
    scopes: config.scopes,
    keychainTokenConfigured: config.keychainTokenConfigured,
    accessTokenAvailable: config.accessTokenAvailable
  });

  const doctor = assertOk("onedrive_doctor", await tool("onedrive_doctor", { checkRootList: true, rootListLimit: 3 }));
  record("doctor health check passes", doctor.ok === true && doctor.summary?.fail === 0 ? "pass" : "fail", {
    status: doctor.status,
    summary: doctor.summary,
    checks: doctor.checks?.map((check) => ({ name: check.name, status: check.status }))
  });

  if (cleanupStale) {
    await cleanupStaleBetaFolders();
  } else if (doctorOnly) {
    results.specialMode = "doctor-only";
    record("doctor-only mode completed", "pass", { tenant: config.tenant });
  } else {
  const me = assertOk("onedrive_me", await tool("onedrive_me"));
  record("profile read", me.userPrincipalName || me.mail ? "pass" : "fail", {
    displayName: me.displayName,
    userPrincipalName: me.userPrincipalName,
    mail: me.mail
  });

  const drive = assertOk("onedrive_drive", await tool("onedrive_drive"));
  record("drive metadata read", drive.id && drive.driveType ? "pass" : "fail", {
    driveType: drive.driveType,
    name: drive.name,
    quotaState: drive.quota?.state
  });

  const presets = assertOk("onedrive_presets", await tool("onedrive_presets"));
  record("path presets available", presets.pathPresets?.documents === "Documents" && presets.pathPresets?.desktop === "Desktop" ? "pass" : "fail", {
    pathPresets: presets.pathPresets
  });

  const presetTraversal = await tool("onedrive_get_info", {
    preset: "documents",
    relativePath: "../Pictures"
  });
  record("preset path traversal is refused", presetTraversal.isError && String(presetTraversal.value).includes("unsafe path segment") ? "pass" : "fail", {
    response: presetTraversal.value
  });

  folder = assertOk("onedrive_create_folder", await tool("onedrive_create_folder", {
    name: folderName,
    conflictBehavior: "fail"
  }));
  record("create folder", folder.id && folder.folder ? "pass" : "fail", { id: folder.id, name: folder.name });

  const movedFolder = assertOk("onedrive_create_folder", await tool("onedrive_create_folder", {
    parentPath: folderName,
    name: movedFolderName,
    conflictBehavior: "fail"
  }));
  record("create nested folder", movedFolder.id && movedFolder.folder ? "pass" : "fail", { id: movedFolder.id, name: movedFolder.name });

  const written = assertOk("onedrive_write_text", await tool("onedrive_write_text", {
    remotePath: `${folderName}/note.txt`,
    content,
    conflictBehavior: "fail"
  }));
  record("create text file", written.item?.id && written.bytesUploaded === Buffer.byteLength(content, "utf8") ? "pass" : "fail", {
    name: written.item?.name,
    bytesUploaded: written.bytesUploaded
  });

  const readBack = assertOk("onedrive_read_text", await tool("onedrive_read_text", {
    path: `${folderName}/note.txt`,
    maxBytes: 10000
  }));
  record("read text file", readBack.content === content ? "pass" : "fail", { bytes: Buffer.byteLength(readBack.content || "", "utf8") });

  await assertOk("write csv", await tool("onedrive_write_text", {
    remotePath: `${folderName}/sheet.csv`,
    content: `name,value\n${unique},42\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write docx-like", await tool("onedrive_write_text", {
    remotePath: `${folderName}/doc.docx`,
    content: `Document helper test ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write pptx-like", await tool("onedrive_write_text", {
    remotePath: `${folderName}/deck.pptx`,
    content: `PowerPoint helper test ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write duplicate A", await tool("onedrive_write_text", {
    remotePath: `${folderName}/duplicate.txt`,
    content: `Duplicate helper test ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write duplicate B", await tool("onedrive_write_text", {
    remotePath: `${folderName}/${movedFolderName}/duplicate.txt`,
    content: `Duplicate helper test ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write batch move A", await tool("onedrive_write_text", {
    remotePath: `${folderName}/batch-move-a.txt`,
    content: `Batch move A ${unique}\n`,
    conflictBehavior: "fail"
  }));
  await assertOk("write batch move B", await tool("onedrive_write_text", {
    remotePath: `${folderName}/batch-move-b.txt`,
    content: `Batch move B ${unique}\n`,
    conflictBehavior: "fail"
  }));
  const batchDeleteTarget = assertOk("write batch delete target", await tool("onedrive_write_text", {
    remotePath: `${folderName}/batch-delete.txt`,
    content: `Batch delete ${unique}\n`,
    conflictBehavior: "fail"
  }));

  const excel = assertOk("download_excel", await tool("onedrive_download_excel", {
    path: `${folderName}/sheet.csv`,
    localPath: excelDownload
  }));
  const word = assertOk("download_word", await tool("onedrive_download_word", {
    path: `${folderName}/doc.docx`,
    localPath: wordDownload
  }));
  const powerpoint = assertOk("download_powerpoint", await tool("onedrive_download_powerpoint", {
    path: `${folderName}/deck.pptx`,
    localPath: powerpointDownload
  }));
  record("office download helpers", excel.bytesWritten > 0 && word.bytesWritten > 0 && powerpoint.bytesWritten > 0 ? "pass" : "fail", {
    excel: excel.localPath,
    word: word.localPath,
    powerpoint: powerpoint.localPath
  });

  const preview = assertOk("preview text file", await tool("onedrive_preview", {
    path: `${folderName}/note.txt`,
    maxBytes: 40
  }));
  record("preview returns bounded content", preview.preview?.includes("OneDrive plugin beta test") && preview.truncated === true ? "pass" : "fail", {
    source: preview.source,
    bytes: preview.bytes,
    truncated: preview.truncated,
    preview: preview.preview
  });

  const exportPdf = await tool("onedrive_export_pdf", {
    path: `${folderName}/doc.docx`,
    localPath: exportPdfDownload,
    overwrite: true
  });
  record("export_pdf succeeds or reports Graph conversion limit", !exportPdf.isError
    ? exportPdf.value.bytesWritten > 0
      ? "pass"
      : "fail"
    : /format|convert|conversion|unsupported|notSupported|cannotOpenFile|NotAcceptable|UnknownError|406|invalid/i.test(String(exportPdf.value))
      ? "pass"
      : "fail", {
    response: exportPdf.value,
    note: "The beta creates a synthetic .docx text payload; Graph may reject conversion for non-Office binary content."
  });

  const exportText = await tool("onedrive_export_text", {
    path: `${folderName}/doc.docx`,
    localPath: exportTextDownload,
    overwrite: true
  });
  record("export_text succeeds or reports Graph conversion limit", !exportText.isError
    ? exportText.value.bytesWritten > 0
      ? "pass"
      : "fail"
    : /format|convert|conversion|unsupported|notSupported|cannotOpenFile|NotAcceptable|UnknownError|406|invalid/i.test(String(exportText.value))
      ? "pass"
      : "fail", {
    response: exportText.value,
    note: "The beta creates a synthetic .docx text payload; Graph may reject conversion for non-Office binary content."
  });

  const renamed = assertOk("onedrive_rename", await tool("onedrive_rename", {
    path: `${folderName}/note.txt`,
    newName: renamedTextFile,
    expectedName: "note.txt",
    dryRun: false,
    confirmed: true
  }));
  const renamedItem = renamed.renamed || renamed;
  record("rename file", renamedItem.name === renamedTextFile && renamed.confirmed === true ? "pass" : "fail", { name: renamedItem.name });

  const moved = assertOk("onedrive_move", await tool("onedrive_move", {
    path: `${folderName}/${renamedTextFile}`,
    destinationParentPath: `${folderName}/${movedFolderName}`,
    newName: movedTextFile,
    expectedName: renamedTextFile,
    dryRun: false,
    confirmed: true
  }));
  const movedItem = moved.moved || moved;
  const movedInfo = assertOk("moved file info", await tool("onedrive_get_info", {
    path: `${folderName}/${movedFolderName}/${movedTextFile}`
  }));
  record("move file with expectedName", movedItem.name === movedTextFile && movedInfo.id === movedItem.id && moved.confirmed === true ? "pass" : "fail", {
    name: movedItem.name,
    id: movedItem.id
  });

  const copied = assertOk("onedrive_copy", await tool("onedrive_copy", {
    path: `${folderName}/${movedFolderName}/${movedTextFile}`,
    destinationParentPath: folderName,
    newName: copyFileName,
    expectedName: movedTextFile,
    dryRun: false,
    confirmed: true,
    waitForCompletion: true,
    timeoutSeconds: 90
  }));
  record("copy file accepted", copied.accepted && copied.monitorUrl ? "pass" : "fail", {
    responseStatus: copied.status,
    monitorComplete: copied.monitor?.complete
  });

  const copiedInfo = assertOk("copied file info", await tool("onedrive_get_info", {
    path: `${folderName}/${copyFileName}`
  }));
  record("copied file resolves", copiedInfo.name === copyFileName ? "pass" : "fail", { id: copiedInfo.id, name: copiedInfo.name });

  const uploaded = assertOk("onedrive_upload", await tool("onedrive_upload", {
    localPath: localUpload,
    remotePath: `${folderName}/uploaded.txt`,
    conflictBehavior: "fail"
  }));
  record("upload local file", uploaded.item?.name === "uploaded.txt" && uploaded.bytesUploaded > 0 && uploaded.uploadMode === "simple" ? "pass" : "fail", {
    name: uploaded.item?.name,
    bytesUploaded: uploaded.bytesUploaded,
    uploadMode: uploaded.uploadMode
  });

  const sessionUploaded = assertOk("onedrive_upload session", await tool("onedrive_upload", {
    localPath: localSessionUpload,
    remotePath: `${folderName}/uploaded-session.txt`,
    conflictBehavior: "fail",
    uploadMode: "session",
    chunkSize: 327680
  }));
  record("upload session file", sessionUploaded.item?.name === "uploaded-session.txt" && sessionUploaded.uploadMode === "session" ? "pass" : "fail", {
    name: sessionUploaded.item?.name,
    bytesUploaded: sessionUploaded.bytesUploaded,
    uploadMode: sessionUploaded.uploadMode,
    chunkSize: sessionUploaded.chunkSize
  });

  const binaryUploaded = assertOk("binary upload", await tool("onedrive_upload", {
    localPath: localBinary,
    remotePath: `${folderName}/binary.bin`,
    conflictBehavior: "fail"
  }));
  record("upload binary file", binaryUploaded.item?.name === "binary.bin" ? "pass" : "fail", { bytesUploaded: binaryUploaded.bytesUploaded });

  const binaryRead = await tool("onedrive_read_text", {
    path: `${folderName}/binary.bin`,
    maxBytes: 1000
  });
  record("binary read refused", binaryRead.isError && String(binaryRead.value).includes("likely binary") ? "pass" : "fail", {
    response: binaryRead.value
  });

  const downloaded = assertOk("onedrive_download", await tool("onedrive_download", {
    path: `${folderName}/uploaded.txt`,
    localPath: localDownload,
    overwrite: false
  }));
  const downloadedContent = await readFile(localDownload, "utf8");
  record("download file", downloaded.bytesWritten === Buffer.byteLength(downloadedContent, "utf8") && downloadedContent.includes(unique) ? "pass" : "fail", {
    localPath: downloaded.localPath,
    bytesWritten: downloaded.bytesWritten
  });

  const checkout = assertOk("update_file checkout", await tool("onedrive_update_file", {
    mode: "checkout",
    remotePath: `${folderName}/uploaded.txt`,
    localPath: updateCheckout,
    manifestPath: updateManifest,
    overwriteLocal: true,
    overwriteManifest: true
  }));
  await writeFile(updateCheckout, `${downloadedContent}Edited through update_file commit: ${unique}\n`, "utf8");
  const committed = assertOk("update_file commit", await tool("onedrive_update_file", {
    mode: "commit",
    remotePath: `${folderName}/uploaded.txt`,
    localPath: updateCheckout,
    manifestPath: updateManifest,
    createBackup: true,
    verify: true
  }));
  const updatedReadBack = assertOk("updated file readback", await tool("onedrive_read_text", {
    path: `${folderName}/uploaded.txt`,
    maxBytes: 10000
  }));
  record("update_file checkout and commit", checkout.mode === "checkout"
    && committed.mode === "commit"
    && committed.backup?.bytesWritten > 0
    && updatedReadBack.content.includes("Edited through update_file commit") ? "pass" : "fail", {
    checkout: { localPath: checkout.localPath, manifestPath: checkout.manifestPath },
    commit: { backupBytes: committed.backup?.bytesWritten, verified: committed.verified?.name }
  });

  const batchMoveDryRun = assertOk("batch move dry-run", await tool("onedrive_batch_move", {
    items: [
      { path: `${folderName}/batch-move-a.txt`, expectedName: "batch-move-a.txt" },
      { path: `${folderName}/batch-move-b.txt`, expectedName: "batch-move-b.txt" }
    ],
    destinationParentPath: `${folderName}/${movedFolderName}`
  }));
  record("batch_move dry-run previews every item", batchMoveDryRun.dryRun === true && batchMoveDryRun.results?.length === 2 ? "pass" : "fail", {
    count: batchMoveDryRun.count
  });

  const batchMoved = assertOk("batch move live", await tool("onedrive_batch_move", {
    items: [
      { path: `${folderName}/batch-move-a.txt`, expectedName: "batch-move-a.txt" },
      { path: `${folderName}/batch-move-b.txt`, expectedName: "batch-move-b.txt" }
    ],
    destinationParentPath: `${folderName}/${movedFolderName}`,
    dryRun: false,
    confirmed: true
  }));
  record("batch_move live succeeds", batchMoved.confirmed === true && batchMoved.results?.length === 2 ? "pass" : "fail", {
    count: batchMoved.count,
    names: batchMoved.results?.map((entry) => entry.moved?.name || entry.item?.name || entry.name)
  });

  const batchDeleteDryRun = assertOk("batch delete dry-run", await tool("onedrive_batch_delete", {
    items: [{ itemId: batchDeleteTarget.item.id, expectedName: "batch-delete.txt" }]
  }));
  record("batch_delete dry-run previews item", batchDeleteDryRun.dryRun === true && batchDeleteDryRun.results?.length === 1 ? "pass" : "fail", {
    count: batchDeleteDryRun.count
  });

  const batchDeleted = assertOk("batch delete live", await toolWithPreview("onedrive_batch_delete", {
    items: [{ itemId: batchDeleteTarget.item.id, expectedName: "batch-delete.txt" }],
    dryRun: false,
    confirmed: true
  }));
  record("batch_delete live succeeds", batchDeleted.confirmed === true && batchDeleted.results?.length === 1 ? "pass" : "fail", {
    count: batchDeleted.count
  });

  const listedFolder = assertOk("onedrive_list", await tool("onedrive_list", { path: folderName, limit: 20 }));
  const childNames = listedFolder.items.map((item) => item.name).sort();
  record("list folder children compactly", childNames.includes(copyFileName) && childNames.includes("uploaded.txt") && listedFolder.items.every((item) => item.type) ? "pass" : "fail", {
    childNames,
    sample: listedFolder.items[0]
  });

  const allFolder = assertOk("onedrive_list_all", await tool("onedrive_list_all", { path: folderName, pageSize: 2, maxItems: 20 }));
  record("list_all follows pagination", allFolder.items.length >= 5 && allFolder.count === allFolder.items.length ? "pass" : "fail", {
    count: allFolder.count,
    truncated: allFolder.truncated
  });

  const scanned = assertOk("onedrive_scan", await tool("onedrive_scan", {
    path: folderName,
    nameContains: "uploaded",
    includeFolders: false,
    maxItems: 50,
    maxFolders: 10,
    maxResults: 10
  }));
  record("scan finds nested test files", scanned.items.some((item) => item.name === "uploaded.txt") && scanned.items.every((item) => item.remotePath) ? "pass" : "fail", {
    summary: scanned.summary,
    names: scanned.items.map((item) => item.name)
  });

  const found = assertOk("onedrive_find", await tool("onedrive_find", {
    query: "uploaded session",
    folderHints: [folderName],
    maxResults: 5,
    scanMaxItems: 50,
    scanMaxFolders: 10
  }));
  record("find ranks test file with cache acceleration", found.items.some((item) => item.name === "uploaded-session.txt") && found.summary?.localIndexUsed === false && found.summary?.persistentCacheUsed === true ? "pass" : "fail", {
    summary: found.summary,
    top: found.items[0]
  });

  const foundAll = assertOk("onedrive_find_all", await tool("onedrive_find_all", {
    query: "uploaded session",
    folderHints: [folderName],
    maxResults: 20,
    scanMaxItems: 100,
    scanMaxFolders: 20
  }));
  record("find_all broad locator works with cache acceleration", foundAll.items.some((item) => item.name === "uploaded-session.txt") && foundAll.summary?.localIndexUsed === false && foundAll.summary?.persistentCacheUsed === true ? "pass" : "fail", {
    summary: foundAll.summary,
    folderPlan: foundAll.folderPlan,
    names: foundAll.items.map((item) => item.name)
  });

  const syncStatusBefore = assertOk("sync status", await tool("onedrive_sync_status", { includeSamples: true }));
  record("sync_status reports cache state", Number.isInteger(syncStatusBefore.itemCount) && Array.isArray(syncStatusBefore.samples) ? "pass" : "fail", {
    itemCount: syncStatusBefore.itemCount,
    sampleCount: syncStatusBefore.samples?.length,
    deltaLinkAvailable: syncStatusBefore.deltaLinkAvailable
  });

  const cacheRefresh = assertOk("cache refresh", await tool("onedrive_cache_refresh", {
    path: folderName,
    mode: "scan",
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("cache_refresh scans selected folder", cacheRefresh.cache?.itemCount >= 1 && ["scan", "scan+delta"].includes(cacheRefresh.mode) ? "pass" : "fail", {
    mode: cacheRefresh.mode,
    itemCount: cacheRefresh.cache?.itemCount,
    scanned: cacheRefresh.scan?.summary?.itemsScanned,
    scanRoot: cacheRefresh.cache?.scanRoot
  });

  const cacheCleared = assertOk("cache clear", await tool("onedrive_cache_clear"));
  record("cache_clear empties metadata cache", cacheCleared.itemCount === 0 ? "pass" : "fail", {
    itemCount: cacheCleared.itemCount,
    updatedAt: cacheCleared.updatedAt
  });

  const cacheRefreshAfterClear = assertOk("cache refresh after clear", await tool("onedrive_cache_refresh", {
    path: folderName,
    mode: "scan",
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("cache_refresh rebuilds cache after clear", cacheRefreshAfterClear.cache?.itemCount >= 1 ? "pass" : "fail", {
    mode: cacheRefreshAfterClear.mode,
    itemCount: cacheRefreshAfterClear.cache?.itemCount,
    scanned: cacheRefreshAfterClear.scan?.summary?.itemsScanned
  });

  const recent = assertOk("recent files", await tool("onedrive_recent", { limit: 10 }));
  record("recent files call succeeds", Array.isArray(recent.items) && recent.count >= 0 ? "pass" : "fail", {
    count: recent.count,
    sample: recent.items?.[0]
  });

  const largeFiles = assertOk("large files", await tool("onedrive_large_files", {
    path: folderName,
    minBytes: 1,
    limit: 10,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("large_files finds test files", largeFiles.items?.some((item) => item.name === "uploaded-session.txt") ? "pass" : "fail", {
    count: largeFiles.count,
    names: largeFiles.items?.map((item) => item.name)
  });

  const duplicates = assertOk("duplicates", await tool("onedrive_duplicates", {
    path: folderName,
    limit: 10,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("duplicates finds duplicate test files", duplicates.groups?.some((group) => group.items?.some((item) => item.name === "duplicate.txt")) ? "pass" : "fail", {
    count: duplicates.count,
    groups: duplicates.groups?.map((group) => ({ key: group.key, count: group.count, names: group.items.map((item) => item.name) }))
  });

  const blockedDownload = await tool("onedrive_download", {
    path: `${folderName}/uploaded.txt`,
    localPath: blockedSyncDownload,
    overwrite: true
  });
  record("download refuses local OneDrive sync path", blockedDownload.isError && String(blockedDownload.value).includes("local OneDrive sync folder") ? "pass" : "fail", {
    response: blockedDownload.value
  });

  const blockedUpload = await tool("onedrive_upload", {
    localPath: blockedSyncUpload,
    remotePath: `${folderName}/blocked-upload.txt`
  });
  record("upload refuses local OneDrive sync path", blockedUpload.isError && String(blockedUpload.value).includes("local OneDrive sync folder") ? "pass" : "fail", {
    response: blockedUpload.value
  });

  const search = assertOk("onedrive_search", await tool("onedrive_search", { query: unique, limit: 10 }));
  record("search unique token", Array.isArray(search.items) ? "pass" : "fail", {
    resultCount: search.items?.length ?? 0,
    note: "OneDrive search indexing may be eventually consistent for freshly created files."
  });

  const searchAllResult = await tool("onedrive_search_all", { query: unique, pageSize: 2, maxItems: 10 });
  if (searchAllResult.isError && String(searchAllResult.value).includes("Substrate Search")) {
    record("search_all tolerates transient Microsoft Search backend", "pass", {
      response: searchAllResult.value,
      note: "OneDrive search can return transient Substrate Search errors for freshly created files."
    });
  } else {
    const searchAll = assertOk("onedrive_search_all", searchAllResult);
    record("search_all call succeeds", Array.isArray(searchAll.items) ? "pass" : "fail", {
      count: searchAll.count,
      truncated: searchAll.truncated,
      note: "Fresh OneDrive search results can lag indexing."
    });
  }

  const delta = assertOk("onedrive_delta", await tool("onedrive_delta", { itemId: folder.id, pageSize: 20, maxItems: 50 }));
  record("delta sync call succeeds", Array.isArray(delta.items) && (delta.nextLink || delta.deltaLink || delta.count >= 0) ? "pass" : "fail", {
    count: delta.count,
    hasNextLink: Boolean(delta.nextLink),
    hasDeltaLink: Boolean(delta.deltaLink)
  });

  const permissionAudit = assertOk("permissions audit", await tool("onedrive_permissions", { path: `${folderName}/${copyFileName}` }));
  record("permission audit call succeeds", Array.isArray(permissionAudit.permissions) ? "pass" : "fail", {
    count: permissionAudit.count,
    item: permissionAudit.item?.name
  });

  const batchInfo = assertOk("batch get info", await tool("onedrive_batch_get_info", {
    items: [{ itemId: copiedInfo.id }, { path: `${folderName}/uploaded.txt` }]
  }));
  record("batch get_info returns per-item results", batchInfo.items?.length === 2 && batchInfo.items.every((entry) => !entry.error) ? "pass" : "fail", {
    count: batchInfo.count,
    names: batchInfo.items?.map((entry) => entry.name || entry.item?.name)
  });

  const batchPermissions = assertOk("batch permissions", await tool("onedrive_batch_permissions", {
    items: [{ itemId: copiedInfo.id }, { path: `${folderName}/uploaded.txt` }]
  }));
  record("batch permissions returns per-item results", batchPermissions.items?.length === 2 && batchPermissions.items.every((entry) => !entry.error) ? "pass" : "fail", {
    count: batchPermissions.count
  });

  const batchDownloaded = assertOk("batch download", await tool("onedrive_batch_download", {
    items: [{ path: `${folderName}/uploaded.txt` }],
    destinationFolder: outDir,
    overwrite: true
  }));
  record("batch download succeeds", batchDownloaded.results?.[0]?.bytesWritten > 0 ? "pass" : "fail", {
    count: batchDownloaded.count,
    first: batchDownloaded.results?.[0]
  });

  const sharingDryRun = assertOk("sharing link dry-run", await tool("onedrive_create_sharing_link", {
    path: `${folderName}/${copyFileName}`,
    type: "view",
    scope: "anonymous",
    password: `${unique}-link-password`,
    expirationDateTime: "2099-01-01T00:00:00Z"
  }));
  record("sharing link dry-run is safe", sharingDryRun.dryRun === true && sharingDryRun.requiredToCreate && sharingDryRun.wouldCreate?.passwordProvided === true && sharingDryRun.wouldCreate?.expirationDateTime === "2099-01-01T00:00:00Z" ? "pass" : "fail", {
    requiredToCreate: sharingDryRun.requiredToCreate,
    beforePermissionCount: sharingDryRun.beforePermissionCount,
    wouldCreate: sharingDryRun.wouldCreate
  });

  const inviteRecipient = me.mail || me.userPrincipalName;
  const inviteDryRun = assertOk("invite permission dry-run", await tool("onedrive_invite_permission", {
    path: `${folderName}/${copyFileName}`,
    recipients: [{ email: inviteRecipient }],
    role: "read",
    password: `${unique}-invite-password`,
    expirationDateTime: "2099-01-01T00:00:00Z"
  }));
  record("invite permission dry-run is safe and silent by default", inviteDryRun.dryRun === true
    && inviteDryRun.requiredToInvite
    && inviteDryRun.wouldInvite?.sendInvitation === false
    && inviteDryRun.wouldInvite?.requireSignIn === true
    && inviteDryRun.wouldInvite?.passwordProvided === true
    && inviteDryRun.wouldInvite?.expirationDateTime === "2099-01-01T00:00:00Z" ? "pass" : "fail", {
    requiredToInvite: inviteDryRun.requiredToInvite,
    beforePermissionCount: inviteDryRun.beforePermissionCount,
    wouldInvite: inviteDryRun.wouldInvite
  });

  const inviteNeedsConfirmation = assertOk("invite permission requires confirmation", await tool("onedrive_invite_permission", {
    itemId: copiedInfo.id,
    recipients: [{ email: inviteRecipient }],
    expectedName: copyFileName,
    dryRun: false
  }));
  record("invite live action requires confirmation", inviteNeedsConfirmation.requiredToInvite && inviteNeedsConfirmation.confirmed === false ? "pass" : "fail", {
    requiredToInvite: inviteNeedsConfirmation.requiredToInvite
  });

  const inviteMissingExpected = assertOk("invite permission requires expected identity", await tool("onedrive_invite_permission", {
    itemId: copiedInfo.id,
    recipients: [{ email: inviteRecipient }],
    dryRun: false,
    confirmed: true
  }));
  record("invite live action requires expected identity", inviteMissingExpected.requiredToInvite?.includes("expectedName or expectedId") ? "pass" : "fail", {
    requiredToInvite: inviteMissingExpected.requiredToInvite
  });

  const inviteLive = await toolWithPreview("onedrive_invite_permission", {
    itemId: copiedInfo.id,
    recipients: [{ email: inviteRecipient }],
    role: "read",
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  });
  if (inviteLive.isError && /owner|same user|already|not.*supported|invalid.*recipient|cannot.*share/i.test(String(inviteLive.value))) {
    record("invite permission live self-grant handled", "pass", {
      response: inviteLive.value,
      note: "The live beta uses the signed-in user as recipient to avoid emailing or exposing access; Microsoft Graph can reject redundant self-grants depending on drive/account type."
    });
  } else {
    const inviteLiveValue = assertOk("invite permission live", inviteLive);
    record("invite permission live silent grant succeeds", inviteLiveValue.confirmed === true && inviteLiveValue.invite?.sendInvitation === false && inviteLiveValue.permissionDiff ? "pass" : "fail", {
      invite: inviteLiveValue.invite,
      diff: inviteLiveValue.permissionDiff
    });
  }

  const sharingLive = assertOk("sharing link live", await toolWithPreview("onedrive_create_sharing_link", {
    itemId: copiedInfo.id,
    type: "view",
    scope: "anonymous",
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  record("sharing link live creates permission diff", sharingLive.confirmed === true && sharingLive.permission?.id && sharingLive.permissionDiff ? "pass" : "fail", {
    permissionId: sharingLive.permission?.id,
    diff: sharingLive.permissionDiff
  });

  const sharedByMe = assertOk("shared by me audit", await tool("onedrive_shared_by_me", {
    path: folderName,
    limit: 20,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  const publicLinks = assertOk("public links audit", await tool("onedrive_public_links", {
    path: folderName,
    limit: 20,
    maxItems: 50,
    maxFolders: 10,
    maxDepth: 5
  }));
  record("sharing audit tools find live anonymous link", sharedByMe.items?.some((entry) => entry.item?.id === copiedInfo.id)
    && publicLinks.items?.some((entry) => entry.item?.id === copiedInfo.id) ? "pass" : "fail", {
    sharedByMe: { count: sharedByMe.count, ids: sharedByMe.items?.map((entry) => entry.item?.id) },
    publicLinks: { count: publicLinks.count, ids: publicLinks.items?.map((entry) => entry.item?.id) }
  });

  const revokeDryRun = assertOk("revoke permission dry-run", await tool("onedrive_revoke_permission", {
    itemId: copiedInfo.id,
    permissionId: sharingLive.permission.id,
    expectedName: copyFileName
  }));
  record("revoke permission dry-run is safe", revokeDryRun.dryRun === true && revokeDryRun.requiredToRevoke ? "pass" : "fail", {
    beforePermissionCount: revokeDryRun.beforePermissionCount,
    requiredToRevoke: revokeDryRun.requiredToRevoke
  });

  const revoked = assertOk("revoke permission live", await toolWithPreview("onedrive_revoke_permission", {
    itemId: copiedInfo.id,
    permissionId: sharingLive.permission.id,
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  record("revoke permission live removes sharing", revoked.confirmed === true && revoked.permissionDiff?.removed?.some((permission) => permission.id === sharingLive.permission.id) ? "pass" : "fail", {
    permissionId: revoked.permissionId,
    diff: revoked.permissionDiff
  });

  const sharingLiveBatch = assertOk("sharing link live for batch revoke", await toolWithPreview("onedrive_create_sharing_link", {
    itemId: copiedInfo.id,
    type: "view",
    scope: "anonymous",
    expectedName: copyFileName,
    dryRun: false,
    confirmed: true
  }));
  const batchRevoked = assertOk("batch revoke permissions live", await toolWithPreview("onedrive_batch_revoke_permissions", {
    items: [{ itemId: copiedInfo.id, permissionId: sharingLiveBatch.permission.id, expectedId: copiedInfo.id }],
    dryRun: false,
    confirmed: true
  }));
  record("batch revoke permissions live succeeds", batchRevoked.confirmed === true && batchRevoked.results?.length === 1 ? "pass" : "fail", {
    count: batchRevoked.count,
    results: batchRevoked.results
  });

  const auditRecent = assertOk("audit recent", await tool("onedrive_audit_recent", { limit: 50 }));
  record("audit recent includes live mutation entries", auditRecent.entries?.some((entry) => entry.tool === "onedrive_revoke_permission") && auditRecent.entries?.some((entry) => entry.tool === "onedrive_batch_revoke_permissions") ? "pass" : "fail", {
    count: auditRecent.count,
    tools: auditRecent.entries?.map((entry) => entry.tool)
  });

  const auditExported = assertOk("audit export", await tool("onedrive_audit_export", {
    localPath: auditExport,
    overwrite: true
  }));
  record("audit export writes local JSONL", auditExported.bytesWritten > 0 && auditExported.localPath === auditExport ? "pass" : "fail", {
    localPath: auditExported.localPath,
    bytesWritten: auditExported.bytesWritten
  });

  const expectedMismatch = await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: "wrong-name",
    dryRun: false,
    confirmed: true
  });
  record("delete expectedName mismatch refused", expectedMismatch.isError && String(expectedMismatch.value).includes("expected item named") ? "pass" : "fail", {
    response: expectedMismatch.value
  });

  const deleteNeedsConfirmation = assertOk("onedrive_delete requires confirmation", await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: folderName,
    dryRun: false
  }));
  record("delete live action requires confirmation", deleteNeedsConfirmation.requiredToDelete && deleteNeedsConfirmation.confirmed === false ? "pass" : "fail", {
    requiredToDelete: deleteNeedsConfirmation.requiredToDelete
  });

  const dryRun = assertOk("onedrive_delete dry-run", await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: folderName,
    dryRun: true
  }));
  record("delete dry-run safety", dryRun.dryRun === true && dryRun.wouldDelete?.id === folder.id ? "pass" : "fail", {
    wouldDelete: dryRun.wouldDelete?.name
  });

  const deleted = assertOk("onedrive_delete", await tool("onedrive_delete", {
    itemId: folder.id,
    expectedName: folderName,
    dryRun: false,
    confirmed: true,
    previewToken: dryRun.previewToken
  }));
  results.cleanup = { deleted: deleted.deleted?.name, id: deleted.deleted?.id };
  record("delete test folder cleanup", deleted.dryRun === false && deleted.deleted?.id === folder.id ? "pass" : "fail", results.cleanup);
  folder = null;

  const restoreDryRun = assertOk("restore deleted dry-run", await tool("onedrive_restore_deleted", {
    itemId: deleted.deleted.id,
    expectedId: deleted.deleted.id
  }));
  record("restore deleted dry-run is safe", restoreDryRun.dryRun === true && restoreDryRun.requiredToRestore ? "pass" : "fail", {
    requiredToRestore: restoreDryRun.requiredToRestore,
    permissionNote: restoreDryRun.permissionNote
  });

  const deletedById = await tool("onedrive_get_info", {
    itemId: deleted.deleted.id,
    includeDeletedItems: true
  });
  record("includeDeletedItems lookup is handled", deletedById.isError || deletedById.value?.id === deleted.deleted.id ? "pass" : "fail", {
    response: deletedById.value,
    note: "Microsoft documents includeDeletedItems as OneDrive Personal-only and itemId-only; some deleted folders may still return itemNotFound."
  });

  const deletedInfo = await tool("onedrive_get_info", { path: folderName });
  record("deleted folder no longer resolves", deletedInfo.isError ? "pass" : "fail", { response: deletedInfo.value });

  const rootDelete = await tool("onedrive_delete", { path: "/", dryRun: true });
  record("root delete is refused", rootDelete.isError && String(rootDelete.value).includes("OneDrive root") ? "pass" : "fail", {
    response: rootDelete.value
  });

  const auditClearNoConfirm = assertOk("audit clear requires confirmation", await tool("onedrive_audit_clear"));
  record("audit_clear requires confirmation", auditClearNoConfirm.requiredToClear && auditClearNoConfirm.confirmed === false ? "pass" : "fail", {
    requiredToClear: auditClearNoConfirm.requiredToClear
  });

  const deviceStart = assertOk("auth device start", await tool("onedrive_auth_device_start", { tenant: "consumers" }));
  record("auth_device_start returns device login metadata", deviceStart.userCode && deviceStart.verificationUri && deviceStart.deviceCodeStoredInMemory === true ? "pass" : "fail", {
    authTenant: deviceStart.authTenant,
    expiresIn: deviceStart.expiresIn,
    verificationUri: deviceStart.verificationUri
  });

  const devicePoll = assertOk("auth device poll pending", await tool("onedrive_auth_device_poll"));
  record("auth_device_poll reports pending authorization safely", devicePoll.authorizationPending === true ? "pass" : "fail", {
    message: devicePoll.message,
    slowDown: devicePoll.slowDown
  });

  const logoutMemoryOnly = assertOk("logout memory only", await tool("onedrive_logout", { deleteKeychainToken: false }));
  record("logout clears memory without deleting Keychain token", logoutMemoryOnly.memoryCleared === true && logoutMemoryOnly.keychainTokenDeleted === false ? "pass" : "fail", logoutMemoryOnly);
  }
} catch (error) {
  results.error = error.stack || error.message;
  if (folder?.id) {
    try {
      const cleanup = await toolWithPreview("onedrive_delete", { itemId: folder.id, expectedName: folderName, dryRun: false, confirmed: true });
      results.cleanup = { attemptedAfterError: true, result: cleanup.value, isError: cleanup.isError };
    } catch (cleanupError) {
      results.cleanup = { attemptedAfterError: true, error: cleanupError.message };
    }
  }
} finally {
  results.stderr = stderr.join("");
  child.stdin.end();
  if (!childExited) child.kill("SIGTERM");
  await Promise.race([
    childExit,
    new Promise((resolve) => setTimeout(resolve, 2_000))
  ]);
  if (!childExited) child.kill("SIGKILL");
}

const passCount = results.checks.filter((check) => check.status === "pass").length;
const failCount = results.checks.filter((check) => check.status === "fail").length;
results.summary = { passCount, failCount, total: results.checks.length };

if (!results.error && failCount === 0 && !keepWork) {
  await rm(outDir, { recursive: true, force: true });
  results.localWorkCleaned = true;
} else {
  results.localWorkDir = outDir;
}

console.log(JSON.stringify(results, null, 2));

if (results.error || failCount > 0) {
  process.exitCode = 1;
}
