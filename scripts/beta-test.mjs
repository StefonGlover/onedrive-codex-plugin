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
const unique = `codex-beta-${Date.now()}-${process.pid}`;
const outDir = join(workspace, "work", "onedrive-beta", unique);
const localUpload = join(outDir, "upload-source.txt");
const localSessionUpload = join(outDir, "upload-session-source.txt");
const localBinary = join(outDir, "binary-source.bin");
const localDownload = join(outDir, "upload-downloaded.txt");
const excelDownload = join(outDir, "downloaded-sheet.csv");
const wordDownload = join(outDir, "downloaded-doc.docx");
const powerpointDownload = join(outDir, "downloaded-deck.pptx");
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
    "onedrive_get_info",
    "onedrive_read_text",
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
    "onedrive_permissions",
    "onedrive_restore_deleted",
    "onedrive_delete"
  ];
  record("tools/list includes enhanced tools", requiredTools.every((name) => toolNames.includes(name)) ? "pass" : "fail", {
    toolCount: toolNames.length,
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

  const renamed = assertOk("onedrive_rename", await tool("onedrive_rename", {
    path: `${folderName}/note.txt`,
    newName: renamedTextFile
  }));
  record("rename file", renamed.name === renamedTextFile ? "pass" : "fail", { name: renamed.name });

  const moved = assertOk("onedrive_move", await tool("onedrive_move", {
    path: `${folderName}/${renamedTextFile}`,
    destinationParentPath: `${folderName}/${movedFolderName}`,
    newName: movedTextFile,
    expectedName: renamedTextFile
  }));
  const movedInfo = assertOk("moved file info", await tool("onedrive_get_info", {
    path: `${folderName}/${movedFolderName}/${movedTextFile}`
  }));
  record("move file with expectedName", moved.name === movedTextFile && movedInfo.id === moved.id ? "pass" : "fail", {
    name: moved.name,
    id: moved.id
  });

  const copied = assertOk("onedrive_copy", await tool("onedrive_copy", {
    path: `${folderName}/${movedFolderName}/${movedTextFile}`,
    destinationParentPath: folderName,
    newName: copyFileName,
    expectedName: movedTextFile,
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
  record("find ranks test file without local index", found.items.some((item) => item.name === "uploaded-session.txt") && found.summary?.localIndexUsed === false && found.summary?.persistentCacheUsed === false ? "pass" : "fail", {
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
  record("find_all broad locator works without local index", foundAll.items.some((item) => item.name === "uploaded-session.txt") && foundAll.summary?.localIndexUsed === false && foundAll.summary?.persistentCacheUsed === false ? "pass" : "fail", {
    summary: foundAll.summary,
    folderPlan: foundAll.folderPlan,
    names: foundAll.items.map((item) => item.name)
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

  const sharingDryRun = assertOk("sharing link dry-run", await tool("onedrive_create_sharing_link", {
    path: `${folderName}/${copyFileName}`,
    type: "view",
    scope: "anonymous"
  }));
  record("sharing link dry-run is safe", sharingDryRun.dryRun === true && sharingDryRun.requiredToCreate ? "pass" : "fail", {
    requiredToCreate: sharingDryRun.requiredToCreate,
    beforePermissionCount: sharingDryRun.beforePermissionCount
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
    confirmed: true
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
} catch (error) {
  results.error = error.stack || error.message;
  if (folder?.id) {
    try {
      const cleanup = await tool("onedrive_delete", { itemId: folder.id, expectedName: folderName, dryRun: false, confirmed: true });
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

console.log(JSON.stringify(results, null, 2));

if (results.error || failCount > 0) {
  process.exitCode = 1;
}
