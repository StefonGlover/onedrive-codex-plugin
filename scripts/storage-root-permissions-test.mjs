#!/usr/bin/env node

import { execFileSync } from "node:child_process";
import { chmod, mkdir, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const workRoot = await mkdtemp(join(tmpdir(), "onedrive-storage-root-test-"));
const fixtureRoot = join(workRoot, "fixtures");

process.env.ONEDRIVE_STORAGE_ROOT ||= join(workRoot, "storage");
process.env.ONEDRIVE_OFFICE_PYCACHE_ROOT ||= join(workRoot, "python-cache");
process.env.ONEDRIVE_OFFICE_PYTHON ||= "/usr/bin/python3";

try {
  try {
    await mkdir(process.env.ONEDRIVE_STORAGE_ROOT, { recursive: true });
    await chmod(process.env.ONEDRIVE_STORAGE_ROOT, 0o555);
  } catch (error) {
    if (!new Set(["EACCES", "EPERM"]).has(error?.code)) throw error;
  }

  execFileSync(process.env.ONEDRIVE_OFFICE_PYTHON, [
    join(pluginRoot, "scripts", "office-openxml-test.py"),
    `--emit-fixtures=${fixtureRoot}`
  ], {
    env: { ...process.env, PYTHONPYCACHEPREFIX: join(workRoot, "fixture-pycache") },
    stdio: "ignore"
  });

  const { runOfficeHelper } = await import("../mcp/server.mjs");
  const workbook = await runOfficeHelper({
    action: "inspect",
    inputPath: join(fixtureRoot, "sample.xlsx"),
    kind: "excel",
    maxCells: 100
  });

  if (!workbook?.sheets?.length) throw new Error("Excel fixture inspection returned no worksheets.");
  const synologyEntrypoint = await readFile(join(pluginRoot, "deploy", "synology", "entrypoint.sh"), "utf8");
  if (!synologyEntrypoint.includes("/data/chatgpt-uploads")) {
    throw new Error("Synology entrypoint must pre-create the ChatGPT attachment staging directory for the unprivileged runtime user.");
  }
  process.stdout.write(`${JSON.stringify({
    ok: true,
    storageRoot: process.env.ONEDRIVE_STORAGE_ROOT,
    sheets: workbook.sheets.length,
    synologyChatgptUploadRootReady: true
  }, null, 2)}\n`);
} finally {
  await rm(workRoot, { recursive: true, force: true });
}
