#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { chmod, copyFile, cp, lstat, mkdir, mkdtemp, open, readdir, rename, rm, stat } from "node:fs/promises";
import { constants, existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptsRoot = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(scriptsRoot, "..");
const manifest = JSON.parse(readFileSync(join(pluginRoot, ".codex-plugin", "plugin.json"), "utf8"));
const versionPattern = /^0\.4\.0\+codex\.\d{14}$/;
const excludedTopLevelNames = new Set([".git", "work", "downloads", "onedrive-beta", "node_modules", "dist", "build", "coverage"]);

function parseArgs(argv) {
  const parsed = { confirmed: false, selfCheck: false, syncEvidence: false, target: null };
  for (const argument of argv) {
    if (argument === "--confirmed") {
      if (parsed.confirmed) throw new Error("--confirmed may only be provided once.");
      parsed.confirmed = true;
      continue;
    }
    if (argument === "--self-check") {
      if (parsed.selfCheck) throw new Error("--self-check may only be provided once.");
      parsed.selfCheck = true;
      continue;
    }
    if (argument === "--sync-evidence") {
      if (parsed.syncEvidence) throw new Error("--sync-evidence may only be provided once.");
      parsed.syncEvidence = true;
      continue;
    }
    if (argument.startsWith("--target=")) {
      if (parsed.target) throw new Error("--target may only be provided once.");
      parsed.target = argument.slice("--target=".length);
      if (!parsed.target) throw new Error("--target requires an absolute path.");
      continue;
    }
    throw new Error(`Unknown option: ${argument}`);
  }
  if (parsed.selfCheck && (parsed.confirmed || parsed.syncEvidence || parsed.target)) throw new Error("--self-check cannot be combined with install options.");
  return parsed;
}

function runPrepackage(extraArgs = []) {
  const result = spawnSync(process.execPath, [join(scriptsRoot, "prepackage-check.mjs"), ...extraArgs], {
    cwd: pluginRoot,
    env: process.env,
    encoding: "utf8",
    timeout: 30_000
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`Prepackage verification failed.\n${result.stdout || ""}${result.stderr || ""}`.trim());
  }
  return (result.stdout || "").trim();
}

const args = parseArgs(process.argv.slice(2));
if (!versionPattern.test(manifest.version || "")) throw new Error(`Refusing unexpected plugin version: ${manifest.version}`);
const codexHome = resolve(process.env.CODEX_HOME || join(homedir(), ".codex"));
const cacheRoot = join(codexHome, "plugins", "cache", "personal", "onedrive");
const expectedTarget = join(cacheRoot, manifest.version);
const operationLockPath = join(cacheRoot, `.${manifest.version}.install.lock`);

async function acquireOperationLock() {
  await mkdir(cacheRoot, { recursive: true, mode: 0o700 });
  let handle;
  try {
    handle = await open(operationLockPath, "wx", 0o600);
  } catch (error) {
    if (error?.code === "EEXIST") {
      throw new Error(`Another install or evidence-sync operation is already in progress: ${operationLockPath}`);
    }
    throw error;
  }
  return async () => {
    await handle.close();
    await rm(operationLockPath, { force: true });
  };
}

async function assertInstalledTargetDirectory() {
  let targetStat;
  try {
    targetStat = await lstat(expectedTarget);
  } catch (error) {
    if (error?.code === "ENOENT") throw new Error(`Installed cache target does not exist: ${expectedTarget}`);
    throw error;
  }
  if (!targetStat.isDirectory() || targetStat.isSymbolicLink()) {
    throw new Error(`Installed cache target must be a real directory, not a file or symlink: ${expectedTarget}`);
  }
}

if (args.selfCheck) {
  const checks = {
    releaseVersionAccepted: versionPattern.test(manifest.version),
    staleVersionRejected: !versionPattern.test("0.3.0+codex.20260711223300"),
    duplicateConfirmationRejected: (() => { try { parseArgs(["--confirmed", "--confirmed"]); return false; } catch { return true; } })(),
    unknownOptionRejected: (() => { try { parseArgs(["--overwrite"]); return false; } catch { return true; } })(),
    evidenceModeParsed: parseArgs(["--sync-evidence"]).syncEvidence === true,
    oldCacheExcludedFromTarget: expectedTarget.endsWith(`/onedrive/${manifest.version}`),
    sourceControlExcluded: excludedTopLevelNames.has(".git")
  };
  const ok = Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, version: manifest.version, expectedTarget, checks }, null, 2));
  process.exit(ok ? 0 : 1);
}

const requestedTarget = args.target ? resolve(args.target) : expectedTarget;
if (requestedTarget !== expectedTarget) {
  throw new Error(`Refusing a non-versioned cache target. The exact target must be ${expectedTarget}`);
}
const preview = {
  dryRun: !args.confirmed,
  confirmed: args.confirmed,
  source: pluginRoot,
  version: manifest.version,
  target: expectedTarget,
  targetAlreadyExists: existsSync(expectedTarget),
  oldCachesPreserved: true
};
if (args.syncEvidence) {
  const evidenceFiles = ["qa-report.md", "qa-report.json"];
  const evidencePreview = {
    ...preview,
    mode: "sync-release-evidence",
    evidenceFiles
  };
  if (!args.confirmed) {
    console.log(JSON.stringify({
      ...evidencePreview,
      requiredToSync: `Review this exact existing 0.4.0 target, then rerun with --sync-evidence --confirmed --target=${expectedTarget}. Only qa-report.md and qa-report.json are replaced.`
    }, null, 2));
    process.exit(0);
  }
  runPrepackage();
  const releaseLock = await acquireOperationLock();
  let transactionRoot = null;
  const records = [];
  try {
    await assertInstalledTargetDirectory();
    const installedManifest = JSON.parse(readFileSync(join(expectedTarget, ".codex-plugin", "plugin.json"), "utf8"));
    if (installedManifest.version !== manifest.version) {
      throw new Error(`Installed cache version ${installedManifest.version} does not match source ${manifest.version}.`);
    }
    transactionRoot = await mkdtemp(join(cacheRoot, `.${manifest.version}.evidence-sync-`));
    await chmod(transactionRoot, 0o700);
    for (const [index, file] of evidenceFiles.entries()) {
      const sourcePath = join(pluginRoot, file);
      const targetPath = join(expectedTarget, file);
      const tempPath = join(transactionRoot, `${index}.new`);
      const backupPath = join(transactionRoot, `${index}.old`);
      const sourceStat = await stat(sourcePath);
      await copyFile(sourcePath, tempPath, constants.COPYFILE_EXCL);
      await chmod(tempPath, sourceStat.mode & 0o777);
      records.push({ file, targetPath, tempPath, backupPath, originalMoved: false, replacementInstalled: false });
    }
    for (const record of records) {
      const targetStat = await lstat(record.targetPath);
      if (!targetStat.isFile() || targetStat.isSymbolicLink()) {
        throw new Error(`Evidence target must be a regular file: ${record.targetPath}`);
      }
      await rename(record.targetPath, record.backupPath);
      record.originalMoved = true;
      await rename(record.tempPath, record.targetPath);
      record.replacementInstalled = true;
    }
    const parityEvidence = runPrepackage(["--installed", expectedTarget]);
    await rm(transactionRoot, { recursive: true, force: true });
    transactionRoot = null;
    console.log(JSON.stringify({
      ...evidencePreview,
      dryRun: false,
      synced: true,
      parityVerified: true,
      parityEvidence
    }, null, 2));
  } catch (error) {
    const rollbackErrors = [];
    for (const record of [...records].reverse()) {
      try {
        if (record.replacementInstalled) await rm(record.targetPath, { force: true });
        if (record.originalMoved) await rename(record.backupPath, record.targetPath);
      } catch (rollbackError) {
        rollbackErrors.push(`${record.file}: ${rollbackError.message}`);
      }
    }
    if (transactionRoot) await rm(transactionRoot, { recursive: true, force: true });
    if (rollbackErrors.length) {
      error.message += `\nEvidence rollback also failed: ${rollbackErrors.join("; ")}`;
    }
    throw error;
  } finally {
    await releaseLock();
  }
} else {
if (!args.confirmed) {
  console.log(JSON.stringify({
    ...preview,
    requiredToInstall: `Review this exact target, then rerun with --confirmed --target=${expectedTarget}. Existing targets are never overwritten.`
  }, null, 2));
  process.exit(0);
}
if (existsSync(expectedTarget)) {
  throw new Error(`Versioned cache target already exists; refusing to overwrite it: ${expectedTarget}`);
}

runPrepackage();
const releaseLock = await acquireOperationLock();
let stagingRoot = null;
let targetPublished = false;
try {
  if (existsSync(expectedTarget)) {
    throw new Error(`Versioned cache target already exists; refusing to overwrite it: ${expectedTarget}`);
  }
  stagingRoot = await mkdtemp(join(cacheRoot, `.${manifest.version}.install-`));
  await chmod(stagingRoot, 0o700);
  const entries = (await readdir(pluginRoot, { withFileTypes: true }))
    .filter((entry) => !excludedTopLevelNames.has(entry.name));
  for (const entry of entries) {
    await cp(join(pluginRoot, entry.name), join(stagingRoot, entry.name), {
      recursive: true,
      force: false,
      errorOnExist: true,
      preserveTimestamps: true,
      verbatimSymlinks: true
    });
  }
  const parityEvidence = runPrepackage(["--installed", stagingRoot]);
  if (existsSync(expectedTarget)) {
    throw new Error(`Versioned cache target appeared during staging; refusing to overwrite it: ${expectedTarget}`);
  }
  await rename(stagingRoot, expectedTarget);
  stagingRoot = null;
  targetPublished = true;
  console.log(JSON.stringify({
    ...preview,
    dryRun: false,
    installed: true,
    filesCopied: entries.length,
    parityVerified: true,
    parityEvidence
  }, null, 2));
} catch (error) {
  if (stagingRoot) await rm(stagingRoot, { recursive: true, force: true });
  if (targetPublished) await rm(expectedTarget, { recursive: true, force: true });
  throw error;
} finally {
  await releaseLock();
}
}
