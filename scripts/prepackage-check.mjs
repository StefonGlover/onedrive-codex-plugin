#!/usr/bin/env node

import { existsSync, lstatSync, readFileSync, readlinkSync, statSync } from "node:fs";
import { readdir } from "node:fs/promises";
import { dirname, extname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { ONEDRIVE_TOOL_CONTRACT, compareToolContract } from "./tool-contract.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const args = process.argv.slice(2);
const problems = [];
const selfCheckFlags = args.filter((arg) => arg === "--self-check");
if (selfCheckFlags.length > 1) problems.push("--self-check may only be provided once.");
const selfCheck = selfCheckFlags.length === 1;
const installedValueIndexes = new Set();
const installedFlags = args.filter((arg) => arg === "--installed" || arg.startsWith("--installed="));
if (installedFlags.length > 1) problems.push("--installed may only be provided once.");
const installedIndex = args.indexOf("--installed");
const installedEquals = args.find((arg) => arg.startsWith("--installed="));
const installedRequested = installedIndex >= 0 || installedEquals !== undefined;
const installedArgument = installedEquals !== undefined
  ? installedEquals.slice("--installed=".length)
  : installedIndex >= 0 ? args[installedIndex + 1] : null;
if (installedIndex >= 0 && installedArgument && !installedArgument.startsWith("--")) {
  installedValueIndexes.add(installedIndex + 1);
}
const installedRoot = installedRequested && installedArgument && !installedArgument.startsWith("--")
  ? resolve(installedArgument)
  : null;
if (installedRequested && !installedRoot) {
  problems.push("--installed requires a path argument.");
}
if (selfCheck && installedRequested) problems.push("--self-check cannot be combined with --installed.");
for (const [index, arg] of args.entries()) {
  if (installedValueIndexes.has(index)) continue;
  if (arg === "--self-check") continue;
  if (arg === "--installed" || arg.startsWith("--installed=")) continue;
  problems.push(arg.startsWith("--") ? `Unknown option: ${arg}` : `Unexpected positional argument: ${arg}`);
}
const ignoredPackageDirs = new Set([".git", "work", "downloads", "onedrive-beta", "node_modules", "dist", "build", "coverage"]);
const ignoredPackageFiles = new Set([".DS_Store"]);
const ignoredPackageFileExtensions = new Set([".log", ".tmp", ".temp", ".bak", ".swp"]);
const forbiddenResidueDirs = new Set([".codex", ".pytest_cache", "__pycache__"]);
const sensitivePackageFileNames = new Set([".env", ".env.local", ".env.development", ".env.production", "credentials.json", "token.json"]);
const sensitivePackageFileExtensions = new Set([".key", ".pem", ".p12", ".pfx"]);
const sensitivePackageNamePattern = /(token|secret|credential)/i;
const expectedPluginVersion = /^0\.5\.1\+codex\.\d{14}$/;
const expectedOfficeOperationKinds = {
  word: [
    "replaceText", "setParagraphText", "setParagraphStyle", "insertParagraph", "setTableCell",
    "setContentControlText", "addHyperlink", "addComment", "insertTable", "insertImage", "replaceImage",
    "createContentControl", "deleteContentControl", "createBookmark", "deleteBookmark", "insertTableRow",
    "deleteTableRow", "insertTableColumn", "deleteTableColumn", "setHeaderFooterText", "setSectionProperties"
  ],
  excel: [
    "setCell", "setFormula", "setRange", "clearRange", "setStyle", "setNumberFormat",
    "addConditionalFormat", "setDataValidation", "freezePanes", "setColumnWidth", "addTableRow", "deleteTableRow",
    "setTableTotals", "createChart", "updateChart", "renameSheet", "setDefinedName", "recalculate",
    "addWorksheet", "deleteWorksheet", "addTable", "deleteTable", "mergeRange", "unmergeRange", "sortRange",
    "setAutoFilter", "setHyperlink", "addNote", "deleteNote", "insertImage", "formatChart",
    "setSheetProtection", "refreshPivot"
  ],
  powerpoint: [
    "replaceText", "setShapeText", "setShapeGeometry", "setTableCell", "addTextBox", "deleteShape",
    "setTextStyle", "replaceImage", "setNotes", "duplicateSlide", "deleteSlide", "moveSlide", "addSlide",
    "addImage", "cropImage", "addTable", "insertTableRow", "deleteTableRow", "insertTableColumn",
    "deleteTableColumn", "setShapeAltText", "setZOrder", "groupShapes", "ungroupShape", "applySlideLayout"
  ]
};
const officeOperationToolNames = {
  word: "onedrive_word_batch_update",
  excel: "onedrive_excel_batch_update",
  powerpoint: "onedrive_powerpoint_batch_update"
};
const requiredQaOfflineGates = [
  "Node 20 and Node 26 syntax/self-checks",
  "Exact MCP contract",
  "Plugin manifest/version alignment",
  "Prepackage negative checks",
  "Office Open XML operations",
  "Office security corpus",
  "Genuine package reopening",
  "Mock Microsoft Graph",
  "Read-only OneDrive doctor and tenant checks",
  "Cleanup preview",
  "Whitespace"
];
const textExtensions = new Set([
  ".json", ".md", ".mjs", ".js", ".zsh", ".sh", ".txt", ".example", ".yaml", ".yml"
]);
const requiredFiles = [
  ".codex-plugin/plugin.json",
  ".mcp.json",
  "mcp/server.mjs",
  "mcp/http-server.mjs",
  "mcp/oauth.mjs",
  "mcp/auth-vault.mjs",
  "mcp/semantic-anchors.mjs",
  "mcp/text-patch.mjs",
  "scripts/auth-vault-test.mjs",
  "scripts/benchmark.mjs",
  "scripts/beta-test.mjs",
  "scripts/tool-contract.mjs",
  "scripts/install-versioned-cache.mjs",
  "scripts/requirements-office-test.txt",
  "scripts/mock-graph-test.mjs",
  "scripts/oauth-http-test.mjs",
  "scripts/storage-root-permissions-test.mjs",
  "scripts/run-chatgpt-oauth-tunnel.mjs",
  "scripts/semantic-anchors-test.mjs",
  "scripts/text-patch-test.mjs",
  "scripts/tool-profile-test.mjs",
  "scripts/office-openxml.py",
  "scripts/office-fixture-factory.py",
  "scripts/office-openxml-test.py",
  "scripts/office-real-fixture-test.py",
  "scripts/office-security-test.py",
  "scripts/prepackage-check.mjs",
  "deploy/synology/Dockerfile",
  "deploy/synology/entrypoint.sh",
  "deploy/synology/compose.yaml",
  "deploy/synology/compose.oauth.example.yaml",
  "deploy/synology/README.md",
  "skills/onedrive/SKILL.md",
  "README.md",
  "qa-report.md",
  "qa-report.json",
  ".github/workflows/ci.yml"
];
const absoluteUsersNeedle = "/" + "Users/";
const todoNeedle = "[TO" + "DO:";

function fail(message) {
  problems.push(message);
}

function readJson(path) {
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    fail(`Invalid JSON: ${path}: ${error.message}`);
    return null;
  }
}

function isSensitivePackageEntryName(name) {
  return sensitivePackageFileNames.has(name)
    || sensitivePackageFileExtensions.has(extname(name).toLowerCase())
    || sensitivePackageNamePattern.test(name);
}

async function walk(dir, files = []) {
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    const extension = extname(entry.name).toLowerCase();
    if (isSensitivePackageEntryName(entry.name)) {
      fail(`Sensitive file must not be packaged: ${relative(pluginRoot, path)}`);
      continue;
    }
    if (entry.isDirectory()) {
      if (ignoredPackageDirs.has(entry.name)) continue;
      if (forbiddenResidueDirs.has(entry.name)) {
        fail(`Packaged test/review residue directory found: ${relative(pluginRoot, path)}`);
        continue;
      }
      await walk(path, files);
    } else if (entry.isFile()) {
      if (ignoredPackageFiles.has(entry.name) || ignoredPackageFileExtensions.has(extension)) {
        fail(`Packaged temporary residue file found: ${relative(pluginRoot, path)}`);
        continue;
      }
      files.push(path);
    } else if (entry.isSymbolicLink()) {
      files.push(path);
    }
  }
  return files;
}

function isProbablyText(path) {
  const extension = extname(path).toLowerCase();
  if (textExtensions.has(extension)) return true;
  return [".mcp.json", "plugin.json", "SKILL.md", "README.md"].some((name) => path.endsWith(name));
}

function checkNoAbsoluteLocalPaths(files) {
  for (const file of files) {
    if (!isProbablyText(file)) continue;
    if (statSync(file).size > 1024 * 1024) continue;
    const rel = relative(pluginRoot, file);
    const text = readFileSync(file, "utf8");
    if (text.includes(absoluteUsersNeedle)) fail(`Absolute local user path found in packaged text: ${rel}`);
    if (text.includes(todoNeedle)) fail(`Placeholder marker found in packaged text: ${rel}`);
  }
}

function checkManifest() {
  const manifestPath = join(pluginRoot, ".codex-plugin", "plugin.json");
  const manifest = readJson(manifestPath);
  if (!manifest) return;
  if (manifest.name !== "onedrive") fail(`Unexpected plugin name: ${manifest.name}`);
  if (!expectedPluginVersion.test(manifest.version || "")) {
    fail(`Plugin version must match 0.5.1+codex.<14-digit timestamp>: ${manifest.version}`);
  }
  const readme = readFileSync(join(pluginRoot, "README.md"), "utf8");
  if (!readme.includes(`Release \`${manifest.version}\``)) {
    fail(`README release version does not match plugin.json: ${manifest.version}`);
  }
  if (manifest.interface?.defaultPrompt?.length > 3) fail("interface.defaultPrompt should contain at most 3 entries.");
  for (const field of ["composerIcon", "logo", "logoDark"]) {
    const value = manifest.interface?.[field];
    if (!value) fail(`Missing interface.${field}`);
    else if (!existsSync(join(pluginRoot, value))) fail(`Missing asset for interface.${field}: ${value}`);
  }
  for (const screenshot of manifest.interface?.screenshots || []) {
    if (!screenshot.startsWith("./assets/") || !screenshot.endsWith(".png")) {
      fail(`Screenshot must be a PNG under ./assets/: ${screenshot}`);
    }
    if (!existsSync(join(pluginRoot, screenshot))) fail(`Missing screenshot asset: ${screenshot}`);
  }
}

function currentHeadCommit() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], {
    cwd: pluginRoot,
    encoding: "utf8",
    timeout: 5_000
  });
  if (result.error || result.status !== 0) {
    fail(`Could not resolve the source HEAD commit: ${result.error?.message || result.stderr || `exit ${result.status}`}`);
    return null;
  }
  const commit = (result.stdout || "").trim();
  if (!/^[0-9a-f]{40}$/.test(commit)) {
    fail(`Source HEAD did not resolve to a full 40-character commit: ${commit || "empty"}`);
    return null;
  }
  return commit;
}

function qaSourceCommitMatchesRelease(commit, headCommit) {
  if (commit === headCommit) return true;
  if (!/^[0-9a-f]{40}$/.test(commit || "") || !/^[0-9a-f]{40}$/.test(headCommit || "")) return false;
  const ancestor = spawnSync("git", ["merge-base", "--is-ancestor", commit, headCommit], {
    cwd: pluginRoot,
    encoding: "utf8",
    timeout: 5_000
  });
  if (ancestor.status !== 0) return false;
  const diff = spawnSync("git", ["diff", "--name-only", `${commit}..${headCommit}`], {
    cwd: pluginRoot,
    encoding: "utf8",
    timeout: 5_000
  });
  if (diff.status !== 0) return false;
  const changed = (diff.stdout || "").trim().split("\n").filter(Boolean);
  return changed.length > 0 && changed.every((file) => file === "qa-report.md" || file === "qa-report.json");
}

function qaAlignmentProblems(qa, markdown, pluginVersion, headCommit = null, contentDigest = null) {
  const issues = [];
  const statusValues = new Set(["pass", "pending", "blocked", "fail", "in_progress"]);
  const decisionMatch = String(qa?.decision || "").match(/^\s*(?:beta\s+)?(pass|passed|pending|blocked|fail|failed)\b/i);
  const decisionStatus = decisionMatch?.[1]?.toLowerCase();
  if (qa?.schemaVersion !== 2) issues.push("qa-report.json schemaVersion must be 2.");
  if (qa?.source?.pluginVersion !== pluginVersion) issues.push("qa-report.json plugin version does not match plugin.json.");
  if (!/^[0-9a-f]{40}$/.test(qa?.source?.commit || "")) issues.push("qa-report.json must record a full 40-character source commit.");
  if (headCommit && !qaSourceCommitMatchesRelease(qa?.source?.commit, headCommit)) {
    issues.push("qa-report.json source commit must match HEAD or an ancestor followed only by qa-report.md/qa-report.json evidence changes.");
  }
  if (qa?.contract?.toolCount !== ONEDRIVE_TOOL_CONTRACT.length || qa?.contract?.exact !== true) issues.push("qa-report.json must record the exact 84-tool contract.");
  if (qa?.installedBuild?.version !== pluginVersion) issues.push("qa-report.json installed-build target version does not match plugin.json.");
  if (qa?.installedBuild?.oldCacheOverwritten !== false) issues.push("qa-report.json must attest that the old installed cache was not overwritten.");
  if (!Array.isArray(qa?.blockedCoverage) || qa.blockedCoverage.length < 4) issues.push("qa-report.json must preserve explicit blocked-resource reasons.");
  for (const gate of qa?.offlineGates || []) {
    if (!statusValues.has(gate?.status)) issues.push(`qa-report.json has an invalid offline gate status: ${gate?.status}`);
  }
  if (!markdown.includes(`Decision: ${qa?.decision}`)) issues.push("qa-report.md decision does not match qa-report.json.");
  if (!markdown.includes(`Plugin version: \`${pluginVersion}\``)) issues.push("qa-report.md plugin version does not match plugin.json.");
  if (!markdown.includes(`Tool contract: ${ONEDRIVE_TOOL_CONTRACT.length} exact tool names`)) issues.push("qa-report.md must record the exact 84-tool contract.");
  if (/\b58(?:-tool| tools?)\b/i.test(markdown) || /0\.[13]\.0\+codex\./.test(markdown)) issues.push("qa-report.md contains stale release evidence.");
  if (!decisionStatus) issues.push("qa-report.json decision must begin with Pass, Pending, Blocked, Fail, or the corresponding Beta form.");
  if (decisionStatus === "pass" || decisionStatus === "passed") {
    const requiredPasses = [
      ["source live beta", qa?.liveRuns?.source?.status],
      ["installed live beta", qa?.liveRuns?.installed?.status],
      ["installed build", qa?.installedBuild?.status],
      ["final cleanup", qa?.cleanup?.status],
      ["source/cache parity", qa?.sourceCacheParity?.status]
    ];
    for (const [label, status] of requiredPasses) {
      if (status !== "pass") issues.push(`qa-report.json cannot claim Pass while ${label} is ${status || "missing"}.`);
    }
    if (!/^[0-9a-f]{64}$/.test(qa?.source?.contentDigest || "")) {
      issues.push("qa-report.json Pass must record the 64-character packaged-content digest.");
    } else if (contentDigest && qa.source.contentDigest !== contentDigest) {
      issues.push("qa-report.json packaged-content digest does not match the current source tree.");
    }
    for (const runKind of ["source", "installed"]) {
      const run = qa?.liveRuns?.[runKind];
      if (typeof run?.runId !== "string" || !run.runId.trim()) issues.push(`qa-report.json Pass must record the ${runKind} live run ID.`);
      if (typeof run?.folderName !== "string" || !run.folderName.trim()) issues.push(`qa-report.json Pass must record the ${runKind} live folder name.`);
      if (!Number.isFinite(run?.runtimeMs) || run.runtimeMs <= 0) issues.push(`qa-report.json Pass must record a positive ${runKind} live runtime.`);
      if (run?.toolCoverage?.contract !== ONEDRIVE_TOOL_CONTRACT.length
        || run?.toolCoverage?.exercised !== ONEDRIVE_TOOL_CONTRACT.length - 1
        || run?.toolCoverage?.blocked !== 1) {
        issues.push(`qa-report.json Pass must record ${ONEDRIVE_TOOL_CONTRACT.length - 1} exercised and one explicitly blocked tool for the ${runKind} live run.`);
      }
      if (run?.cleanupVerified !== true) issues.push(`qa-report.json Pass must attest exact cleanup for the ${runKind} live run.`);
    }
    const expectedInstalledQaPath = `$CODEX_HOME/plugins/cache/personal/onedrive/${pluginVersion}`;
    if (qa?.installedBuild?.path !== expectedInstalledQaPath) {
      issues.push(`qa-report.json Pass must record the exact versioned cache path relative to CODEX_HOME: ${expectedInstalledQaPath}.`);
    }
    const offlineGateEntries = Array.isArray(qa?.offlineGates) ? qa.offlineGates : [];
    const offlineGateCounts = new Map();
    for (const gate of offlineGateEntries) {
      offlineGateCounts.set(gate?.name, (offlineGateCounts.get(gate?.name) || 0) + 1);
    }
    const actualOfflineGateNames = [...offlineGateCounts.keys()].filter((name) => typeof name === "string").sort();
    if (JSON.stringify(actualOfflineGateNames) !== JSON.stringify([...requiredQaOfflineGates].sort())) {
      issues.push(`qa-report.json Pass must record exactly the required offline gate set: ${requiredQaOfflineGates.join(", ")}.`);
    }
    for (const gateName of requiredQaOfflineGates) {
      const matches = offlineGateEntries.filter((gate) => gate?.name === gateName);
      if (matches.length !== 1) {
        issues.push(`qa-report.json Pass must record ${gateName} exactly once.`);
      } else if (matches[0].status !== "pass") {
        issues.push(`qa-report.json cannot claim Pass while offline gate ${gateName} is ${matches[0].status || "missing"}.`);
      }
    }
    if (qa?.cleanup?.remoteTestRootsRemaining !== 0
      || qa?.cleanup?.permissionsRemaining !== 0
      || qa?.cleanup?.anonymousLinksRemaining !== 0
      || qa?.cleanup?.isolatedLocalResidue !== false) {
      issues.push("qa-report.json cannot claim Pass without explicit zero-residue cleanup evidence.");
    }
    for (const component of ["bytes", "modes", "types", "symlinkTargets"]) {
      if (qa?.sourceCacheParity?.[component] !== "pass") {
        issues.push(`qa-report.json cannot claim Pass while source/cache ${component} parity is ${qa?.sourceCacheParity?.[component] || "missing"}.`);
      }
    }
  }
  return issues;
}

function checkQaReports(contentDigest) {
  const pluginManifest = readJson(join(pluginRoot, ".codex-plugin", "plugin.json"));
  const qa = readJson(join(pluginRoot, "qa-report.json"));
  if (!pluginManifest || !qa) return;
  const markdown = readFileSync(join(pluginRoot, "qa-report.md"), "utf8");
  const headCommit = currentHeadCommit();
  for (const issue of qaAlignmentProblems(qa, markdown, pluginManifest.version, headCommit, contentDigest)) fail(issue);
}

function checkMcp() {
  const mcp = readJson(join(pluginRoot, ".mcp.json"));
  const server = mcp?.mcpServers?.onedrive;
  if (!server) return fail(".mcp.json missing mcpServers.onedrive");
  if (server.command !== "node") fail(`Unexpected MCP command: ${server.command}`);
  if (server.cwd !== ".") fail(`MCP cwd should be '.': ${server.cwd}`);
  if (!server.args?.includes("./mcp/server.mjs")) fail("MCP args should include ./mcp/server.mjs");
  for (const key of ["ONEDRIVE_TENANT", "ONEDRIVE_SCOPES"]) {
    if (Object.hasOwn(server.env || {}, key)) fail(`.mcp.json must not override ${key}; use environment, config.json, or server defaults.`);
  }
}

function checkRequiredFiles() {
  for (const file of requiredFiles) {
    if (!existsSync(join(pluginRoot, file))) fail(`Missing required file: ${file}`);
  }
}

async function packageSnapshot(root) {
  const snapshot = new Map();
  for (const path of await walk(root, [])) {
    const rel = relative(root, path);
    const stat = lstatSync(path);
    const type = stat.isSymbolicLink() ? "symlink" : stat.isFile() ? "file" : "other";
    const entry = { type, mode: stat.mode & 0o777 };
    if (type === "file") entry.bytes = readFileSync(path);
    if (type === "symlink") {
      entry.target = readlinkSync(path);
      const resolvedTarget = resolve(dirname(path), entry.target);
      if (entry.target.startsWith(sep) || (resolvedTarget !== root && !resolvedTarget.startsWith(`${root}${sep}`))) {
        fail(`Packaged symlink escapes the plugin root: ${rel} -> ${entry.target}`);
      }
    }
    snapshot.set(rel, entry);
  }
  return snapshot;
}

function packagedContentDigest(snapshot) {
  const hash = createHash("sha256");
  for (const [file, entry] of [...snapshot.entries()].sort(([left], [right]) => left.localeCompare(right))) {
    if (file === "qa-report.md" || file === "qa-report.json") continue;
    hash.update(`${file}\0${entry.type}\0${entry.mode.toString(8)}\0`);
    if (entry.type === "file") {
      hash.update(`${entry.bytes.length}\0`);
      hash.update(entry.bytes);
    } else if (entry.type === "symlink") {
      hash.update(`${Buffer.byteLength(entry.target)}\0${entry.target}`);
    }
    hash.update("\0");
  }
  return hash.digest("hex");
}

async function checkInstalledDrift() {
  if (!installedRoot) return;
  if (!existsSync(installedRoot)) return fail(`Installed root does not exist: ${installedRoot}`);
  if (installedRoot === pluginRoot) return fail("--installed must point to a separate installed plugin cache, not the source root.");
  const source = await packageSnapshot(pluginRoot);
  const installed = await packageSnapshot(installedRoot);
  const sourceFiles = [...source.keys()].sort();
  const installedFiles = [...installed.keys()].sort();
  const missing = sourceFiles.filter((file) => !installed.has(file));
  const extra = installedFiles.filter((file) => !source.has(file));
  const changedBytes = [];
  const changedModes = [];
  const changedTypes = [];
  const changedSymlinkTargets = [];
  for (const file of sourceFiles) {
    if (!installed.has(file)) continue;
    const sourceEntry = source.get(file);
    const installedEntry = installed.get(file);
    if (sourceEntry.type !== installedEntry.type) changedTypes.push(file);
    if (sourceEntry.mode !== installedEntry.mode) changedModes.push({ file, source: sourceEntry.mode.toString(8), installed: installedEntry.mode.toString(8) });
    if (sourceEntry.type === "file" && installedEntry.type === "file" && !sourceEntry.bytes.equals(installedEntry.bytes)) changedBytes.push(file);
    if (sourceEntry.type === "symlink" && installedEntry.type === "symlink" && sourceEntry.target !== installedEntry.target) changedSymlinkTargets.push(file);
  }
  if (missing.length || extra.length || changedBytes.length || changedModes.length || changedTypes.length || changedSymlinkTargets.length) {
    fail(`Installed cache differs from source:\n${JSON.stringify({ missing, extra, changedBytes, changedModes, changedTypes, changedSymlinkTargets }, null, 2)}`);
  }
}

function schemaConsistencyProblems(toolName, schema) {
  const issues = [];
  const visit = (node, path, inheritedProperties = null) => {
    if (!node || typeof node !== "object" || Array.isArray(node)) return;
    let branchProperties = inheritedProperties;
    if (node.type === "object") {
      if (node.additionalProperties !== false) {
        issues.push(`${toolName} ${path} must set additionalProperties: false.`);
      }
      const properties = node.properties;
      if (!properties || typeof properties !== "object" || Array.isArray(properties)) {
        issues.push(`${toolName} ${path}.properties must be an object.`);
        branchProperties = {};
      } else {
        branchProperties = properties;
      }
      for (const required of node.required || []) {
        if (!Object.hasOwn(branchProperties, required)) {
          issues.push(`${toolName} ${path} requires undeclared property: ${required}`);
        }
      }
      for (const [property, propertySchema] of Object.entries(branchProperties)) {
        visit(propertySchema, `${path}.properties.${property}`, null);
      }
    } else if (node.required) {
      for (const required of node.required) {
        if (!inheritedProperties || !Object.hasOwn(inheritedProperties, required)) {
          issues.push(`${toolName} ${path} requires undeclared property: ${required}`);
        }
      }
    }
    if (node.items) visit(node.items, `${path}.items`, null);
    for (const keyword of ["anyOf", "oneOf", "allOf"]) {
      for (const [index, branch] of (node[keyword] || []).entries()) {
        visit(branch, `${path}.${keyword}[${index}]`, branchProperties);
      }
    }
  };
  visit(schema, "inputSchema");
  return issues;
}

function operationKindsFromSchema(schema = {}) {
  return (schema.items?.anyOf || schema.items?.oneOf || [])
    .map((branch) => branch?.properties?.type?.const)
    .filter((kind) => typeof kind === "string");
}

function compareExactValues(actual, expected) {
  const actualCounts = new Map();
  for (const value of actual) actualCounts.set(value, (actualCounts.get(value) || 0) + 1);
  const expectedSet = new Set(expected);
  return {
    missing: expected.filter((value) => !actualCounts.has(value)),
    extra: [...actualCounts.keys()].filter((value) => !expectedSet.has(value)),
    duplicates: [...actualCounts.entries()].filter(([, count]) => count > 1).map(([value]) => value)
  };
}

function exactValuesMatch(actual, expected) {
  const comparison = compareExactValues(actual, expected);
  return actual.length === expected.length
    && comparison.missing.length === 0
    && comparison.extra.length === 0
    && comparison.duplicates.length === 0;
}

function checkToolSchemas() {
  const serverPath = join(pluginRoot, "mcp", "server.mjs");
  const input = [
    JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "prepackage-check", version: "1" } } }),
    JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} }),
    ""
  ].join("\n");
  const result = spawnSync(process.execPath, [serverPath], {
    cwd: pluginRoot,
    input,
    encoding: "utf8",
    env: { ...process.env, ONEDRIVE_TEST_ACCESS_TOKEN: "prepackage-schema-check" },
    timeout: 10_000
  });
  if (result.error) return fail(`Could not inspect MCP tool schemas: ${result.error.message}`);
  if (result.status !== 0 && result.signal !== "SIGTERM") {
    return fail(`MCP schema inspection failed: ${result.stderr || result.stdout}`);
  }
  const lines = (result.stdout || "").trim().split("\n").filter(Boolean);
  const listMessage = lines.map((line) => {
    try {
      return JSON.parse(line);
    } catch {
      return null;
    }
  }).find((message) => message?.id === 2);
  const tools = listMessage?.result?.tools;
  if (!Array.isArray(tools)) return fail("MCP tools/list did not return a tools array during schema inspection.");
  const contract = compareToolContract(tools.map((tool) => tool.name));
  if (!contract.ok) fail(`MCP tool set does not exactly match the 84-tool contract: ${JSON.stringify(contract)}`);
  const seenToolNames = new Set();
  for (const tool of tools) {
    if (seenToolNames.has(tool.name)) fail(`Duplicate MCP tool registered: ${tool.name}`);
    seenToolNames.add(tool.name);
    if (!tool.description || typeof tool.description !== "string") fail(`${tool.name} must have a non-empty description.`);
    if (!tool.title || typeof tool.title !== "string") fail(`${tool.name} must have a non-empty title.`);
    for (const hint of ["readOnlyHint", "openWorldHint", "destructiveHint"]) {
      if (typeof tool.annotations?.[hint] !== "boolean") fail(`${tool.name} must define boolean annotations.${hint}.`);
    }
    const expectedSecuritySchemes = [{ type: "noauth" }];
    if (JSON.stringify(tool.securitySchemes) !== JSON.stringify(expectedSecuritySchemes)) {
      fail(`${tool.name} must advertise the standard noauth security scheme.`);
    }
    if (JSON.stringify(tool._meta?.securitySchemes) !== JSON.stringify(expectedSecuritySchemes)) {
      fail(`${tool.name} must mirror the noauth security scheme in _meta for ChatGPT compatibility.`);
    }
    if (tool.annotations.readOnlyHint && (tool.annotations.openWorldHint || tool.annotations.destructiveHint)) {
      fail(`${tool.name} is read-only but advertises write-only impact hints.`);
    }
    for (const issue of schemaConsistencyProblems(tool.name, tool.inputSchema || {})) fail(issue);
  }
  for (const [officeKind, expectedKinds] of Object.entries(expectedOfficeOperationKinds)) {
    const toolName = officeOperationToolNames[officeKind];
    const tool = tools.find((entry) => entry.name === toolName);
    const actualKinds = operationKindsFromSchema(tool?.inputSchema?.properties?.operations);
    const comparison = compareExactValues(actualKinds, expectedKinds);
    if (comparison.missing.length || comparison.extra.length || comparison.duplicates.length || actualKinds.length !== expectedKinds.length) {
      fail(`${toolName} operation schema does not exactly match the advertised ${expectedKinds.length}-operation contract: ${JSON.stringify({ actualKinds, ...comparison })}`);
    }
  }
  const officeBatch = tools.find((entry) => entry.name === "onedrive_office_batch_transform");
  const officeBatchItem = officeBatch?.inputSchema?.properties?.items?.items;
  const batchKinds = officeBatchItem?.properties?.kind?.enum || [];
  if (!exactValuesMatch(batchKinds, Object.keys(expectedOfficeOperationKinds))) {
    fail(`onedrive_office_batch_transform kind enum must be exactly word, excel, and powerpoint: ${JSON.stringify(batchKinds)}`);
  }
  const batchOperationBranches = officeBatchItem?.properties?.operations?.anyOf || [];
  const batchOperationKindSets = batchOperationBranches.map((branch) => operationKindsFromSchema(branch));
  for (const [officeKind, expectedKinds] of Object.entries(expectedOfficeOperationKinds)) {
    if (!batchOperationKindSets.some((actualKinds) => exactValuesMatch(actualKinds, expectedKinds))) {
      fail(`onedrive_office_batch_transform is missing the exact ${officeKind} ${expectedKinds.length}-operation schema.`);
    }
  }
  if (batchOperationKindSets.length !== Object.keys(expectedOfficeOperationKinds).length) {
    fail(`onedrive_office_batch_transform must expose exactly three typed operation-schema branches, got ${batchOperationKindSets.length}.`);
  }
}

if (selfCheck) {
  const exactContract = compareToolContract(ONEDRIVE_TOOL_CONTRACT);
  const missingContract = compareToolContract(ONEDRIVE_TOOL_CONTRACT.slice(1));
  const extraContract = compareToolContract([...ONEDRIVE_TOOL_CONTRACT, "onedrive_unexpected"]);
  const duplicateContract = compareToolContract([...ONEDRIVE_TOOL_CONTRACT.slice(0, -1), ONEDRIVE_TOOL_CONTRACT[0]]);
  const qaFixture = JSON.parse(readFileSync(join(pluginRoot, "qa-report.json"), "utf8"));
  const qaMarkdownFixture = readFileSync(join(pluginRoot, "qa-report.md"), "utf8");
  const headCommit = currentHeadCommit();
  const finalPassFixture = {
    ...qaFixture,
    decision: "Pass",
    source: { ...qaFixture.source, contentDigest: "a".repeat(64) },
    offlineGates: requiredQaOfflineGates.map((name) => ({ name, status: "pass" })),
    liveRuns: {
      source: { status: "pass", runId: "source-run", folderName: "source-folder", runtimeMs: 1, toolCoverage: { contract: 84, exercised: 83, blocked: 1 }, cleanupVerified: true },
      installed: { status: "pass", runId: "installed-run", folderName: "installed-folder", runtimeMs: 1, toolCoverage: { contract: 84, exercised: 83, blocked: 1 }, cleanupVerified: true }
    },
    installedBuild: { ...qaFixture.installedBuild, status: "pass", path: `$CODEX_HOME/plugins/cache/personal/onedrive/${qaFixture.source.pluginVersion}` },
    cleanup: {
      status: "pass",
      remoteTestRootsRemaining: 0,
      permissionsRemaining: 0,
      anonymousLinksRemaining: 0,
      isolatedLocalResidue: false
    },
    sourceCacheParity: { status: "pass", bytes: "pass", modes: "pass", types: "pass", symlinkTargets: "pass" }
  };
  const intentionallyIncompletePassFixture = {
    ...finalPassFixture,
    liveRuns: {
      ...finalPassFixture.liveRuns,
      source: { ...finalPassFixture.liveRuns.source, status: "pending" }
    }
  };
  const naturalLanguageIncompletePassFixture = {
    ...intentionallyIncompletePassFixture,
    decision: "Beta Passed — Restart Required"
  };
  const checks = {
    exactToolContractAccepted: exactContract.ok,
    missingToolRejected: !missingContract.ok && missingContract.missing.length === 1,
    extraToolRejected: !extraContract.ok && extraContract.extra.includes("onedrive_unexpected"),
    duplicateToolRejected: !duplicateContract.ok && duplicateContract.duplicates.length === 1,
    currentVersionAccepted: expectedPluginVersion.test("0.5.1+codex.20260719224717"),
    staleVersionRejected: !expectedPluginVersion.test("0.4.0+codex.20260713105951"),
    sensitiveFileNamesRecognized: isSensitivePackageEntryName(".env.local")
      && isSensitivePackageEntryName("signing.pem")
      && isSensitivePackageEntryName("refresh-token.txt")
      && isSensitivePackageEntryName("client-secret.json")
      && isSensitivePackageEntryName("azure-credentials.backup")
      && !isSensitivePackageEntryName("tool-contract.mjs"),
    residueDirectoriesRecognized: forbiddenResidueDirs.has("__pycache__"),
    nestedLooseObjectSchemaRejected: schemaConsistencyProblems("negative", {
      type: "object",
      properties: { nested: { type: "object", properties: {} } },
      additionalProperties: false
    }).some((issue) => issue.includes("nested") && issue.includes("additionalProperties")),
    nestedUndeclaredRequiredRejected: schemaConsistencyProblems("negative", {
      type: "object",
      properties: { nested: { type: "object", required: ["missing"], properties: {}, additionalProperties: false } },
      additionalProperties: false
    }).some((issue) => issue.includes("requires undeclared property: missing")),
    exactValueComparisonRejectsDuplicates: !exactValuesMatch(["word", "word", "excel"], ["word", "excel", "powerpoint"]),
    staleQaToolCountRejected: qaAlignmentProblems({ ...qaFixture, contract: { toolCount: 58, exact: true } }, qaMarkdownFixture, qaFixture.source.pluginVersion)
      .some((issue) => issue.includes("84-tool")),
    falsePassQaDecisionRejected: qaAlignmentProblems(intentionallyIncompletePassFixture, qaMarkdownFixture.replace(qaFixture.decision, "Pass"), qaFixture.source.pluginVersion)
      .some((issue) => issue.includes("cannot claim Pass")),
    naturalLanguageFalsePassRejected: qaAlignmentProblems(
      naturalLanguageIncompletePassFixture,
      qaMarkdownFixture.replace(qaFixture.decision, naturalLanguageIncompletePassFixture.decision),
      qaFixture.source.pluginVersion
    ).some((issue) => issue.includes("cannot claim Pass")),
    missingOfflineGateRejected: qaAlignmentProblems(
      { ...finalPassFixture, offlineGates: finalPassFixture.offlineGates.slice(1) },
      qaMarkdownFixture.replace(qaFixture.decision, "Pass"),
      qaFixture.source.pluginVersion,
      headCommit
    ).some((issue) => issue.includes("exactly the required offline gate set")),
    staleHeadCommitRejected: qaAlignmentProblems(
      { ...qaFixture, source: { ...qaFixture.source, commit: "0000000000000000000000000000000000000000" } },
      qaMarkdownFixture,
      qaFixture.source.pluginVersion,
      headCommit
    ).some((issue) => issue.includes("source commit must match HEAD")),
    incompleteLiveEvidenceRejected: qaAlignmentProblems(
      { ...finalPassFixture, liveRuns: { ...finalPassFixture.liveRuns, source: { status: "pass" } } },
      qaMarkdownFixture.replace(qaFixture.decision, "Pass"),
      qaFixture.source.pluginVersion,
      headCommit,
      "a".repeat(64)
    ).some((issue) => issue.includes("source live run ID")),
    incompleteParityEvidenceRejected: qaAlignmentProblems(
      { ...finalPassFixture, sourceCacheParity: { status: "pass" } },
      qaMarkdownFixture.replace(qaFixture.decision, "Pass"),
      qaFixture.source.pluginVersion,
      headCommit,
      "a".repeat(64)
    ).some((issue) => issue.includes("source/cache bytes parity")),
    staleContentDigestRejected: qaAlignmentProblems(
      finalPassFixture,
      qaMarkdownFixture.replace(qaFixture.decision, "Pass"),
      qaFixture.source.pluginVersion,
      headCommit,
      "b".repeat(64)
    ).some((issue) => issue.includes("packaged-content digest"))
  };
  const ok = problems.length === 0 && Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, checks, problems }, null, 2));
  process.exit(ok ? 0 : 1);
}

const files = await walk(pluginRoot);
checkRequiredFiles();
checkManifest();
checkMcp();
checkToolSchemas();
checkNoAbsoluteLocalPaths(files);
const sourceSnapshot = await packageSnapshot(pluginRoot);
checkQaReports(packagedContentDigest(sourceSnapshot));
await checkInstalledDrift();

if (problems.length) {
  console.error(JSON.stringify({ ok: false, pluginRoot, installedRoot, problems }, null, 2));
  process.exit(1);
}

console.log(JSON.stringify({
  ok: true,
  pluginRoot,
  installedRoot,
  filesChecked: files.length
}, null, 2));
