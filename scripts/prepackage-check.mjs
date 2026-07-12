#!/usr/bin/env node

import { existsSync, readFileSync, statSync } from "node:fs";
import { readdir } from "node:fs/promises";
import { dirname, extname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const args = process.argv.slice(2);
const problems = [];
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
for (const [index, arg] of args.entries()) {
  if (installedValueIndexes.has(index)) continue;
  if (arg === "--installed" || arg.startsWith("--installed=")) continue;
  problems.push(arg.startsWith("--") ? `Unknown option: ${arg}` : `Unexpected positional argument: ${arg}`);
}
const ignoredPackageDirs = new Set([".git", "work", "downloads", "onedrive-beta", "node_modules", "dist", "build", "coverage"]);
const ignoredPackageFiles = new Set([".DS_Store"]);
const ignoredPackageFileExtensions = new Set([".log", ".tmp", ".temp", ".bak", ".swp"]);
const ignoredPackageFileNames = new Set([".env", ".env.local", ".env.development", ".env.production"]);

const textExtensions = new Set([
  ".json", ".md", ".mjs", ".js", ".zsh", ".sh", ".txt", ".example", ".yaml", ".yml"
]);
const requiredFiles = [
  ".codex-plugin/plugin.json",
  ".mcp.json",
  "mcp/server.mjs",
  "scripts/benchmark.mjs",
  "scripts/beta-test.mjs",
  "scripts/mock-graph-test.mjs",
  "scripts/office-openxml.py",
  "scripts/office-openxml-test.py",
  "scripts/office-real-fixture-test.py",
  "scripts/office-security-test.py",
  "office-addin/manifest.xml",
  "office-addin/taskpane.html",
  "office-addin/taskpane.js",
  "office-addin/taskpane-test.mjs",
  "office-addin/README.md",
  "scripts/prepackage-check.mjs",
  "skills/onedrive/SKILL.md",
  "README.md"
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

async function walk(dir, files = []) {
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (ignoredPackageDirs.has(entry.name)) continue;
      if (entry.name === ".codex") fail(`Packaged review/artifact directory found: ${path}`);
      await walk(path, files);
    } else if (entry.isFile()) {
      if (ignoredPackageFiles.has(entry.name)) continue;
      if (ignoredPackageFileNames.has(entry.name)) continue;
      if (ignoredPackageFileExtensions.has(extname(entry.name).toLowerCase())) continue;
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
  if (!manifest.version || !manifest.version.includes("+codex.")) fail(`Version should include Codex cachebuster: ${manifest.version}`);
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

function checkMcp() {
  const mcp = readJson(join(pluginRoot, ".mcp.json"));
  const server = mcp?.mcpServers?.onedrive;
  if (!server) return fail(".mcp.json missing mcpServers.onedrive");
  if (server.command !== "node") fail(`Unexpected MCP command: ${server.command}`);
  if (server.cwd !== ".") fail(`MCP cwd should be '.': ${server.cwd}`);
  if (!server.args?.includes("./mcp/server.mjs")) fail("MCP args should include ./mcp/server.mjs");
}

function checkRequiredFiles() {
  for (const file of requiredFiles) {
    if (!existsSync(join(pluginRoot, file))) fail(`Missing required file: ${file}`);
  }
}

async function packageFileSet(root) {
  return (await walk(root, [])).map((file) => relative(root, file)).sort();
}

async function checkInstalledDrift() {
  if (!installedRoot) return;
  if (!existsSync(installedRoot)) return fail(`Installed root does not exist: ${installedRoot}`);
  if (installedRoot === pluginRoot) return fail("--installed must point to a separate installed plugin cache, not the source root.");
  const sourceFiles = await packageFileSet(pluginRoot);
  const installedFiles = await packageFileSet(installedRoot);
  const sourceSet = new Set(sourceFiles);
  const installedSet = new Set(installedFiles);
  const missing = sourceFiles.filter((file) => !installedSet.has(file));
  const extra = installedFiles.filter((file) => !sourceSet.has(file));
  const changed = [];
  for (const file of sourceFiles) {
    if (!installedSet.has(file)) continue;
    const source = readFileSync(join(pluginRoot, file));
    const installed = readFileSync(join(installedRoot, file));
    if (!source.equals(installed)) changed.push(file);
  }
  if (missing.length || extra.length || changed.length) {
    fail(`Installed cache differs from source:\n${JSON.stringify({ missing, extra, changed }, null, 2)}`);
  }
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
  const seenToolNames = new Set();
  for (const tool of tools) {
    if (seenToolNames.has(tool.name)) fail(`Duplicate MCP tool registered: ${tool.name}`);
    seenToolNames.add(tool.name);
    const schema = tool.inputSchema || {};
    const properties = schema.properties || {};
    for (const required of schema.required || []) {
      if (!Object.hasOwn(properties, required)) fail(`${tool.name} schema requires undeclared property: ${required}`);
    }
    for (const branch of schema.anyOf || []) {
      for (const required of branch.required || []) {
        if (!Object.hasOwn(properties, required)) fail(`${tool.name} anyOf requires undeclared property: ${required}`);
      }
    }
  }
}

const files = await walk(pluginRoot);
checkRequiredFiles();
checkManifest();
checkMcp();
checkToolSchemas();
checkNoAbsoluteLocalPaths(files);
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
