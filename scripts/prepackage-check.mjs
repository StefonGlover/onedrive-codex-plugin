#!/usr/bin/env node

import { existsSync, readFileSync, statSync } from "node:fs";
import { readdir } from "node:fs/promises";
import { dirname, extname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const args = process.argv.slice(2);
const installedIndex = args.indexOf("--installed");
const installedRoot = installedIndex >= 0 ? resolve(args[installedIndex + 1] || "") : null;
const problems = [];

const textExtensions = new Set([
  ".json", ".md", ".mjs", ".js", ".zsh", ".sh", ".txt", ".example", ".yaml", ".yml"
]);
const requiredFiles = [
  ".codex-plugin/plugin.json",
  ".mcp.json",
  "mcp/server.mjs",
  "scripts/beta-test.mjs",
  "scripts/mock-graph-test.mjs",
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
      if (entry.name === ".codex") fail(`Packaged review/artifact directory found: ${path}`);
      await walk(path, files);
    } else if (entry.isFile()) {
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

function checkInstalledDrift() {
  if (!installedRoot) return;
  if (!existsSync(installedRoot)) return fail(`Installed root does not exist: ${installedRoot}`);
  const diff = spawnSync("diff", ["-qr", pluginRoot, installedRoot], { encoding: "utf8" });
  const output = `${diff.stdout || ""}${diff.stderr || ""}`.trim();
  if (diff.status !== 0) fail(`Installed cache differs from source:\n${output}`);
}

const files = await walk(pluginRoot);
checkRequiredFiles();
checkManifest();
checkMcp();
checkNoAbsoluteLocalPaths(files);
checkInstalledDrift();

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
