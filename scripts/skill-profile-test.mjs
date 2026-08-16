#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync, readdirSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptPath = fileURLToPath(import.meta.url);
const pluginRoot = resolve(dirname(scriptPath), "..");

if (process.argv.includes("--probe")) {
  const { processMcpMessage, shutdownOneDriveServer } = await import("../mcp/server.mjs");
  await processMcpMessage({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "onedrive-skill-profile-test", version: "1" }
    }
  });
  const listed = await processMcpMessage({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
  const output = JSON.stringify({
    profile: process.env.ONEDRIVE_TOOL_PROFILE || "full",
    tools: listed.result?.tools || []
  });
  await shutdownOneDriveServer();
  await new Promise((resolveWrite, rejectWrite) => {
    process.stdout.write(`${output}\n`, (error) => error ? rejectWrite(error) : resolveWrite());
  });
  process.exit(0);
}

if (process.argv.length !== 2) {
  console.error(JSON.stringify({ ok: false, error: `Unknown arguments: ${process.argv.slice(2).join(" ")}` }, null, 2));
  process.exit(1);
}

function fail(message, details = undefined) {
  const error = new Error(message);
  error.details = details;
  throw error;
}

function parseScalar(value, lineNumber) {
  if (value === "true") return true;
  if (value === "false") return false;
  if (value === "null") return null;
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/u.test(value)) return Number(value);
  if (value.startsWith('"')) {
    try {
      return JSON.parse(value);
    } catch (error) {
      fail(`Invalid quoted YAML scalar on line ${lineNumber}: ${error.message}`);
    }
  }
  if (value.startsWith("'")) {
    if (!value.endsWith("'")) fail(`Unterminated YAML scalar on line ${lineNumber}.`);
    return value.slice(1, -1).replaceAll("''", "'");
  }
  if (/[\[\]{},&*!|>@`]/u.test(value)) {
    fail(`Unsupported YAML syntax on line ${lineNumber}; keep the profile declaration to mappings, lists, and scalars.`);
  }
  return value;
}

function parseSimpleYaml(text) {
  const rawLines = text.split(/\r?\n/u);
  const lines = rawLines
    .map((raw, index) => ({ raw, number: index + 1 }))
    .filter(({ raw }) => raw.trim() && !raw.trimStart().startsWith("#"));
  const root = {};
  const stack = [{ indent: -1, value: root }];

  for (let index = 0; index < lines.length; index += 1) {
    const { raw, number } = lines[index];
    if (raw.includes("\t")) fail(`Tabs are not allowed in the YAML declaration (line ${number}).`);
    const indent = raw.length - raw.trimStart().length;
    if (indent % 2 !== 0) fail(`YAML indentation must use two-space levels (line ${number}).`);
    const content = raw.trim();
    while (stack.length > 1 && indent <= stack.at(-1).indent) stack.pop();
    const parent = stack.at(-1).value;

    if (content.startsWith("- ")) {
      if (!Array.isArray(parent)) fail(`Unexpected YAML list item on line ${number}.`);
      parent.push(parseScalar(content.slice(2).trim(), number));
      continue;
    }

    if (Array.isArray(parent)) fail(`Unexpected YAML mapping on line ${number}.`);
    const separator = content.indexOf(":");
    if (separator < 1) fail(`Expected a YAML key/value mapping on line ${number}.`);
    const key = content.slice(0, separator).trim();
    const rawValue = content.slice(separator + 1).trim();
    if (!/^[A-Za-z0-9_][A-Za-z0-9_.-]*$/u.test(key)) fail(`Invalid YAML key on line ${number}: ${key}`);
    if (Object.hasOwn(parent, key)) fail(`Duplicate YAML key on line ${number}: ${key}`);

    if (rawValue) {
      parent[key] = parseScalar(rawValue, number);
      continue;
    }

    const next = lines[index + 1];
    const nextIndent = next ? next.raw.length - next.raw.trimStart().length : -1;
    if (!next || nextIndent <= indent) {
      parent[key] = {};
      continue;
    }
    const child = next.raw.trim().startsWith("- ") ? [] : {};
    parent[key] = child;
    stack.push({ indent, value: child });
  }

  return root;
}

function probe(profile) {
  const result = spawnSync(process.execPath, [scriptPath, "--probe"], {
    cwd: pluginRoot,
    env: {
      ...process.env,
      ONEDRIVE_TOOL_PROFILE: profile,
      ONEDRIVE_MCP_AUTH_MODE: "noauth",
      ONEDRIVE_TEST_ACCESS_TOKEN: "skill-profile-test"
    },
    encoding: "utf8",
    timeout: 15_000
  });
  if (result.error) fail(`Could not inspect the ${profile} tools/list contract: ${result.error.message}`);
  if (result.status !== 0) fail(`${profile} tools/list probe failed.`, { stdout: result.stdout, stderr: result.stderr });
  try {
    return JSON.parse(result.stdout.trim());
  } catch (error) {
    fail(`${profile} tools/list probe returned invalid JSON: ${error.message}`, result.stdout);
  }
}

function compareExactSet(actual, expected) {
  const actualCounts = new Map();
  for (const value of actual) actualCounts.set(value, (actualCounts.get(value) || 0) + 1);
  const expectedCounts = new Map();
  for (const value of expected) expectedCounts.set(value, (expectedCounts.get(value) || 0) + 1);
  return {
    missing: [...expectedCounts.keys()].filter((value) => !actualCounts.has(value)).sort(),
    extra: [...actualCounts.keys()].filter((value) => !expectedCounts.has(value)).sort(),
    duplicates: [...actualCounts.entries()].filter(([, count]) => count > 1).map(([value]) => value).sort(),
    expectedDuplicates: [...expectedCounts.entries()].filter(([, count]) => count > 1).map(([value]) => value).sort()
  };
}

function assertExactSet(actual, expected, label) {
  if (!Array.isArray(actual) || !actual.every((value) => typeof value === "string" && value)) {
    fail(`${label} must be a non-empty string array.`, actual);
  }
  const comparison = compareExactSet(actual, expected);
  if (actual.length !== expected.length || Object.values(comparison).some((values) => values.length > 0)) {
    fail(`${label} differs from the advertised schema.`, comparison);
  }
}

function operationEnum(tool) {
  const operation = tool?.inputSchema?.properties?.actions?.items?.properties?.operation;
  if (!Array.isArray(operation?.enum)) fail(`${tool?.name || "Composite tool"} must expose an actions[].operation enum.`, operation);
  return operation.enum;
}

function referencedToolNames(markdown) {
  const names = new Set();
  for (const match of markdown.matchAll(/\bonedrive_[a-z0-9_]+\b/gu)) names.add(match[0]);
  for (const code of markdown.matchAll(/`([^`\n]+)`/gu)) {
    if (/\bfetch\b/u.test(code[1])) names.add("fetch");
  }
  return [...names].sort();
}

try {
  const declarationPath = join(pluginRoot, "skills", "onedrive", "references", "tool-profiles.yaml");
  const declaration = parseSimpleYaml(readFileSync(declarationPath, "utf8"));
  if (declaration.schema_version !== 1) fail("tool-profiles.yaml schema_version must be 1.");
  const focusedDeclaration = declaration.chatgpt_focused;
  if (!focusedDeclaration || typeof focusedDeclaration !== "object") fail("tool-profiles.yaml must declare chatgpt_focused.");

  const chatgpt = probe("chatgpt");
  const full = probe("full");
  const chatgptNames = chatgpt.tools.map((tool) => tool.name);
  const fullNames = full.tools.map((tool) => tool.name);
  assertExactSet(focusedDeclaration.tools, chatgptNames, "Declared chatgpt_focused.tools");

  const compositeDeclarations = {
    onedrive_read_actions: "onedrive_read_actions_operations",
    onedrive_preview_actions: "onedrive_preview_actions_operations",
    onedrive_commit_actions: "onedrive_commit_actions_operations"
  };
  for (const [toolName, declarationKey] of Object.entries(compositeDeclarations)) {
    const tool = chatgpt.tools.find((candidate) => candidate.name === toolName);
    if (!tool) fail(`Focused tools/list is missing ${toolName}.`);
    assertExactSet(focusedDeclaration[declarationKey], operationEnum(tool), `Declared ${declarationKey}`);
  }

  const skillsRoot = join(pluginRoot, "skills");
  const skillDirectories = readdirSync(skillsRoot, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && (entry.name === "onedrive" || entry.name.startsWith("onedrive-")))
    .map((entry) => entry.name)
    .sort();
  const requiredSkills = ["onedrive", "onedrive-excel", "onedrive-powerpoint", "onedrive-review", "onedrive-word"];
  assertExactSet(skillDirectories, requiredSkills, "Focused OneDrive skill directories");

  const focusedNames = new Set(chatgptNames);
  const fullOnlyNames = new Set(fullNames.filter((name) => !focusedNames.has(name)));
  const allRuntimeNames = new Set([...chatgptNames, ...fullNames]);
  const skillReferences = {};
  for (const skillDirectory of skillDirectories) {
    const skillPath = join(skillsRoot, skillDirectory, "SKILL.md");
    const references = referencedToolNames(readFileSync(skillPath, "utf8"));
    const hidden = references.filter((name) => fullOnlyNames.has(name));
    const unknown = references.filter((name) => !allRuntimeNames.has(name));
    if (hidden.length || unknown.length) {
      fail(`Focused skill ${relative(pluginRoot, skillPath)} contains unavailable tool references.`, { hiddenFullOnly: hidden, unknown });
    }
    skillReferences[skillDirectory] = references;
  }

  const server = await import("../mcp/server.mjs");
  const initialized = await server.processMcpMessage({
    jsonrpc: "2.0",
    id: 10,
    method: "initialize",
    params: { protocolVersion: "2025-11-25", capabilities: {}, clientInfo: { name: "remote-skill-test", version: "1" } }
  });
  if (!initialized.result?.capabilities?.extensions?.["io.modelcontextprotocol/skills"]) {
    fail("MCP initialize must advertise the remote skills extension.", initialized.result?.capabilities);
  }
  const listedSkills = await server.processMcpMessage({ jsonrpc: "2.0", id: 11, method: "skills/list", params: {} });
  const remoteSkills = listedSkills.result?.skills;
  if (!Array.isArray(remoteSkills)) fail("skills/list must return a skill catalog.", listedSkills);
  assertExactSet(remoteSkills.map((skill) => skill.frontmatter?.name), requiredSkills, "Remote skills/list names");
  for (const skill of remoteSkills) {
    if (skill.uri !== `skill://onedrive/${skill.frontmatter.name}/SKILL.md`) {
      fail("Remote skill URI must use the stable OneDrive skill namespace.", skill);
    }
    if (!skill.frontmatter.description || !Array.isArray(skill.resources) || !skill.resources.length) {
      fail("Remote skills must include front matter and a complete resource manifest.", skill);
    }
    const fetched = await server.processMcpMessage({ jsonrpc: "2.0", id: 12, method: "skills/get", params: { uri: skill.uri } });
    if (JSON.stringify(fetched.result?.skill) !== JSON.stringify(skill)) fail("skills/get must return the exact catalog entry.", fetched);
    for (const resource of skill.resources) {
      const read = await server.processMcpMessage({ jsonrpc: "2.0", id: 13, method: "resources/read", params: { uri: resource.uri } });
      const content = read.result?.contents?.[0];
      if (!content || content.uri !== resource.uri) fail("resources/read must return exactly the requested skill resource.", read);
      const bytes = typeof content.text === "string" ? Buffer.from(content.text, "utf8") : Buffer.from(content.blob || "", "base64");
      const digest = `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
      if (digest !== resource.digest) fail("Remote skill resource digest does not match resources/read.", { resource, digest });
    }
  }
  await server.shutdownOneDriveServer();

  console.log(JSON.stringify({
    ok: true,
    declaration: relative(pluginRoot, declarationPath),
    profiles: { chatgpt: chatgptNames.length, full: fullNames.length },
    compositeOperations: Object.fromEntries(
      Object.entries(compositeDeclarations).map(([toolName, declarationKey]) => [toolName, focusedDeclaration[declarationKey]])
    ),
    skillReferences,
    remoteSkills: {
      count: remoteSkills.length,
      resources: remoteSkills.reduce((total, skill) => total + skill.resources.length, 0)
    }
  }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ ok: false, error: error.message, details: error.details }, null, 2));
  process.exit(1);
}
