#!/usr/bin/env node

import { randomUUID } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const officeRoot = dirname(fileURLToPath(import.meta.url));
const sourceManifestPath = join(officeRoot, "manifest.xml");
const runIdPattern = /^[A-Za-z0-9](?:[A-Za-z0-9._-]{0,62}[A-Za-z0-9])?$/;

function parseArgs(argv) {
  const values = {};
  for (const argument of argv) {
    if (argument === "--self-check") {
      if (values.selfCheck) throw new Error("--self-check may only be provided once.");
      values.selfCheck = true;
      continue;
    }
    const match = argument.match(/^--(run-id|output)=(.+)$/);
    if (!match) throw new Error(`Unknown or malformed option: ${argument}`);
    if (values[match[1]]) throw new Error(`--${match[1]} may only be provided once.`);
    values[match[1]] = match[2];
  }
  return values;
}

function generateTestManifest(source, runId, manifestId = randomUUID()) {
  if (!runIdPattern.test(runId)) throw new Error("run-id must be 1-64 safe filename characters.");
  const displayName = `Codex OneDrive Office Companion Test ${runId}`;
  const idMatches = source.match(/<Id>[^<]+<\/Id>/g) || [];
  const nameMatches = source.match(/<DisplayName DefaultValue="[^"]+"\s*\/>/g) || [];
  if (idMatches.length !== 1 || nameMatches.length !== 1) {
    throw new Error("Source Office manifest must contain exactly one Id and one DisplayName element.");
  }
  const manifest = source
    .replace(idMatches[0], `<Id>${manifestId}</Id>`)
    .replace(nameMatches[0], `<DisplayName DefaultValue="${displayName}"/>`);
  return { manifest, manifestId, displayName };
}

const args = parseArgs(process.argv.slice(2));
const source = await readFile(sourceManifestPath, "utf8");

if (args.selfCheck) {
  if (args["run-id"] || args.output) throw new Error("--self-check cannot be combined with output options.");
  const fixtureId = "11111111-2222-4333-8444-555555555555";
  const generated = generateTestManifest(source, "office-test-20260713", fixtureId);
  const ok = generated.manifest.includes(`<Id>${fixtureId}</Id>`)
    && generated.manifest.includes("Codex OneDrive Office Companion Test office-test-20260713")
    && generated.manifest.includes("<Version>1.1.1.0</Version>")
    && generated.manifest.includes("https://127.0.0.1:3443/")
    && !generated.manifest.includes("86e130ba-2570-4ec6-8533-1b17273953ce");
  console.log(JSON.stringify({ ok, manifestId: generated.manifestId, displayName: generated.displayName }));
  process.exit(ok ? 0 : 1);
}

const runId = args["run-id"];
if (!runId || !runIdPattern.test(runId)) throw new Error("--run-id=<1-64 safe filename characters> is required.");
const expectedOutput = resolve("/tmp", `codex-onedrive-office-${runId}`, `codex-onedrive-office-${runId}.xml`);
const outputPath = args.output ? resolve(args.output) : expectedOutput;
if (outputPath !== expectedOutput) {
  throw new Error(`--output must be the exact run-scoped path ${expectedOutput}`);
}
const generated = generateTestManifest(source, runId);
await mkdir(dirname(outputPath), { recursive: false }).catch((error) => {
  if (error.code !== "EEXIST") throw error;
});
await writeFile(outputPath, generated.manifest, { encoding: "utf8", flag: "wx", mode: 0o600 });
console.log(JSON.stringify({
  ok: true,
  runId,
  outputPath,
  manifestId: generated.manifestId,
  displayName: generated.displayName
}));
