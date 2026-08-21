#!/usr/bin/env node

import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const pluginRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const allowedFlags = new Set(["output", "self-check"]);

function parseArgs(argv = []) {
  const result = {};
  for (const raw of argv) {
    if (!raw.startsWith("--")) throw new Error(`Unexpected positional argument: ${raw}`);
    const [key, ...rest] = raw.slice(2).split("=");
    if (!allowedFlags.has(key)) throw new Error(`Unknown flag: --${key}`);
    if (Object.hasOwn(result, key)) throw new Error(`Duplicate flag: --${key}`);
    result[key] = rest.length ? rest.join("=") : true;
  }
  return result;
}

function parseRequirements(text) {
  const components = [];
  for (const [index, raw] of text.split(/\r?\n/u).entries()) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const match = line.match(/^([A-Za-z0-9_.-]+)==([A-Za-z0-9_.+!-]+)$/u);
    if (!match) throw new Error(`Requirement line ${index + 1} is not exactly pinned with ==.`);
    components.push({ type: "library", name: match[1], version: match[2], purl: `pkg:pypi/${match[1].toLowerCase()}@${match[2]}` });
  }
  if (!components.length) throw new Error("Pinned Python requirements are empty.");
  return components;
}

function parseDockerfile(text) {
  if (/\b(?:latest|edge)\b/iu.test(text)) throw new Error("Dockerfile must not use floating latest/edge identifiers.");
  const images = [...text.matchAll(/^FROM\s+([^\s]+)(?:\s+AS\s+([^\s]+))?$/gimu)].map((match) => ({ image: match[1], stage: match[2] || null }));
  if (images.length !== 2) throw new Error("Dockerfile must declare the reviewed builder and runtime stages exactly once.");
  for (const entry of images) {
    const pinnedByDigest = /@sha256:[0-9a-f]{64}$/iu.test(entry.image);
    const pinnedDatedDebian = /^debian:bookworm-\d{8}-slim$/u.test(entry.image);
    if (!pinnedByDigest && !pinnedDatedDebian) throw new Error(`Container base is not immutable/dated: ${entry.image}`);
  }
  const tunnelVersion = text.match(/^ARG TUNNEL_CLIENT_VERSION=([0-9]+\.[0-9]+\.[0-9]+)$/mu)?.[1];
  if (!tunnelVersion) throw new Error("Tunnel client version is not exactly pinned.");
  for (const required of ["SHA256SUMS.txt", "sha256sum -c -"]) {
    if (!text.includes(required)) throw new Error(`Tunnel client checksum verification is missing: ${required}`);
  }
  return { images, tunnelVersion };
}

function digest(value) {
  return createHash("sha256").update(value).digest("hex");
}

function buildSbom(manifest, requirements, dockerfile, parsedDocker) {
  const components = [
    { type: "application", name: manifest.name, version: manifest.version, "bom-ref": `pkg:generic/${manifest.name}@${manifest.version}` },
    ...requirements,
    ...parsedDocker.images.map((entry) => ({
      type: "container",
      name: entry.image.split(/[:@]/u)[0],
      version: entry.image.slice(entry.image.indexOf(":") + 1),
      "bom-ref": `container:${entry.image}`,
      properties: [{ name: "onedrive.build.stage", value: entry.stage || "runtime" }]
    })),
    { type: "application", name: "openai-tunnel-client", version: parsedDocker.tunnelVersion, "bom-ref": `pkg:github/openai/tunnel-client@${parsedDocker.tunnelVersion}` }
  ];
  return {
    bomFormat: "CycloneDX",
    specVersion: "1.6",
    version: 1,
    metadata: {
      component: components[0],
      properties: [
        { name: "onedrive.source.requirements.sha256", value: digest(readFileSync(resolve(pluginRoot, "scripts/requirements-office-test.txt"))) },
        { name: "onedrive.source.dockerfile.sha256", value: digest(dockerfile) }
      ]
    },
    components,
    dependencies: [{ ref: components[0]["bom-ref"], dependsOn: components.slice(1).map((component) => component["bom-ref"] || component.purl) }]
  };
}

async function selfCheck() {
  const requirements = parseRequirements("Example_Package==1.2.3\n");
  const docker = parseDockerfile("FROM debian:bookworm-20260803-slim AS tunnel-client\nARG TUNNEL_CLIENT_VERSION=0.0.10\nRUN curl SHA256SUMS.txt && sha256sum -c -\nFROM node:24-bookworm-slim@sha256:3638d9a6fe4030bd716be989438248074489337ba3275657f93595428be4fc03\n");
  const checks = {
    requirementParsed: requirements[0].purl === "pkg:pypi/example_package@1.2.3",
    datedBuilderAccepted: docker.images[0].image === "debian:bookworm-20260803-slim",
    digestRuntimeAccepted: docker.images[1].image.includes("@sha256:"),
    floatingRequirementRejected: false,
    floatingImageRejected: false
  };
  try { parseRequirements("example>=1"); } catch { checks.floatingRequirementRejected = true; }
  try { parseDockerfile("FROM debian:latest AS tunnel-client\nARG TUNNEL_CLIENT_VERSION=0.0.10\nRUN curl SHA256SUMS.txt && sha256sum -c -\nFROM node:latest\n"); } catch { checks.floatingImageRejected = true; }
  const ok = Object.values(checks).every(Boolean);
  console.log(JSON.stringify({ ok, checks }, null, 2));
  return ok;
}

const args = parseArgs(process.argv.slice(2));
if (args["self-check"] !== undefined) process.exit(await selfCheck() ? 0 : 1);

const manifest = JSON.parse(readFileSync(resolve(pluginRoot, ".codex-plugin/plugin.json"), "utf8"));
const requirementsText = readFileSync(resolve(pluginRoot, "scripts/requirements-office-test.txt"), "utf8");
const dockerfile = readFileSync(resolve(pluginRoot, "deploy/synology/Dockerfile"), "utf8");
const requirements = parseRequirements(requirementsText);
const parsedDocker = parseDockerfile(dockerfile);
const sbom = buildSbom(manifest, requirements, dockerfile, parsedDocker);
const serialized = `${JSON.stringify(sbom, null, 2)}\n`;

if (args.output !== undefined) {
  if (args.output === true) throw new Error("--output requires an explicit file path.");
  const output = resolve(String(args.output));
  await writeFile(output, serialized, { encoding: "utf8", mode: 0o600, flag: "wx" });
  console.log(JSON.stringify({ ok: true, output, componentCount: sbom.components.length, sha256: digest(serialized) }, null, 2));
} else {
  console.log(JSON.stringify({ ok: true, componentCount: sbom.components.length, sha256: digest(serialized), sbom }, null, 2));
}
