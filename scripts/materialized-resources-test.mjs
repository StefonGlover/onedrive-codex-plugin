#!/usr/bin/env node

import assert from "node:assert/strict";
import {
  chmodSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  rmSync,
  symlinkSync,
  writeFileSync
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import {
  createMaterializedResourceRegistry,
  renderPdfPages
} from "../mcp/materialized-resources.mjs";

const testRoot = mkdtempSync(join(tmpdir(), "onedrive-materialized-resource-test-"));
const registryRoot = join(testRoot, "registry");
let clock = 1_000;
let checks = 0;

function check(value, message) {
  assert.ok(value, message);
  checks += 1;
}

function makeOwnedFile(name, contents) {
  mkdirSync(registryRoot, { recursive: true, mode: 0o700 });
  const path = join(registryRoot, name);
  writeFileSync(path, contents, { mode: 0o600 });
  return path;
}

function assertResourceUnavailable(callback) {
  assert.throws(callback, (error) => {
    assert.equal(error.code, "RESOURCE_NOT_FOUND");
    assert.doesNotMatch(error.message, /registry|private|tmp|Users/i);
    return true;
  });
  checks += 1;
}

try {
  const registry = createMaterializedResourceRegistry({
    rootDir: registryRoot,
    ttlMs: 50,
    maxEntriesPerScope: 2,
    maxReadBytes: 64,
    now: () => clock
  });
  check((lstatSync(registryRoot).mode & 0o777) === 0o700, "registry root should be private");

  const firstPath = makeOwnedFile("first.txt", "first resource");
  const first = registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: firstPath,
    name: "first.txt",
    mimeType: "Text/Plain"
  });
  check(/^onedrive-resource:\/\/[0-9a-f-]{36}$/.test(first.uri), "resource URI should be opaque");
  check(first.mimeType === "text/plain", "MIME type should be normalized");
  check(!JSON.stringify(first).includes(firstPath), "registration metadata must not expose a path");
  check((lstatSync(firstPath).mode & 0o777) === 0o600, "registered file should be private");

  const sameScopeRead = registry.read({ scopeKey: "subject-a:drive-a", uri: first.uri });
  assert.equal(sameScopeRead.data.toString("utf8"), "first resource");
  assert.equal(sameScopeRead.name, "first.txt");
  assert.equal(sameScopeRead.mimeType, "text/plain");
  check(!Object.hasOwn(sameScopeRead, "filePath"), "read metadata must not expose a path");
  assertResourceUnavailable(() => registry.read({ scopeKey: "subject-b:drive-a", uri: first.uri }));

  for (const invalidUri of [
    "onedrive-resource://not-a-uuid",
    `${first.uri}/extra`,
    `${first.uri}?scope=subject-a`,
    first.uri.toUpperCase(),
    "https://example.com/resource"
  ]) {
    assertResourceUnavailable(() => registry.read({ scopeKey: "subject-a:drive-a", uri: invalidUri }));
  }

  assert.throws(() => registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: firstPath,
    name: "duplicate.txt",
    mimeType: "text/plain"
  }), /only one active/i);
  checks += 1;

  const outsidePath = join(testRoot, "outside.txt");
  writeFileSync(outsidePath, "outside", { mode: 0o600 });
  assert.throws(() => registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: outsidePath,
    name: "outside.txt",
    mimeType: "text/plain"
  }), /inside the materialized-resource root/i);
  checks += 1;

  const symlinkPath = join(registryRoot, "linked.txt");
  symlinkSync(outsidePath, symlinkPath);
  assert.throws(() => registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: symlinkPath,
    name: "linked.txt",
    mimeType: "text/plain"
  }), /regular file/i);
  checks += 1;

  const boundedPath = makeOwnedFile("bounded.bin", Buffer.alloc(32, 0x42));
  const bounded = registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: boundedPath,
    name: "bounded.bin",
    mimeType: "application/octet-stream"
  });
  assert.throws(
    () => registry.read({ scopeKey: "subject-a:drive-a", uri: bounded.uri, maxBytes: 8 }),
    (error) => error.code === "RESOURCE_TOO_LARGE"
  );
  assert.throws(
    () => registry.read({ scopeKey: "subject-a:drive-a", uri: bounded.uri, maxBytes: 65 }),
    (error) => error.code === "INVALID_ARGUMENT"
  );
  checks += 2;

  const oversizedRegistrationPath = makeOwnedFile("oversized-registration.bin", Buffer.alloc(65, 0x43));
  assert.throws(
    () => registry.register({
      scopeKey: "subject-a:drive-a",
      filePath: oversizedRegistrationPath,
      name: "oversized-registration.bin",
      mimeType: "application/octet-stream"
    }),
    (error) => error.code === "RESOURCE_TOO_LARGE"
  );
  checks += 1;

  const newestPath = makeOwnedFile("newest.txt", "newest");
  const newest = registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: newestPath,
    name: "newest.txt",
    mimeType: "text/plain"
  });
  assertResourceUnavailable(() => registry.read({ scopeKey: "subject-a:drive-a", uri: first.uri }));
  check(!existsSync(firstPath), "per-scope eviction should delete the evicted owned file");
  assert.equal(registry.read({ scopeKey: "subject-a:drive-a", uri: newest.uri }).data.toString(), "newest");
  checks += 1;

  clock += 51;
  assert.equal(registry.prune(), 2);
  check(!existsSync(boundedPath) && !existsSync(newestPath), "pruning should delete expired owned files");
  assertResourceUnavailable(() => registry.read({ scopeKey: "subject-a:drive-a", uri: newest.uri }));

  const retainedPath = makeOwnedFile("retained.txt", "retained until clear");
  const retained = registry.register({
    scopeKey: "subject-a:drive-a",
    filePath: retainedPath,
    name: "retained.txt",
    mimeType: "text/plain",
    deleteOnExpiry: false
  });
  clock += 51;
  assert.equal(registry.prune(), 1);
  check(existsSync(retainedPath), "deleteOnExpiry false should retain an expired file until clear");
  assertResourceUnavailable(() => registry.read({ scopeKey: "subject-a:drive-a", uri: retained.uri }));
  assert.equal(registry.clear(), 0);
  check(!existsSync(registryRoot), "clear should remove the owned resource root even after entries expire");

  const scopeClearPath = makeOwnedFile("scope-clear.txt", "scope clear");
  registry.register({
    scopeKey: "subject-c:drive-c",
    filePath: scopeClearPath,
    name: "scope-clear.txt",
    mimeType: "text/plain",
    deleteOnExpiry: false
  });
  assert.equal(registry.clear({ scopeKey: "subject-c:drive-c" }), 1);
  check(!existsSync(scopeClearPath), "scope clear should force-delete its owned files");
  registry.close();

  const globalEntryRoot = join(testRoot, "global-entry-registry");
  const globalEntryRegistry = createMaterializedResourceRegistry({
    rootDir: globalEntryRoot,
    ttlMs: 1_000,
    maxEntriesPerScope: 2,
    maxEntries: 2,
    maxReadBytes: 64,
    maxTotalBytes: 128,
    now: () => clock
  });
  const globalFirstPath = join(globalEntryRoot, "scope-a", "first.txt");
  const globalSecondPath = join(globalEntryRoot, "scope-b", "second.txt");
  const globalThirdPath = join(globalEntryRoot, "scope-c", "third.txt");
  for (const [path, contents] of [[globalFirstPath, "first"], [globalSecondPath, "second"], [globalThirdPath, "third"]]) {
    mkdirSync(join(path, ".."), { recursive: true, mode: 0o700 });
    writeFileSync(path, contents, { mode: 0o600 });
  }
  const globalFirst = globalEntryRegistry.register({ scopeKey: "scope-a", filePath: globalFirstPath, name: "first.txt", mimeType: "text/plain", deleteOnExpiry: false });
  const globalSecond = globalEntryRegistry.register({ scopeKey: "scope-b", filePath: globalSecondPath, name: "second.txt", mimeType: "text/plain" });
  const globalThird = globalEntryRegistry.register({ scopeKey: "scope-c", filePath: globalThirdPath, name: "third.txt", mimeType: "text/plain" });
  assertResourceUnavailable(() => globalEntryRegistry.read({ scopeKey: "scope-a", uri: globalFirst.uri }));
  check(!existsSync(globalFirstPath), "global entry eviction should force-delete the oldest file even when expiry retention was requested");
  check(globalEntryRegistry.read({ scopeKey: "scope-b", uri: globalSecond.uri }).data.toString() === "second", "global entry eviction should retain the next-oldest cross-scope resource");
  check(globalEntryRegistry.read({ scopeKey: "scope-c", uri: globalThird.uri }).data.toString() === "third", "global entry eviction should retain the newest cross-scope resource");
  globalEntryRegistry.close();

  const globalByteRoot = join(testRoot, "global-byte-registry");
  const globalByteRegistry = createMaterializedResourceRegistry({
    rootDir: globalByteRoot,
    ttlMs: 1_000,
    maxEntriesPerScope: 4,
    maxEntries: 4,
    maxReadBytes: 8,
    maxTotalBytes: 10,
    now: () => clock
  });
  const byteFirstPath = join(globalByteRoot, "scope-a", "six-a.bin");
  const byteSecondPath = join(globalByteRoot, "scope-b", "six-b.bin");
  for (const path of [byteFirstPath, byteSecondPath]) {
    mkdirSync(join(path, ".."), { recursive: true, mode: 0o700 });
    writeFileSync(path, Buffer.alloc(6, 0x44), { mode: 0o600 });
  }
  const byteFirst = globalByteRegistry.register({ scopeKey: "scope-a", filePath: byteFirstPath, name: "six-a.bin", mimeType: "application/octet-stream" });
  const byteSecond = globalByteRegistry.register({ scopeKey: "scope-b", filePath: byteSecondPath, name: "six-b.bin", mimeType: "application/octet-stream" });
  assertResourceUnavailable(() => globalByteRegistry.read({ scopeKey: "scope-a", uri: byteFirst.uri }));
  check(!existsSync(byteFirstPath), "global byte quota should delete the oldest file before admitting a cross-scope resource");
  check(globalByteRegistry.read({ scopeKey: "scope-b", uri: byteSecond.uri }).sizeBytes === 6, "global byte quota should retain the newly admitted resource within quota");
  globalByteRegistry.close();

  const scopeFairnessRoot = join(testRoot, "scope-byte-fairness-registry");
  const scopeFairnessRegistry = createMaterializedResourceRegistry({
    rootDir: scopeFairnessRoot,
    ttlMs: 1_000,
    maxEntriesPerScope: 8,
    maxBytesPerScope: 8,
    maxEntries: 8,
    maxReadBytes: 8,
    maxTotalBytes: 24,
    now: () => clock
  });
  const fairnessPaths = ["a-first", "a-second", "b-first"].map((name) => join(scopeFairnessRoot, name, `${name}.bin`));
  for (const path of fairnessPaths) {
    mkdirSync(join(path, ".."), { recursive: true, mode: 0o700 });
    writeFileSync(path, Buffer.alloc(6, 0x46), { mode: 0o600 });
  }
  const scopeAFirst = scopeFairnessRegistry.register({ scopeKey: "scope-a", filePath: fairnessPaths[0], name: "a-first.bin", mimeType: "application/octet-stream" });
  const scopeASecond = scopeFairnessRegistry.register({ scopeKey: "scope-a", filePath: fairnessPaths[1], name: "a-second.bin", mimeType: "application/octet-stream" });
  assertResourceUnavailable(() => scopeFairnessRegistry.read({ scopeKey: "scope-a", uri: scopeAFirst.uri }));
  check(scopeFairnessRegistry.read({ scopeKey: "scope-a", uri: scopeASecond.uri }).sizeBytes === 6, "per-scope byte quota should retain only the newest resource in a noisy scope");
  const scopeBFirst = scopeFairnessRegistry.register({ scopeKey: "scope-b", filePath: fairnessPaths[2], name: "b-first.bin", mimeType: "application/octet-stream" });
  check(scopeFairnessRegistry.read({ scopeKey: "scope-b", uri: scopeBFirst.uri }).sizeBytes === 6, "a different scope should retain capacity after one scope reaches its byte quota");
  scopeFairnessRegistry.close();

  const orphanRoot = join(testRoot, "startup-orphan-registry");
  const orphanPath = join(orphanRoot, "abandoned-process", "orphan.bin");
  mkdirSync(join(orphanPath, ".."), { recursive: true, mode: 0o700 });
  writeFileSync(orphanPath, "unrecoverable prior-process bytes", { mode: 0o600 });
  const orphanRegistry = createMaterializedResourceRegistry({
    rootDir: orphanRoot,
    ttlMs: 1_000,
    maxReadBytes: 64,
    maxTotalBytes: 128
  });
  check(!existsSync(orphanPath), "registry startup should sweep crash-orphaned files that cannot be quota-accounted or rebound");
  const postSweepPath = join(orphanRoot, "current-process", "current.txt");
  mkdirSync(join(postSweepPath, ".."), { recursive: true, mode: 0o700 });
  writeFileSync(postSweepPath, "current", { mode: 0o600 });
  const postSweep = orphanRegistry.register({ scopeKey: "current-scope", filePath: postSweepPath, name: "current.txt", mimeType: "text/plain" });
  check(orphanRegistry.read({ scopeKey: "current-scope", uri: postSweep.uri }).data.toString() === "current", "registry should accept quota-accounted resources after startup orphan cleanup");
  orphanRegistry.close();

  const autonomousRoot = join(testRoot, "autonomous-registry");
  mkdirSync(autonomousRoot, { recursive: true, mode: 0o700 });
  const autonomousPath = join(autonomousRoot, "expires.txt");
  const removedMetadata = [];
  const autonomousRegistry = createMaterializedResourceRegistry({
    rootDir: autonomousRoot,
    ttlMs: 30,
    pruneIntervalMs: 5,
    maxReadBytes: 64,
    onRemove: (metadata) => removedMetadata.push(metadata)
  });
  writeFileSync(autonomousPath, "expires without another registry operation", { mode: 0o600 });
  const autonomous = autonomousRegistry.register({
    scopeKey: "subject-live:drive-live",
    filePath: autonomousPath,
    name: "expires.txt",
    mimeType: "text/plain"
  });
  const expiryDeadline = Date.now() + 1_000;
  while (existsSync(autonomousPath) && Date.now() < expiryDeadline) await delay(10);
  check(!existsSync(autonomousPath), "the autonomous TTL timer should delete expired files without another registry operation");
  check(
    removedMetadata.some((entry) => entry.uri === autonomous.uri
      && entry.scopeKey === "subject-live:drive-live"
      && !Object.hasOwn(entry, "filePath")),
    "autonomous removal callbacks should contain only opaque resource metadata"
  );
  assertResourceUnavailable(() => autonomousRegistry.read({ scopeKey: "subject-live:drive-live", uri: autonomous.uri }));
  autonomousRegistry.close();
  assert.throws(() => autonomousRegistry.register({
    scopeKey: "subject-live:drive-live",
    filePath: autonomousPath,
    name: "closed.txt",
    mimeType: "text/plain"
  }), /closed/i);
  checks += 1;

  const pdfPath = join(testRoot, "input.pdf");
  const largePdfPath = join(testRoot, "large.pdf");
  const failingPdfPath = join(testRoot, "failing.pdf");
  const invalidPdfPath = join(testRoot, "not-pdf.bin");
  const renderRoot = join(testRoot, "render output");
  const fakeRenderer = join(testRoot, "fake pdftoppm");
  const renderAdmission = { admissionSubject: "materialized-resource-test" };
  writeFileSync(pdfPath, "%PDF-1.4\n%%EOF\n", { mode: 0o600 });
  writeFileSync(largePdfPath, "%PDF-1.4\nBIG\n%%EOF\n", { mode: 0o600 });
  writeFileSync(failingPdfPath, "%PDF-1.4\nFAIL\n%%EOF\n", { mode: 0o600 });
  writeFileSync(invalidPdfPath, "not a PDF", { mode: 0o600 });
  writeFileSync(fakeRenderer, `#!/usr/bin/env python3
import resource
import sys
args = sys.argv[1:]
source = args[-2]
prefix = args[-1]
page = int(args[args.index("-f") + 1])
scale_index = args.index("-scale-to") if "-scale-to" in args else -1
cpu_limit = resource.getrlimit(resource.RLIMIT_CPU)
address_limit = resource.getrlimit(resource.RLIMIT_AS)
file_limit = resource.getrlimit(resource.RLIMIT_FSIZE)
if scale_index < 0 or args[scale_index + 1] != "4096":
    raise SystemExit(9)
if cpu_limit != (5, 5) or address_limit != (805306368, 805306368):
    raise SystemExit(10)
if not 4096 < file_limit[0] <= 4160 or file_limit[0] != file_limit[1]:
    raise SystemExit(11)
with open(source, "r", encoding="utf-8") as input_file:
    input_value = input_file.read()
if "FAIL" in input_value:
    raise SystemExit(7)
length = 128 if "BIG" in input_value else 16
signature = bytes([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
with open(prefix + ".png", "wb") as output_file:
    output_file.write(signature + bytes([page]) * (length - 8))
`, { mode: 0o700 });
  chmodSync(fakeRenderer, 0o700);

  if (process.platform === "darwin") {
    await assert.rejects(
      renderPdfPages({
        ...renderAdmission,
        pdfPath,
        pages: [1],
        dpi: 100,
        outputRoot: renderRoot,
        pdftoppmPath: fakeRenderer,
        maxOutputBytes: 64,
        timeoutMs: 5_000
      }),
      /resource limits/i
    );
    check(readdirSync(renderRoot).length === 0, "unsupported address-space limits should fail closed and clean outputs");
  } else {
    const rendered = await renderPdfPages({
      ...renderAdmission,
      pdfPath,
      pages: [1, 3],
      dpi: 100,
      outputRoot: renderRoot,
      pdftoppmPath: fakeRenderer,
      maxOutputBytes: 64,
      timeoutMs: 5_000
    });
    assert.equal(rendered.dpi, 100);
    assert.equal(rendered.totalBytes, 32);
    assert.deepEqual(rendered.pages.map(({ page }) => page), [1, 3]);
    check(rendered.pages.every(({ data, mimeType }) => Buffer.isBuffer(data) && mimeType === "image/png"), "rendered pages should be in-memory PNGs");
    check(rendered.pages.length === 2, "the renderer should receive pixel, CPU, address-space, and file-size limits");
    check(rendered.pages.every((page) => !Object.hasOwn(page, "filePath")), "render results must not expose local paths");
    check(readdirSync(renderRoot).length === 0, "successful rendering should clean its private output directory");
    check((lstatSync(renderRoot).mode & 0o777) === 0o700, "render output root should be private");
  }

  await assert.rejects(
    renderPdfPages({ ...renderAdmission, pdfPath, pages: [1], dpi: 100, outputRoot: renderRoot, pdftoppmPath: fakeRenderer, maxOutputBytes: 64, timeoutMs: 0 }),
    /timeoutMs/i
  );
  await assert.rejects(
    renderPdfPages({ ...renderAdmission, pdfPath, pages: [], dpi: 100, outputRoot: renderRoot, pdftoppmPath: fakeRenderer }),
    /pages/i
  );
  await assert.rejects(
    renderPdfPages({ ...renderAdmission, pdfPath, pages: [1, 1], dpi: 100, outputRoot: renderRoot, pdftoppmPath: fakeRenderer }),
    /duplicates/i
  );
  await assert.rejects(
    renderPdfPages({ ...renderAdmission, pdfPath, pages: [1], dpi: 71, outputRoot: renderRoot, pdftoppmPath: fakeRenderer }),
    /dpi/i
  );
  await assert.rejects(
    renderPdfPages({ ...renderAdmission, pdfPath, pages: [1], dpi: 201, outputRoot: renderRoot, pdftoppmPath: fakeRenderer }),
    /dpi/i
  );
  await assert.rejects(
    renderPdfPages({ ...renderAdmission, pdfPath: invalidPdfPath, pages: [1], dpi: 100, outputRoot: renderRoot, pdftoppmPath: fakeRenderer }),
    /PDF file/i
  );
  checks += 6;

  if (process.platform !== "darwin") {
    await assert.rejects(
      renderPdfPages({
        ...renderAdmission,
        pdfPath: largePdfPath,
        pages: [1],
        dpi: 100,
        outputRoot: renderRoot,
        pdftoppmPath: fakeRenderer,
        maxOutputBytes: 64
      }),
      (error) => error.code === "RENDER_OUTPUT_INVALID"
    );
    check(readdirSync(renderRoot).length === 0, "oversized render failure should clean outputs");
    await assert.rejects(
      renderPdfPages({
        ...renderAdmission,
        pdfPath: failingPdfPath,
        pages: [1],
        dpi: 100,
        outputRoot: renderRoot,
        pdftoppmPath: fakeRenderer,
        maxOutputBytes: 64
      }),
      /renderer failed/i
    );
    check(readdirSync(renderRoot).length === 0, "renderer process failure should clean outputs");
  }

  console.log(JSON.stringify({ ok: true, checks }));
} finally {
  rmSync(testRoot, { recursive: true, force: true });
}
