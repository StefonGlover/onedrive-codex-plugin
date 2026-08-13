#!/usr/bin/env node

import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import {
  chmodSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  assertReusableHeavyweightSubprocessLease,
  createHeavyweightSubprocessAdmissionController,
  heavyweightSubprocessAdmissionDefaults,
  heavyweightSubprocessBusyError
} from "../mcp/heavyweight-subprocess-admission.mjs";
import { renderPdfPages } from "../mcp/materialized-resources.mjs";

let checks = 0;
function check(value, message) {
  assert.ok(value, message);
  checks += 1;
}

const unit = createHeavyweightSubprocessAdmissionController();
const office = unit.acquire({ subject: "subject-a", kind: "office" });
check(office.admitted, "the first heavyweight child should be admitted");
const sameSubject = unit.acquire({ subject: "subject-a", kind: "renderer" });
const otherSubject = unit.acquire({ subject: "subject-b", kind: "renderer" });
check(!sameSubject.admitted && sameSubject.retryable, "same-subject contention should be retryable");
check(!otherSubject.admitted && otherSubject.retryable, "global contention should reject another subject");
check(!JSON.stringify(otherSubject).includes("subject-a"), "admission rejection must not leak the active subject");
const childClosed = new EventEmitter();
office.releaseOnChildCompletion(childClosed);
childClosed.emit("close", 0);
check(unit.acquire({ subject: "subject-b", kind: "renderer" }).admitted, "child close should release admission");

const reuseController = createHeavyweightSubprocessAdmissionController();
const bufferLease = reuseController.acquire({ subject: "buffer-subject", kind: "buffer" });
check(assertReusableHeavyweightSubprocessLease(bufferLease) === bufferLease, "a genuine active buffer lease should be reusable internally");
assert.throws(
  () => assertReusableHeavyweightSubprocessLease({ admitted: true, kind: "buffer", release() {} }),
  /current internal heavyweight-operation lease/i
);
bufferLease.release();
assert.throws(
  () => assertReusableHeavyweightSubprocessLease(bufferLease),
  /current internal heavyweight-operation lease/i
);
checks += 2;

const errorController = createHeavyweightSubprocessAdmissionController();
const errored = errorController.acquire({ subject: "subject-error", kind: "office" });
const childErrored = new EventEmitter();
errored.releaseOnChildCompletion(childErrored);
childErrored.emit("error", new Error("simulated spawn error"));
check(errorController.acquire({ subject: "subject-next", kind: "renderer" }).admitted, "child error should release admission");

const batchController = createHeavyweightSubprocessAdmissionController();
const batchAdmissions = Array.from({ length: 16 }, (_, index) => batchController.acquire({
  subject: `batch-subject-${index}`,
  kind: index % 2 ? "office" : "renderer"
}));
check(batchAdmissions.filter((entry) => entry.admitted).length === 1, "a 16-call batch must admit at most one heavyweight child");
batchAdmissions.find((entry) => entry.admitted)?.release();

const busyError = heavyweightSubprocessBusyError(otherSubject);
check(
  busyError.code === "HEAVYWEIGHT_SUBPROCESS_BUSY"
    && busyError.retryable === true
    && busyError.retryAfterSeconds === 1,
  "busy errors should carry bounded retry metadata"
);
check(
  heavyweightSubprocessAdmissionDefaults.maxGlobal === 1
    && heavyweightSubprocessAdmissionDefaults.maxPerSubject === 1,
  "production admission defaults should allow one heavyweight child total"
);

const testRoot = mkdtempSync(join(tmpdir(), "onedrive-heavyweight-admission-"));
try {
  const pdfPath = join(testRoot, "input.pdf");
  const failingPdfPath = join(testRoot, "failing.pdf");
  const outputRoot = join(testRoot, "render-output");
  const limiterPath = join(testRoot, "test-limiter.py");
  const rendererPath = join(testRoot, "test-renderer.py");
  const spawnLog = join(testRoot, "spawn.log");
  writeFileSync(pdfPath, "%PDF-1.4\n%%EOF\n", { mode: 0o600 });
  writeFileSync(failingPdfPath, "%PDF-1.4\nFAIL\n%%EOF\n", { mode: 0o600 });
  writeFileSync(limiterPath, `#!/usr/bin/env python3
import os, sys
separator = sys.argv.index("--")
command = sys.argv[separator + 1:]
os.execvp(command[0], command)
`, { mode: 0o700 });
  writeFileSync(rendererPath, `#!/usr/bin/env python3
import sys, time
source = sys.argv[-2]
prefix = sys.argv[-1]
with open(${JSON.stringify(spawnLog)}, "a", encoding="utf-8") as log:
    log.write("spawn\\n")
with open(source, "r", encoding="utf-8") as input_file:
    if "FAIL" in input_file.read():
        raise SystemExit(7)
time.sleep(0.2)
with open(prefix + ".png", "wb") as output:
    output.write(bytes([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1]))
`, { mode: 0o700 });
  chmodSync(limiterPath, 0o700);
  chmodSync(rendererPath, 0o700);
  const rendererSpawnCount = () => existsSync(spawnLog)
    ? readFileSync(spawnLog, "utf8").trim().split("\n").filter(Boolean).length
    : 0;

  const crossKind = createHeavyweightSubprocessAdmissionController();
  const heldOffice = crossKind.acquire({ subject: "office-subject", kind: "office" });
  await assert.rejects(
    renderPdfPages({
      pdfPath,
      pages: [1],
      outputRoot,
      pdftoppmPath: rendererPath,
      resourceLimiterPath: limiterPath,
      resourceLimiterPythonPath: process.env.ONEDRIVE_OFFICE_TEST_PYTHON || "python3",
      admissionController: crossKind,
      admissionSubject: "renderer-subject",
      maxOutputBytes: 64,
      timeoutMs: 5_000
    }),
    (error) => error.code === "HEAVYWEIGHT_SUBPROCESS_BUSY" && error.retryable === true
  );
  check(!existsSync(outputRoot), "an Office lease should reject a renderer before staging");
  heldOffice.release();

  process.env.ONEDRIVE_TEST_ACCESS_TOKEN ||= "heavyweight-admission-test-token";
  process.env.ONEDRIVE_TEST_AUTH_CONTEXT_ID ||= "heavyweight-admission-test-subject";
  process.env.ONEDRIVE_STORAGE_ROOT ||= join(testRoot, "server-storage");
  process.env.ONEDRIVE_CACHE_ROOT ||= join(testRoot, "server-cache");
  process.env.ONEDRIVE_OFFICE_PYTHON ||= process.env.ONEDRIVE_OFFICE_TEST_PYTHON || "python3";
  const { runOfficeHelper } = await import("../mcp/server.mjs");
  const heldRenderer = crossKind.acquire({ subject: "renderer-subject", kind: "renderer" });
  await assert.rejects(
    runOfficeHelper({ action: "inspect", inputPath: "/never-staged.docx", kind: "word" }, {
      admissionController: crossKind,
      admissionSubject: "office-subject"
    }),
    (error) => error.code === "HEAVYWEIGHT_SUBPROCESS_BUSY" && error.retryable === true
  );
  check(heldRenderer.admitted, "a renderer lease should block the Office helper before spawn");
  heldRenderer.release();

  const nestedController = createHeavyweightSubprocessAdmissionController();
  const outerBufferLease = nestedController.acquire({ subject: "nested-subject", kind: "buffer" });
  await assert.rejects(
    runOfficeHelper({ action: "inspect", inputPath: "/does-not-exist.docx", kind: "word" }, {
      _heavyweightAdmissionLease: outerBufferLease
    }),
    (error) => error.code !== "HEAVYWEIGHT_SUBPROCESS_BUSY"
  );
  const nestedRender = await renderPdfPages({
    pdfPath,
    pages: [1],
    outputRoot: join(testRoot, "nested-render"),
    pdftoppmPath: rendererPath,
    resourceLimiterPath: limiterPath,
    resourceLimiterPythonPath: process.env.ONEDRIVE_OFFICE_TEST_PYTHON || "python3",
    _heavyweightAdmissionLease: outerBufferLease,
    maxOutputBytes: 64,
    timeoutMs: 5_000
  });
  check(nestedRender.pages.length === 1 && rendererSpawnCount() === 1, "a renderer should reuse a genuine enclosing buffer lease without reacquiring admission");
  check(
    !nestedController.acquire({ subject: "competing-subject", kind: "renderer" }).admitted,
    "nested Office and renderer helpers must not release their enclosing buffer lease"
  );
  outerBufferLease.release();
  check(
    nestedController.acquire({ subject: "competing-subject", kind: "renderer" }).admitted,
    "the enclosing operation should release admission after nested helper completion"
  );
  await assert.rejects(
    runOfficeHelper({ action: "inspect", inputPath: "/does-not-exist.docx", kind: "word" }, {
      _heavyweightAdmissionLease: { admitted: true, kind: "buffer" }
    }),
    /current internal heavyweight-operation lease/i
  );
  checks += 1;
  await assert.rejects(
    renderPdfPages({
      pdfPath,
      pages: [1],
      outputRoot: join(testRoot, "forged-render"),
      pdftoppmPath: rendererPath,
      resourceLimiterPath: limiterPath,
      resourceLimiterPythonPath: process.env.ONEDRIVE_OFFICE_TEST_PYTHON || "python3",
      _heavyweightAdmissionLease: { admitted: true, kind: "buffer" },
      maxOutputBytes: 64,
      timeoutMs: 5_000
    }),
    /current internal heavyweight-operation lease/i
  );
  check(!existsSync(join(testRoot, "forged-render")), "a forged nested renderer lease must be rejected before staging or spawn");

  const parallel = createHeavyweightSubprocessAdmissionController();
  const renderArguments = (index) => ({
    pdfPath,
    pages: [1],
    outputRoot: join(testRoot, `parallel-${index}`),
    pdftoppmPath: rendererPath,
    resourceLimiterPath: limiterPath,
    resourceLimiterPythonPath: process.env.ONEDRIVE_OFFICE_TEST_PYTHON || "python3",
    admissionController: parallel,
    admissionSubject: `parallel-subject-${index}`,
    maxOutputBytes: 64,
    timeoutMs: 5_000
  });
  const parallelSpawnBefore = rendererSpawnCount();
  const parallelResults = await Promise.allSettled(
    Array.from({ length: 16 }, (_, index) => renderPdfPages(renderArguments(index)))
  );
  check(parallelResults.filter((entry) => entry.status === "fulfilled").length === 1, "a parallel 16-call render batch should complete only one child");
  check(
    parallelResults.filter((entry) => entry.status === "rejected"
      && entry.reason?.code === "HEAVYWEIGHT_SUBPROCESS_BUSY"
      && entry.reason?.retryable === true).length === 15,
    "the remaining parallel render calls should fail fast with retry metadata"
  );
  check(rendererSpawnCount() === parallelSpawnBefore + 1, "parallel calls should spawn exactly one renderer process");

  const sequentialSpawnBefore = rendererSpawnCount();
  await renderPdfPages(renderArguments(20));
  check(rendererSpawnCount() === sequentialSpawnBefore + 1, "sequential rendering should resume after child close");

  const failureSpawnBefore = rendererSpawnCount();
  await assert.rejects(
    renderPdfPages({ ...renderArguments(21), pdfPath: failingPdfPath }),
    /renderer failed/i
  );
  await renderPdfPages(renderArguments(22));
  check(rendererSpawnCount() === failureSpawnBefore + 2, "renderer failure should release admission for the next child");
} finally {
  rmSync(testRoot, { recursive: true, force: true });
}

console.log(JSON.stringify({ ok: true, checks }));
