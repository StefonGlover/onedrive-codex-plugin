#!/usr/bin/env node

import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createManagedArtifactQuota } from "../mcp/managed-artifact-quota.mjs";

const root = await mkdtemp(join(tmpdir(), "onedrive-managed-artifact-quota-"));
let checks = 0;
const check = (condition, message) => { assert.ok(condition, message); checks += 1; };

try {
  const downloads = join(root, "downloads");
  const updates = join(root, "updates");
  const backups = join(root, "backups");
  const controller = () => createManagedArtifactQuota({
    storageRoot: root,
    categoryRoots: [downloads, updates, backups],
    maxEntries: 2,
    maxBytes: 10,
    lockTimeoutMs: 2_000,
    staleLockMs: 5_000
  });
  const first = controller();
  await mkdir(downloads, { recursive: true });
  await writeFile(join(downloads, "existing.bin"), Buffer.alloc(4));
  const reconciled = await first.reconcile();
  check(reconciled.entries === 1 && reconciled.bytes === 4, "startup reconciliation should count existing managed artifacts");

  const backupPath = join(backups, "new.bin");
  await mkdir(backups, { recursive: true });
  await first.withReservation({ targetPaths: [backupPath], expectedEntries: 1, expectedBytes: 6 }, async () => {
    await writeFile(backupPath, Buffer.alloc(6));
  });
  const full = await first.inventory();
  check(full.entries === 2 && full.bytes === 10, "successful reservations should remain within the shared entry and byte quotas");

  const refusedPath = join(updates, "refused.bin");
  await mkdir(updates, { recursive: true });
  await assert.rejects(
    first.withReservation({ targetPaths: [refusedPath], expectedEntries: 1, expectedBytes: 1 }, async () => {
      await writeFile(refusedPath, "x");
    }),
    (error) => error.code === "MANAGED_ARTIFACT_QUOTA_EXCEEDED" && /storage quota exceeded/iu.test(error.message)
  );
  checks += 1;

  await rm(backupPath, { force: true });
  const second = controller();
  const concurrentTargets = [join(backups, "race-a.bin"), join(backups, "race-b.bin")];
  const concurrent = await Promise.allSettled(concurrentTargets.map((targetPath, index) => {
    const instance = index === 0 ? first : second;
    return instance.withReservation({ targetPaths: [targetPath], expectedEntries: 1, expectedBytes: 6 }, async () => {
      await new Promise((resolvePromise) => setTimeout(resolvePromise, 50));
      await writeFile(targetPath, Buffer.alloc(6));
    });
  }));
  check(concurrent.filter((entry) => entry.status === "fulfilled").length === 1, "cross-controller writers should serialize and admit only one reservation");
  check(concurrent.filter((entry) => entry.status === "rejected" && entry.reason?.code === "MANAGED_ARTIFACT_QUOTA_EXCEEDED").length === 1, "the racing writer should fail closed without silent eviction");
  check((await first.inventory()).bytes === 10, "concurrent quota enforcement should preserve the hard byte ceiling");
  check(first.managedRootForPath(join(root, "outside.bin")) === null, "explicit paths outside managed roots should remain outside plugin quota accounting");

  const fairnessRoot = join(root, "fairness");
  const fairnessBackups = join(fairnessRoot, "backups");
  const scopeA = "a".repeat(64);
  const scopeB = "b".repeat(64);
  const scopeC = "c".repeat(64);
  const fairness = createManagedArtifactQuota({
    storageRoot: fairnessRoot,
    categoryRoots: [fairnessBackups],
    maxEntries: 2,
    maxBytes: 12,
    perScopeRoot: fairnessBackups,
    maxEntriesPerScope: 1,
    maxBytesPerScope: 6,
    lockTimeoutMs: 2_000,
    staleLockMs: 5_000
  });
  const scopedTarget = (scope, name) => join(fairnessBackups, scope, name);
  const scopeAFirst = scopedTarget(scopeA, "first.bin");
  await mkdir(join(fairnessBackups, scopeA), { recursive: true });
  await fairness.withReservation({ targetPaths: [scopeAFirst], expectedEntries: 1, expectedBytes: 6 }, async () => {
    await writeFile(scopeAFirst, Buffer.alloc(6));
  });
  await assert.rejects(
    fairness.withReservation({ targetPaths: [scopedTarget(scopeA, "second.bin")], expectedEntries: 1, expectedBytes: 1 }, async () => {}),
    (error) => error.code === "MANAGED_ARTIFACT_QUOTA_EXCEEDED" && /account-and-drive usage/iu.test(error.message)
  );
  checks += 1;
  const scopeBFirst = scopedTarget(scopeB, "first.bin");
  await mkdir(join(fairnessBackups, scopeB), { recursive: true });
  await fairness.withReservation({ targetPaths: [scopeBFirst], expectedEntries: 1, expectedBytes: 6 }, async () => {
    await writeFile(scopeBFirst, Buffer.alloc(6));
  });
  check((await fairness.inventory()).scopeUsage[scopeB]?.bytes === 6, "a second scope should retain its independent fair-share capacity");
  const scopeCFirst = scopedTarget(scopeC, "first.bin");
  await mkdir(join(fairnessBackups, scopeC), { recursive: true });
  await assert.rejects(
    fairness.withReservation({ targetPaths: [scopeCFirst], expectedEntries: 1, expectedBytes: 1 }, async () => {}),
    (error) => error.code === "MANAGED_ARTIFACT_QUOTA_EXCEEDED"
  );
  checks += 1;
  check((await fairness.inventory()).entries === 2, "global quota should span scopes without silently evicting retained backups");
} finally {
  await rm(root, { recursive: true, force: true });
}

console.log(JSON.stringify({ ok: true, checks }));
