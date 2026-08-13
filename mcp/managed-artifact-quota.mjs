import { lstat, mkdir, open, readdir, rm, stat } from "node:fs/promises";
import { isAbsolute, join, relative, resolve, sep } from "node:path";

const defaultLockTimeoutMs = 30_000;
const defaultStaleLockMs = 5 * 60_000;

function positiveInteger(value, label, maximum = Number.MAX_SAFE_INTEGER) {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw new Error(`${label} must be an integer from 1 through ${maximum}.`);
  }
  return value;
}

function pathIsWithin(root, candidate) {
  const difference = relative(root, candidate);
  return difference === "" || (!difference.startsWith(`..${sep}`) && difference !== ".." && !isAbsolute(difference));
}

function quotaError({ maxEntries, maxBytes, entries, bytes, requestedEntries = 0, requestedBytes = 0 }) {
  const error = new Error(
    `Plugin-managed downloads, updates, and backups have reached their shared storage quota `
    + `(${entries}/${maxEntries} files, ${bytes}/${maxBytes} bytes; requested ${requestedEntries} files and ${requestedBytes} bytes). `
    + "Delete unneeded files under the plugin-managed downloads, updates, or backups directories; managed Office backups can be deleted explicitly with onedrive_office_backups."
  );
  error.code = "MANAGED_ARTIFACT_QUOTA_EXCEEDED";
  return error;
}

async function sleep(milliseconds) {
  await new Promise((resolvePromise) => setTimeout(resolvePromise, milliseconds));
}

export function createManagedArtifactQuota({
  storageRoot,
  categoryRoots,
  maxEntries,
  maxBytes,
  perScopeRoot = undefined,
  maxEntriesPerScope = undefined,
  maxBytesPerScope = undefined,
  quotaLabel = "plugin-managed artifacts",
  cleanupInstruction = "Delete unneeded plugin-managed files before retrying.",
  lockTimeoutMs = defaultLockTimeoutMs,
  staleLockMs = defaultStaleLockMs
} = {}) {
  if (typeof storageRoot !== "string" || !isAbsolute(storageRoot) || storageRoot.includes("\u0000")) {
    throw new Error("storageRoot must be an absolute plugin-managed directory.");
  }
  if (!Array.isArray(categoryRoots) || categoryRoots.length < 1 || categoryRoots.length > 16) {
    throw new Error("categoryRoots must contain one through sixteen plugin-managed roots.");
  }
  const root = resolve(storageRoot);
  const roots = categoryRoots.map((candidate) => {
    if (typeof candidate !== "string" || !isAbsolute(candidate) || candidate.includes("\u0000")) {
      throw new Error("Every managed artifact category root must be absolute.");
    }
    const resolved = resolve(candidate);
    if (!pathIsWithin(root, resolved) || resolved === root) {
      throw new Error("Managed artifact category roots must be dedicated descendants of storageRoot.");
    }
    return resolved;
  });
  if (new Set(roots).size !== roots.length) throw new Error("Managed artifact category roots must be unique.");
  positiveInteger(maxEntries, "maxEntries", 100_000);
  positiveInteger(maxBytes, "maxBytes", Number.MAX_SAFE_INTEGER);
  const scopedRoot = perScopeRoot === undefined ? null : resolve(perScopeRoot);
  if (scopedRoot !== null && !roots.includes(scopedRoot)) throw new Error("perScopeRoot must be one of the managed category roots.");
  if ((maxEntriesPerScope === undefined) !== (maxBytesPerScope === undefined) || (scopedRoot === null) !== (maxEntriesPerScope === undefined)) {
    throw new Error("perScopeRoot, maxEntriesPerScope, and maxBytesPerScope must be configured together.");
  }
  if (scopedRoot !== null) {
    positiveInteger(maxEntriesPerScope, "maxEntriesPerScope", maxEntries);
    positiveInteger(maxBytesPerScope, "maxBytesPerScope", maxBytes);
  }
  if (typeof quotaLabel !== "string" || quotaLabel.length < 1 || quotaLabel.length > 128) throw new Error("quotaLabel must be a bounded string.");
  if (typeof cleanupInstruction !== "string" || cleanupInstruction.length < 1 || cleanupInstruction.length > 512) throw new Error("cleanupInstruction must be a bounded string.");
  positiveInteger(lockTimeoutMs, "lockTimeoutMs", 5 * 60_000);
  positiveInteger(staleLockMs, "staleLockMs", 60 * 60_000);
  const lockPath = join(root, ".managed-artifacts.lock");
  const reservationStates = new WeakMap();

  function managedRootForPath(candidate) {
    if (typeof candidate !== "string" || !isAbsolute(candidate) || candidate.includes("\u0000")) return null;
    const resolved = resolve(candidate);
    return roots.find((categoryRoot) => pathIsWithin(categoryRoot, resolved) && resolved !== categoryRoot) || null;
  }

  async function scanDirectory(directory, inventory) {
    let entries;
    try {
      entries = await readdir(directory, { withFileTypes: true });
    } catch (error) {
      if (error.code === "ENOENT") return;
      throw error;
    }
    for (const entry of entries) {
      const path = join(directory, entry.name);
      if (entry.isSymbolicLink()) throw new Error("Plugin-managed artifact storage contains a symbolic link and is unavailable until it is removed.");
      if (entry.isDirectory()) {
        await scanDirectory(path, inventory);
        continue;
      }
      if (!entry.isFile()) throw new Error("Plugin-managed artifact storage contains an unsupported filesystem entry.");
      const metadata = await lstat(path);
      if (!metadata.isFile() || metadata.isSymbolicLink() || metadata.nlink !== 1) {
        throw new Error("Plugin-managed artifact storage contains an unsafe file entry.");
      }
      inventory.entries += 1;
      inventory.bytes += metadata.size;
      const scopeId = scopeIdForPath(path);
      if (scopeId) {
        const usage = inventory.scopeUsage[scopeId] || { entries: 0, bytes: 0 };
        usage.entries += 1;
        usage.bytes += metadata.size;
        inventory.scopeUsage[scopeId] = usage;
      }
    }
  }

  function scopeIdForPath(candidate) {
    if (!scopedRoot || !pathIsWithin(scopedRoot, candidate)) return null;
    const parts = relative(scopedRoot, candidate).split(sep).filter(Boolean);
    return parts.length >= 2 && /^[0-9a-f]{64}$/u.test(parts[0]) ? parts[0] : null;
  }

  async function inventory() {
    const value = { entries: 0, bytes: 0, maxEntries, maxBytes, scopeUsage: {}, overQuota: false };
    for (const categoryRoot of roots) await scanDirectory(categoryRoot, value);
    value.overQuota = value.entries > maxEntries || value.bytes > maxBytes;
    if (scopedRoot) {
      value.overQuota ||= Object.values(value.scopeUsage).some((usage) => (
        usage.entries > maxEntriesPerScope || usage.bytes > maxBytesPerScope
      ));
    }
    return value;
  }

  async function withLock(callback) {
    await mkdir(root, { recursive: true, mode: 0o700 });
    const startedAt = Date.now();
    let handle;
    while (!handle) {
      try {
        handle = await open(lockPath, "wx", 0o600);
      } catch (error) {
        if (error.code !== "EEXIST") throw error;
        try {
          const metadata = await stat(lockPath);
          if (Date.now() - metadata.mtimeMs > staleLockMs) {
            await rm(lockPath, { force: true });
            continue;
          }
        } catch (statError) {
          if (statError.code !== "ENOENT") throw statError;
          continue;
        }
        if (Date.now() - startedAt >= lockTimeoutMs) {
          const error = new Error("Timed out waiting for the plugin-managed artifact quota lock.");
          error.code = "MANAGED_ARTIFACT_QUOTA_BUSY";
          throw error;
        }
        await sleep(Math.min(100, 10 + Math.floor((Date.now() - startedAt) / 10)));
      }
    }
    try {
      return await callback();
    } finally {
      await handle.close().catch(() => null);
      await rm(lockPath, { force: true }).catch(() => null);
    }
  }

  async function reconcile() {
    return await withLock(async () => await inventory());
  }

  function assertReservation(candidate) {
    const state = candidate && typeof candidate === "object" ? reservationStates.get(candidate) : null;
    if (!state?.active) throw new Error("A current internal managed-artifact quota reservation is required.");
    return candidate;
  }

  async function withReservation({ targetPaths = [], expectedEntries = 0, expectedBytes = 0, reservation = undefined } = {}, writer) {
    if (typeof writer !== "function") throw new Error("A managed artifact writer is required.");
    if (!Array.isArray(targetPaths) || targetPaths.length < 1 || targetPaths.length > 100) {
      throw new Error("targetPaths must contain one through one hundred managed artifact paths.");
    }
    for (const targetPath of targetPaths) {
      if (!managedRootForPath(targetPath)) throw new Error("Managed artifact quota reservations may target only plugin-managed downloads, updates, or backups.");
    }
    const scopeIds = [...new Set(targetPaths.map(scopeIdForPath).filter(Boolean))];
    if (scopeIds.length > 1) throw new Error("One managed-artifact reservation may target only one account-and-drive scope.");
    if (!Number.isSafeInteger(expectedEntries) || expectedEntries < 0 || expectedEntries > 100) {
      throw new Error("expectedEntries must be an integer from zero through one hundred.");
    }
    if (!Number.isSafeInteger(expectedBytes) || expectedBytes < 0) {
      throw new Error("expectedBytes must be a non-negative safe integer.");
    }
    if (reservation !== undefined) {
      const reusable = assertReservation(reservation);
      return await writer(reusable);
    }
    return await withLock(async () => {
      const before = await inventory();
      const scopeUsage = scopeIds.length ? (before.scopeUsage[scopeIds[0]] || { entries: 0, bytes: 0 }) : null;
      if (before.overQuota
        || before.entries + expectedEntries > maxEntries
        || before.bytes + expectedBytes > maxBytes
        || (scopeUsage && (
          scopeUsage.entries + expectedEntries > maxEntriesPerScope
          || scopeUsage.bytes + expectedBytes > maxBytesPerScope
        ))) {
        const error = quotaError({ ...before, requestedEntries: expectedEntries, requestedBytes: expectedBytes });
        const fairness = scopeUsage
          ? ` Current account-and-drive usage is ${scopeUsage.entries}/${maxEntriesPerScope} files and ${scopeUsage.bytes}/${maxBytesPerScope} bytes.`
          : "";
        error.message = `${quotaLabel} storage quota exceeded.${fairness} ${error.message.replace(/^Plugin-managed downloads, updates, and backups have reached their shared storage quota /u, "Global usage ")} ${cleanupInstruction}`;
        throw error;
      }
      const state = { active: true };
      const lease = Object.freeze({ quotaLabel });
      reservationStates.set(lease, state);
      try {
        const result = await writer(lease);
        const after = await inventory();
        if (after.overQuota) {
          const error = quotaError({ ...after });
          error.message = `${quotaLabel} storage quota exceeded after write verification. ${cleanupInstruction}`;
          throw error;
        }
        return result;
      } finally {
        state.active = false;
      }
    });
  }

  async function withMaintenance(callback) {
    if (typeof callback !== "function") throw new Error("A managed artifact maintenance callback is required.");
    return await withLock(callback);
  }

  return Object.freeze({ inventory, managedRootForPath, reconcile, withMaintenance, withReservation });
}
