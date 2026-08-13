import { randomUUID } from "node:crypto";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import {
  assertReusableHeavyweightSubprocessLease,
  heavyweightSubprocessAdmission,
  heavyweightSubprocessBusyError
} from "./heavyweight-subprocess-admission.mjs";
import {
  chmodSync,
  closeSync,
  constants,
  fstatSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  openSync,
  readFileSync,
  readSync,
  readdirSync,
  realpathSync,
  rmSync,
  rmdirSync,
  statSync
} from "node:fs";
import { isAbsolute, join, relative, resolve, sep } from "node:path";

const resourceUriPattern = /^onedrive-resource:\/\/([0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})$/;
const mimeTypePattern = /^[a-z0-9][a-z0-9!#$&^_.+-]{0,126}\/[a-z0-9][a-z0-9!#$&^_.+-]{0,126}$/i;
const defaultTtlMs = 10 * 60 * 1000;
const defaultMaxEntriesPerScope = 16;
const defaultMaxEntries = 32;
const maximumConfiguredEntries = 128;
const defaultMaxReadBytes = 25 * 1024 * 1024;
const maximumConfiguredReadBytes = 25 * 1024 * 1024;
const defaultMaxTotalBytes = 128 * 1024 * 1024;
const maximumConfiguredTotalBytes = 512 * 1024 * 1024;
const minimumPruneIntervalMs = 5;
const maximumPruneIntervalMs = 60 * 60 * 1000;
const maximumRenderPages = 8;
const minimumRenderDpi = 72;
const maximumRenderDpi = 200;
const maximumRenderDimensionPixels = 4096;
const maximumRendererCpuSeconds = 20;
const rendererAddressSpaceBytes = 768 * 1024 * 1024;
const rendererFileSizeOverheadBytes = 4096;
const defaultRenderOutputBytes = 24 * 1024 * 1024;
const maximumRenderOutputBytes = 25 * 1024 * 1024;
const maximumPdfInputBytes = 25 * 1024 * 1024;
const rendererStderrLimit = 32 * 1024;
const defaultResourceLimiterPath = fileURLToPath(new URL("../scripts/pdftoppm-limited.py", import.meta.url));

function validationError(message) {
  const error = new Error(message);
  error.code = "INVALID_ARGUMENT";
  return error;
}

function unavailableError() {
  const error = new Error("The materialized resource was not found or has expired.");
  error.code = "RESOURCE_NOT_FOUND";
  return error;
}

function storageError(message) {
  const error = new Error(message);
  error.code = "RESOURCE_STORAGE_UNAVAILABLE";
  return error;
}

function validateScopeKey(value) {
  if (
    typeof value !== "string"
    || value.length < 1
    || value.length > 512
    || value !== value.trim()
    || /[\u0000-\u001f\u007f]/.test(value)
  ) {
    throw validationError("scopeKey must be a non-empty, bounded string without control characters.");
  }
  return value;
}

function validateName(value) {
  if (
    typeof value !== "string"
    || value.length < 1
    || value.length > 255
    || value !== value.trim()
    || /[\\/\u0000-\u001f\u007f]/.test(value)
  ) {
    throw validationError("name must be a plain file name without path separators or control characters.");
  }
  return value;
}

function validateMimeType(value) {
  if (typeof value !== "string" || !mimeTypePattern.test(value)) {
    throw validationError("mimeType must be a valid type/subtype media type without parameters.");
  }
  return value.toLowerCase();
}

function validatePositiveInteger(value, label, maximum) {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw validationError(`${label} must be an integer from 1 through ${maximum}.`);
  }
  return value;
}

function tokenFromUri(uri) {
  if (typeof uri !== "string") throw unavailableError();
  const match = resourceUriPattern.exec(uri);
  if (!match) throw unavailableError();
  return match[1];
}

function pathIsWithin(root, candidate) {
  const difference = relative(root, candidate);
  return difference === "" || (!difference.startsWith(`..${sep}`) && difference !== ".." && !isAbsolute(difference));
}

function ensurePrivateDirectory(path) {
  try {
    mkdirSync(path, { recursive: true, mode: 0o700 });
    const metadata = lstatSync(path);
    if (!metadata.isDirectory() || metadata.isSymbolicLink()) {
      throw validationError("The materialized-resource root must be a regular directory, not a symbolic link.");
    }
    chmodSync(path, 0o700);
    return realpathSync(path);
  } catch (error) {
    if (error?.code === "INVALID_ARGUMENT") throw error;
    throw storageError("The private materialized-resource directory could not be prepared.");
  }
}

function requireOwnedRegularFile(root, filePath) {
  if (typeof filePath !== "string" || !isAbsolute(filePath) || filePath.includes("\u0000")) {
    throw validationError("filePath must be an absolute path inside the materialized-resource root.");
  }
  let metadata;
  let canonicalPath;
  try {
    metadata = lstatSync(filePath);
    if (!metadata.isFile() || metadata.isSymbolicLink()) throw new Error("not a regular file");
    canonicalPath = realpathSync(filePath);
  } catch {
    throw validationError("filePath must identify an existing regular file inside the materialized-resource root.");
  }
  if (!pathIsWithin(root, canonicalPath) || canonicalPath === root) {
    throw validationError("filePath must identify an existing regular file inside the materialized-resource root.");
  }
  try {
    chmodSync(canonicalPath, 0o600);
    metadata = statSync(canonicalPath);
  } catch {
    throw storageError("The materialized resource could not be secured for private access.");
  }
  if (metadata.nlink !== 1) {
    throw validationError("filePath must not be a hard-linked file.");
  }
  return { canonicalPath, metadata };
}

function removeEmptyParentsBestEffort(root, filePath) {
  let current = resolve(filePath, "..");
  while (current !== root && pathIsWithin(root, current)) {
    try {
      rmdirSync(current);
    } catch {
      break;
    }
    current = resolve(current, "..");
  }
}

function sweepCrashOrphans(root) {
  try {
    for (const entry of readdirSync(root)) {
      rmSync(join(root, entry), { recursive: true, force: true });
    }
  } catch {
    throw storageError("Crash-orphaned materialized resources could not be removed safely.");
  }
}

function readRegisteredFile(root, entry, limit) {
  let descriptor;
  try {
    const current = requireOwnedRegularFile(root, entry.filePath);
    if (
      current.metadata.dev !== entry.dev
      || current.metadata.ino !== entry.ino
      || current.metadata.size !== entry.registeredSize
      || current.metadata.mtimeMs !== entry.mtimeMs
    ) throw unavailableError();
    if (current.metadata.size > limit) {
      const error = new Error(`The materialized resource exceeds the ${limit}-byte read limit.`);
      error.code = "RESOURCE_TOO_LARGE";
      throw error;
    }
    const flags = constants.O_RDONLY | (constants.O_NOFOLLOW || 0);
    descriptor = openSync(current.canonicalPath, flags);
    const openedMetadata = fstatSync(descriptor);
    if (
      !openedMetadata.isFile()
      || openedMetadata.nlink !== 1
      || openedMetadata.dev !== entry.dev
      || openedMetadata.ino !== entry.ino
      || openedMetadata.size !== entry.registeredSize
      || openedMetadata.mtimeMs !== entry.mtimeMs
    ) {
      throw unavailableError();
    }
    const data = readFileSync(descriptor);
    if (data.length > limit) {
      const error = new Error(`The materialized resource exceeds the ${limit}-byte read limit.`);
      error.code = "RESOURCE_TOO_LARGE";
      throw error;
    }
    return data;
  } catch (error) {
    if (error?.code === "RESOURCE_TOO_LARGE" || error?.code === "RESOURCE_NOT_FOUND") throw error;
    throw unavailableError();
  } finally {
    if (descriptor !== undefined) {
      try { closeSync(descriptor); } catch {}
    }
  }
}

/**
 * Create a synchronous, scope-bound registry for temporary files beneath a
 * dedicated root. Resource metadata never contains a local path.
 */
export function createMaterializedResourceRegistry({
  rootDir,
  ttlMs = defaultTtlMs,
  maxEntriesPerScope = defaultMaxEntriesPerScope,
  maxEntries = defaultMaxEntries,
  maxReadBytes = defaultMaxReadBytes,
  maxTotalBytes = defaultMaxTotalBytes,
  maxBytesPerScope = maxTotalBytes,
  pruneIntervalMs,
  onRemove,
  now = Date.now
} = {}) {
  if (typeof rootDir !== "string" || !isAbsolute(rootDir) || rootDir.includes("\u0000")) {
    throw validationError("rootDir must be an absolute path dedicated to materialized resources.");
  }
  validatePositiveInteger(ttlMs, "ttlMs", 24 * 60 * 60 * 1000);
  validatePositiveInteger(maxEntriesPerScope, "maxEntriesPerScope", 128);
  validatePositiveInteger(maxEntries, "maxEntries", maximumConfiguredEntries);
  validatePositiveInteger(maxReadBytes, "maxReadBytes", maximumConfiguredReadBytes);
  validatePositiveInteger(maxTotalBytes, "maxTotalBytes", maximumConfiguredTotalBytes);
  validatePositiveInteger(maxBytesPerScope, "maxBytesPerScope", maximumConfiguredTotalBytes);
  if (maxEntriesPerScope > maxEntries) {
    throw validationError("maxEntriesPerScope must not exceed maxEntries.");
  }
  if (maxBytesPerScope > maxTotalBytes) throw validationError("maxBytesPerScope must not exceed maxTotalBytes.");
  const effectivePruneIntervalMs = pruneIntervalMs === undefined
    ? Math.min(60_000, Math.max(1_000, Math.floor(ttlMs / 4)))
    : validatePositiveInteger(pruneIntervalMs, "pruneIntervalMs", maximumPruneIntervalMs);
  if (effectivePruneIntervalMs < minimumPruneIntervalMs) {
    throw validationError(`pruneIntervalMs must be at least ${minimumPruneIntervalMs}.`);
  }
  if (onRemove !== undefined && typeof onRemove !== "function") throw validationError("onRemove must be a function.");
  if (typeof now !== "function") throw validationError("now must be a function.");

  const configuredRoot = resolve(rootDir);
  let canonicalRoot = ensurePrivateDirectory(configuredRoot);
  // Resource URIs and their subject/scope bindings are intentionally
  // process-local. Nothing beneath this dedicated root can be recovered after
  // a crash or restart, so remove it before accepting new quota-accounted
  // registrations.
  sweepCrashOrphans(canonicalRoot);
  const entriesByScope = new Map();
  const entriesByToken = new Map();
  const tokenByPath = new Map();
  let totalRegisteredBytes = 0;
  let closed = false;

  function currentTime() {
    const value = now();
    if (!Number.isSafeInteger(value) || value < 0) throw new Error("The materialized-resource clock returned an invalid value.");
    return value;
  }

  function ensureRoot() {
    canonicalRoot = ensurePrivateDirectory(configuredRoot);
    return canonicalRoot;
  }

  function removeEntry(scopeKey, token, entry, { forceDelete = false } = {}) {
    const scopeEntries = entriesByScope.get(scopeKey);
    if (scopeEntries?.get(token) === entry) scopeEntries.delete(token);
    if (scopeEntries?.size === 0) entriesByScope.delete(scopeKey);
    const globalEntry = entriesByToken.get(token);
    if (globalEntry?.scopeKey === scopeKey && globalEntry.entry === entry) {
      entriesByToken.delete(token);
      totalRegisteredBytes = Math.max(0, totalRegisteredBytes - entry.registeredSize);
    }
    if (tokenByPath.get(entry.filePath) === token) tokenByPath.delete(entry.filePath);
    if (forceDelete || entry.deleteOnExpiry) {
      try { rmSync(entry.filePath, { force: true }); } catch {}
      removeEmptyParentsBestEffort(canonicalRoot, entry.filePath);
    }
    try {
      onRemove?.({
        uri: `onedrive-resource://${token}`,
        scopeKey,
        expiresAt: entry.expiresAt
      });
    } catch {}
  }

  function evictOldestGlobalEntry() {
    const oldest = entriesByToken.entries().next().value;
    if (!oldest) return false;
    const [token, registered] = oldest;
    removeEntry(registered.scopeKey, token, registered.entry, { forceDelete: true });
    return true;
  }

  function prune() {
    const timestamp = currentTime();
    let removed = 0;
    for (const [scopeKey, scopeEntries] of [...entriesByScope]) {
      for (const [token, entry] of [...scopeEntries]) {
        if (entry.expiresAt <= timestamp) {
          removeEntry(scopeKey, token, entry);
          removed += 1;
        }
      }
    }
    return removed;
  }

  function register({ scopeKey, filePath, name, mimeType, deleteOnExpiry = true } = {}) {
    if (closed) throw storageError("The materialized-resource registry is closed.");
    const validScopeKey = validateScopeKey(scopeKey);
    const validName = validateName(name);
    const validMimeType = validateMimeType(mimeType);
    if (typeof deleteOnExpiry !== "boolean") throw validationError("deleteOnExpiry must be a boolean.");
    prune();
    const root = ensureRoot();
    const { canonicalPath, metadata } = requireOwnedRegularFile(root, filePath);
    if (metadata.size > maxReadBytes) {
      const error = new Error(`The materialized resource exceeds the ${maxReadBytes}-byte registration limit.`);
      error.code = "RESOURCE_TOO_LARGE";
      throw error;
    }
    if (metadata.size > maxTotalBytes) {
      const error = new Error(`The materialized resource exceeds the ${maxTotalBytes}-byte global storage quota.`);
      error.code = "RESOURCE_TOO_LARGE";
      throw error;
    }
    if (tokenByPath.has(canonicalPath)) {
      throw validationError("A materialized file may have only one active resource registration.");
    }

    let scopeEntries = entriesByScope.get(validScopeKey);
    if (!scopeEntries) {
      scopeEntries = new Map();
      entriesByScope.set(validScopeKey, scopeEntries);
    }
    if (scopeEntries.size >= maxEntriesPerScope) {
      const oldest = scopeEntries.entries().next().value;
      removeEntry(validScopeKey, oldest[0], oldest[1], { forceDelete: true });
    }
    let scopeBytes = [...scopeEntries.values()].reduce((sum, entry) => sum + entry.registeredSize, 0);
    while (scopeEntries.size && scopeBytes + metadata.size > maxBytesPerScope) {
      const oldest = scopeEntries.entries().next().value;
      scopeBytes -= oldest[1].registeredSize;
      removeEntry(validScopeKey, oldest[0], oldest[1], { forceDelete: true });
    }
    while (entriesByToken.size >= maxEntries || totalRegisteredBytes + metadata.size > maxTotalBytes) {
      if (!evictOldestGlobalEntry()) break;
    }
    scopeEntries = entriesByScope.get(validScopeKey);
    if (!scopeEntries) {
      scopeEntries = new Map();
      entriesByScope.set(validScopeKey, scopeEntries);
    }

    const token = randomUUID();
    const createdAt = currentTime();
    const entry = {
      createdAt,
      deleteOnExpiry,
      dev: metadata.dev,
      expiresAt: createdAt + ttlMs,
      filePath: canonicalPath,
      ino: metadata.ino,
      mtimeMs: metadata.mtimeMs,
      mimeType: validMimeType,
      name: validName,
      registeredSize: metadata.size,
      sizeBytes: metadata.size
    };
    scopeEntries.set(token, entry);
    entriesByToken.set(token, { scopeKey: validScopeKey, entry });
    tokenByPath.set(canonicalPath, token);
    totalRegisteredBytes += metadata.size;
    return {
      uri: `onedrive-resource://${token}`,
      name: validName,
      mimeType: validMimeType,
      sizeBytes: metadata.size,
      expiresAt: entry.expiresAt
    };
  }

  function read({ scopeKey, uri, maxBytes = maxReadBytes } = {}) {
    if (closed) throw unavailableError();
    const validScopeKey = validateScopeKey(scopeKey);
    validatePositiveInteger(maxBytes, "maxBytes", maxReadBytes);
    const token = tokenFromUri(uri);
    const scopeEntries = entriesByScope.get(validScopeKey);
    const entry = scopeEntries?.get(token);
    if (!entry) throw unavailableError();
    if (entry.expiresAt <= currentTime()) {
      removeEntry(validScopeKey, token, entry);
      throw unavailableError();
    }
    const data = readRegisteredFile(ensureRoot(), entry, maxBytes);
    return {
      uri,
      name: entry.name,
      mimeType: entry.mimeType,
      sizeBytes: data.length,
      data
    };
  }

  function clear(options = {}) {
    if (!options || typeof options !== "object" || Array.isArray(options)) {
      throw validationError("clear options must be an object.");
    }
    const hasScope = Object.prototype.hasOwnProperty.call(options, "scopeKey");
    const requestedScope = hasScope ? validateScopeKey(options.scopeKey) : undefined;
    let removed = 0;
    if (hasScope) {
      const scopeEntries = entriesByScope.get(requestedScope);
      for (const [token, entry] of [...(scopeEntries || [])]) {
        removeEntry(requestedScope, token, entry, { forceDelete: true });
        removed += 1;
      }
      return removed;
    }
    for (const [scopeKey, scopeEntries] of [...entriesByScope]) {
      for (const [token, entry] of [...scopeEntries]) {
        removeEntry(scopeKey, token, entry, { forceDelete: true });
        removed += 1;
      }
    }
    try { rmSync(configuredRoot, { recursive: true, force: true }); } catch {}
    entriesByScope.clear();
    entriesByToken.clear();
    tokenByPath.clear();
    totalRegisteredBytes = 0;
    return removed;
  }

  const pruneTimer = setInterval(() => {
    if (closed) return;
    try { prune(); } catch {}
  }, effectivePruneIntervalMs);
  pruneTimer.unref?.();

  function close() {
    if (closed) return 0;
    closed = true;
    clearInterval(pruneTimer);
    return clear();
  }

  return Object.freeze({ register, read, prune, clear, close });
}

function validatePdfInput(pdfPath) {
  if (typeof pdfPath !== "string" || !isAbsolute(pdfPath) || pdfPath.includes("\u0000")) {
    throw validationError("pdfPath must be an absolute path to a PDF file.");
  }
  let descriptor;
  try {
    const metadata = lstatSync(pdfPath);
    if (!metadata.isFile() || metadata.isSymbolicLink() || metadata.size > maximumPdfInputBytes) {
      throw new Error("invalid PDF input");
    }
    descriptor = openSync(pdfPath, constants.O_RDONLY | (constants.O_NOFOLLOW || 0));
    const signature = Buffer.alloc(5);
    const bytesRead = readSync(descriptor, signature, 0, signature.length, 0);
    if (bytesRead !== signature.length || signature.toString("ascii") !== "%PDF-") {
      throw new Error("invalid PDF signature");
    }
  } catch {
    throw validationError("pdfPath must identify a bounded regular PDF file.");
  } finally {
    if (descriptor !== undefined) {
      try { closeSync(descriptor); } catch {}
    }
  }
}

function validateRendererCommand(value, label) {
  if (
    typeof value !== "string"
    || value.length < 1
    || value.length > 4096
    || value.includes("\u0000")
  ) {
    throw validationError(`${label} must be a non-empty command or absolute executable path.`);
  }
  return value;
}

function validateResourceLimiter(path) {
  if (typeof path !== "string" || !isAbsolute(path) || path.includes("\u0000")) {
    throw validationError("resourceLimiterPath must be an absolute regular Python file.");
  }
  try {
    const metadata = lstatSync(path);
    if (!metadata.isFile() || metadata.isSymbolicLink()) throw new Error("invalid limiter");
  } catch {
    throw validationError("resourceLimiterPath must be an absolute regular Python file.");
  }
  return path;
}

function runRenderer(command, argumentsList, timeoutMs, {
  maxFileBytes,
  resourceLimiterPath,
  resourceLimiterPythonPath
}) {
  return new Promise((resolvePromise, rejectPromise) => {
    let settled = false;
    let stderrBytes = 0;
    let timedOut = false;
    let outputOverflow = false;
    const cpuSeconds = Math.max(1, Math.min(maximumRendererCpuSeconds, Math.ceil(timeoutMs / 1000)));
    const child = spawn(resourceLimiterPythonPath, [
      resourceLimiterPath,
      "--cpu-seconds", String(cpuSeconds),
      "--address-space-bytes", String(rendererAddressSpaceBytes),
      "--file-size-bytes", String(maxFileBytes),
      "--",
      command,
      ...argumentsList
    ], {
      shell: false,
      stdio: ["ignore", "ignore", "pipe"],
      windowsHide: true
    });
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill("SIGKILL");
    }, timeoutMs);
    timer.unref?.();

    function finish(error) {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (error) rejectPromise(error);
      else resolvePromise();
    }

    child.stderr?.on("data", (chunk) => {
      stderrBytes += chunk.length;
      if (stderrBytes > rendererStderrLimit) {
        outputOverflow = true;
        child.kill("SIGKILL");
      }
    });
    child.once("error", () => finish(new Error("The PDF renderer could not be started.")));
    child.once("close", (code) => {
      if (timedOut) return finish(new Error("The PDF renderer timed out."));
      if (outputOverflow) return finish(new Error("The PDF renderer produced too much diagnostic output."));
      if (code === 70) return finish(new Error("The secure PDF renderer could not establish required process resource limits on this host."));
      if (code !== 0) return finish(new Error("The PDF renderer failed."));
      return finish();
    });
  });
}

function readRenderedPng(path, remainingBytes) {
  let descriptor;
  try {
    const metadata = lstatSync(path);
    if (!metadata.isFile() || metadata.isSymbolicLink() || metadata.size < 8 || metadata.size > remainingBytes) {
      throw new Error("invalid rendered output");
    }
    chmodSync(path, 0o600);
    descriptor = openSync(path, constants.O_RDONLY | (constants.O_NOFOLLOW || 0));
    const openedMetadata = fstatSync(descriptor);
    if (!openedMetadata.isFile() || openedMetadata.size > remainingBytes) throw new Error("oversized rendered output");
    const data = readFileSync(descriptor);
    if (data.length > remainingBytes || !data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
      throw new Error("invalid PNG output");
    }
    return data;
  } finally {
    if (descriptor !== undefined) {
      try { closeSync(descriptor); } catch {}
    }
  }
}

/** Render selected PDF pages into bounded in-memory PNGs without using a shell. */
export async function renderPdfPages({
  pdfPath,
  pages,
  dpi = 120,
  outputRoot,
  pdftoppmPath = "pdftoppm",
  resourceLimiterPath = defaultResourceLimiterPath,
  resourceLimiterPythonPath = "python3",
  admissionController = heavyweightSubprocessAdmission,
  admissionSubject,
  _heavyweightAdmissionLease,
  maxOutputBytes = defaultRenderOutputBytes,
  timeoutMs = 30_000
} = {}) {
  validatePdfInput(pdfPath);
  if (!Array.isArray(pages) || pages.length < 1 || pages.length > maximumRenderPages) {
    throw validationError(`pages must contain 1 through ${maximumRenderPages} page numbers.`);
  }
  const validPages = pages.map((page) => validatePositiveInteger(page, "Each page", 10_000));
  if (new Set(validPages).size !== validPages.length) throw validationError("pages must not contain duplicates.");
  if (!Number.isSafeInteger(dpi) || dpi < minimumRenderDpi || dpi > maximumRenderDpi) {
    throw validationError(`dpi must be an integer from ${minimumRenderDpi} through ${maximumRenderDpi}.`);
  }
  validatePositiveInteger(maxOutputBytes, "maxOutputBytes", maximumRenderOutputBytes);
  validatePositiveInteger(timeoutMs, "timeoutMs", 120_000);
  if (typeof outputRoot !== "string" || !isAbsolute(outputRoot) || outputRoot.includes("\u0000")) {
    throw validationError("outputRoot must be an absolute temporary directory path.");
  }
  validateRendererCommand(pdftoppmPath, "pdftoppmPath");
  validateRendererCommand(resourceLimiterPythonPath, "resourceLimiterPythonPath");
  validateResourceLimiter(resourceLimiterPath);
  let admission;
  let ownsAdmission = false;
  if (_heavyweightAdmissionLease !== undefined) {
    admission = assertReusableHeavyweightSubprocessLease(_heavyweightAdmissionLease);
  } else {
    if (!admissionController || typeof admissionController.acquire !== "function") {
      throw validationError("admissionController must provide heavyweight-process admission.");
    }
    admission = admissionController.acquire({ subject: admissionSubject, kind: "renderer" });
    if (!admission.admitted) throw heavyweightSubprocessBusyError(admission);
    ownsAdmission = true;
  }

  let renderDirectory;
  try {
    const privateOutputRoot = ensurePrivateDirectory(resolve(outputRoot));
    try {
      renderDirectory = mkdtempSync(join(privateOutputRoot, "render-"));
      chmodSync(renderDirectory, 0o700);
    } catch {
      throw storageError("A private PDF-rendering directory could not be prepared.");
    }
    const renderedPages = [];
    let totalBytes = 0;
    for (const page of validPages) {
      const outputPrefix = join(renderDirectory, `page-${page}`);
      await runRenderer(pdftoppmPath, [
        "-f", String(page),
        "-l", String(page),
        "-r", String(dpi),
        "-scale-to", String(maximumRenderDimensionPixels),
        "-png",
        "-singlefile",
        pdfPath,
        outputPrefix
      ], timeoutMs, {
        maxFileBytes: Math.min(
          maximumRenderOutputBytes + rendererFileSizeOverheadBytes,
          maxOutputBytes - totalBytes + rendererFileSizeOverheadBytes
        ),
        resourceLimiterPath,
        resourceLimiterPythonPath
      });
      let data;
      try {
        data = readRenderedPng(`${outputPrefix}.png`, maxOutputBytes - totalBytes);
      } catch {
        const error = new Error("The PDF renderer produced an invalid or oversized PNG.");
        error.code = "RENDER_OUTPUT_INVALID";
        throw error;
      }
      totalBytes += data.length;
      renderedPages.push({ page, mimeType: "image/png", sizeBytes: data.length, data });
    }
    return { dpi, totalBytes, pages: renderedPages };
  } finally {
    if (renderDirectory) {
      try { rmSync(renderDirectory, { recursive: true, force: true }); } catch {}
    }
    if (ownsAdmission) admission.release();
  }
}
