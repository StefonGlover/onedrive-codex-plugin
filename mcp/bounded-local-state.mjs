const kibibyte = 1024;
const mebibyte = 1024 * kibibyte;

export const LOCAL_STATE_LIMITS = Object.freeze({
  runtimeScopes: 8,
  diskScopes: 8,
  contentIndex: Object.freeze({
    entries: 1000,
    serializedBytes: 16 * mebibyte,
    entrySerializedBytes: 768 * kibibyte,
    entryTextBytes: 256 * kibibyte,
    aggregateTextBytes: 8 * mebibyte,
    entrySegmentBytes: 128 * kibibyte,
    aggregateSegmentBytes: 2 * mebibyte,
    segmentTextBytes: 16 * kibibyte,
    segmentAnchorBytes: 8 * kibibyte,
    segmentsPerEntry: 2048,
    tokenBytes: 64 * kibibyte,
    tokensPerEntry: 4096,
    itemBytes: 32 * kibibyte
  }),
  metadataCache: Object.freeze({
    entries: 20_000,
    serializedBytes: 16 * mebibyte,
    entrySerializedBytes: 32 * kibibyte,
    pathRoots: 2048,
    tombstones: 2000,
    envelopeValueBytes: 64 * kibibyte
  })
});

function jsonStringByteLength(value) {
  let bytes = 2;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x22 || code === 0x5c || code === 0x08 || code === 0x09 || code === 0x0a || code === 0x0c || code === 0x0d) {
      bytes += 2;
    } else if (code <= 0x1f || (code >= 0xd800 && code <= 0xdfff && !(code <= 0xdbff && index + 1 < value.length && value.charCodeAt(index + 1) >= 0xdc00 && value.charCodeAt(index + 1) <= 0xdfff))) {
      bytes += 6;
    } else if (code <= 0x7f) {
      bytes += 1;
    } else if (code <= 0x7ff) {
      bytes += 2;
    } else if (code >= 0xd800 && code <= 0xdbff) {
      bytes += 4;
      index += 1;
    } else {
      bytes += 3;
    }
  }
  return bytes;
}

/** Exact UTF-8 byte count of JSON-compatible data, with early budget/depth/node rejection. */
export function boundedJsonByteLength(value, maximumBytes) {
  if (!Number.isSafeInteger(maximumBytes) || maximumBytes < 0) return null;
  const ancestors = new Set();
  let nodes = 0;

  const measure = (current, remaining, depth, arrayValue = false) => {
    nodes += 1;
    if (nodes > 500_000 || depth > 64 || remaining < 0) return null;
    if (current === null) return remaining >= 4 ? 4 : null;
    if (typeof current === "string") {
      const size = jsonStringByteLength(current);
      return size <= remaining ? size : null;
    }
    if (typeof current === "boolean") {
      const size = current ? 4 : 5;
      return size <= remaining ? size : null;
    }
    if (typeof current === "number") {
      const serialized = Number.isFinite(current) ? String(current) : "null";
      return serialized.length <= remaining ? serialized.length : null;
    }
    if (typeof current === "bigint") return null;
    if (["undefined", "function", "symbol"].includes(typeof current)) return arrayValue && remaining >= 4 ? 4 : undefined;
    if (!current || typeof current !== "object" || ancestors.has(current)) return null;

    ancestors.add(current);
    let size = 2;
    if (size > remaining) {
      ancestors.delete(current);
      return null;
    }
    if (Array.isArray(current)) {
      for (let index = 0; index < current.length; index += 1) {
        if (index > 0) size += 1;
        const child = measure(current[index], remaining - size, depth + 1, true);
        if (child === null || child === undefined) {
          ancestors.delete(current);
          return null;
        }
        size += child;
        if (size > remaining) {
          ancestors.delete(current);
          return null;
        }
      }
    } else {
      let emitted = 0;
      for (const key of Object.keys(current)) {
        const child = measure(current[key], remaining - size, depth + 1, false);
        if (child === undefined) continue;
        if (child === null) {
          ancestors.delete(current);
          return null;
        }
        if (emitted > 0) size += 1;
        size += jsonStringByteLength(key) + 1 + child;
        emitted += 1;
        if (size > remaining) {
          ancestors.delete(current);
          return null;
        }
      }
    }
    ancestors.delete(current);
    return size <= remaining ? size : null;
  };

  const result = measure(value, maximumBytes, 0, false);
  return result === undefined ? null : result;
}

export function boundedJsonClone(value, maximumBytes) {
  if (boundedJsonByteLength(value, maximumBytes) === null) return null;
  try {
    const serialized = JSON.stringify(value);
    if (serialized === undefined || Buffer.byteLength(serialized, "utf8") > maximumBytes) return null;
    return JSON.parse(serialized);
  } catch {
    return null;
  }
}

export function boundedUtf8Prefix(value, maximumBytes) {
  const source = String(value ?? "");
  if (!Number.isSafeInteger(maximumBytes) || maximumBytes < 0) return { value: "", bytes: 0, truncated: source.length > 0 };
  let bytes = 0;
  let end = 0;
  for (const character of source) {
    const characterBytes = Buffer.byteLength(character, "utf8");
    if (bytes + characterBytes > maximumBytes) break;
    bytes += characterBytes;
    end += character.length;
  }
  return { value: source.slice(0, end), bytes, truncated: end < source.length };
}

export function scopedStateDirectoriesToEvict(entries = [], currentScopeId, maximumScopes = LOCAL_STATE_LIMITS.diskScopes) {
  if (!Number.isSafeInteger(maximumScopes) || maximumScopes < 1) throw new Error("maximumScopes must be a positive integer.");
  const candidates = entries
    .filter((entry) => entry?.name !== currentScopeId)
    .sort((left, right) => Number(left.modifiedAt || 0) - Number(right.modifiedAt || 0) || String(left.name).localeCompare(String(right.name)));
  return candidates.slice(0, Math.max(0, candidates.length - (maximumScopes - 1)));
}

function boundedScalarString(value, maximumBytes) {
  if (value === null || value === undefined) return null;
  const source = String(value);
  return Buffer.byteLength(source, "utf8") <= maximumBytes ? source : null;
}

function boundedContentSegments(segments, limits) {
  const retained = [];
  let bytes = 2;
  let truncated = false;
  if (!Array.isArray(segments)) return { segments: retained, bytes, truncated: Boolean(segments) };
  for (const segment of segments) {
    if (retained.length >= limits.segmentsPerEntry) {
      truncated = true;
      break;
    }
    const text = boundedUtf8Prefix(segment?.text ?? "", limits.segmentTextBytes);
    const anchor = boundedJsonClone(segment?.anchor ?? null, limits.segmentAnchorBytes);
    if (!text.value || anchor === null) {
      truncated ||= Boolean(segment?.text) || segment?.anchor !== null;
      continue;
    }
    const candidate = { text: text.value, anchor };
    const candidateBytes = boundedJsonByteLength(candidate, limits.entrySegmentBytes - bytes);
    const separatorBytes = retained.length ? 1 : 0;
    if (candidateBytes === null || bytes + separatorBytes + candidateBytes > limits.entrySegmentBytes) {
      truncated = true;
      break;
    }
    retained.push(candidate);
    bytes += separatorBytes + candidateBytes;
    truncated ||= text.truncated;
  }
  truncated ||= retained.length < segments.length;
  return { segments: retained, bytes, truncated };
}

function boundedTokens(tokens, limits) {
  if (!Array.isArray(tokens)) return [];
  const retained = [];
  let bytes = 2;
  for (const token of tokens) {
    if (retained.length >= limits.tokensPerEntry) break;
    const bounded = boundedUtf8Prefix(token, 256);
    if (!bounded.value) continue;
    const tokenBytes = boundedJsonByteLength(bounded.value, limits.tokenBytes - bytes);
    const separatorBytes = retained.length ? 1 : 0;
    if (tokenBytes === null || bytes + separatorBytes + tokenBytes > limits.tokenBytes) break;
    retained.push(bounded.value);
    bytes += separatorBytes + tokenBytes;
  }
  return retained;
}

export function boundContentIndexEntry(id, entry, limits = LOCAL_STATE_LIMITS.contentIndex) {
  if (typeof id !== "string" || !id || Buffer.byteLength(id, "utf8") > 1024 || !entry || typeof entry !== "object") return null;
  const item = boundedJsonClone(entry.item, limits.itemBytes);
  if (!item || item.id !== id) return null;
  const text = boundedUtf8Prefix(entry.text ?? "", limits.entryTextBytes);
  const normalizedText = boundedUtf8Prefix(entry.normalizedText ?? text.value, limits.entryTextBytes);
  const boundedSegments = boundedContentSegments(entry.segments, limits);
  const candidate = {
    item,
    text: text.value,
    normalizedText: normalizedText.value,
    tokens: boundedTokens(entry.tokens, limits),
    indexedAt: boundedScalarString(entry.indexedAt, 128),
    source: boundedScalarString(entry.source, 256),
    bytesRead: Number.isSafeInteger(entry.bytesRead) && entry.bytesRead >= 0 ? entry.bytesRead : 0,
    textBytes: text.bytes,
    truncated: Boolean(entry.truncated || text.truncated || normalizedText.truncated || boundedSegments.truncated),
    segments: boundedSegments.segments,
    structuredKind: boundedScalarString(entry.structuredKind, 128),
    eTag: boundedScalarString(entry.eTag, 4096),
    cTag: boundedScalarString(entry.cTag, 4096),
    lastModifiedDateTime: boundedScalarString(entry.lastModifiedDateTime, 128),
    size: Number.isSafeInteger(entry.size) && entry.size >= 0 ? entry.size : 0
  };
  const serializedBytes = boundedJsonByteLength(candidate, limits.entrySerializedBytes);
  if (serializedBytes === null) return null;
  return { entry: candidate, serializedBytes, textBytes: text.bytes, segmentBytes: boundedSegments.bytes };
}

export function createContentIndexAdmissionLedger(scope, limits = LOCAL_STATE_LIMITS.contentIndex) {
  const retained = new Map();
  let entrySerializedBytes = 0;
  let textBytes = 0;
  let segmentBytes = 0;
  const serializedEntryBudget = limits.serializedBytes - mebibyte;

  const admit = (id, value) => {
    const bounded = value?.entry && Number.isSafeInteger(value.serializedBytes)
      ? value
      : boundContentIndexEntry(id, value, limits);
    if (!bounded) return { accepted: false, reason: "invalid-or-oversized-entry" };
    const previous = retained.get(id);
    const previousPropertyBytes = previous
      ? jsonStringByteLength(id) + 1 + previous.serializedBytes + (retained.size > 1 ? 1 : 0)
      : 0;
    const nextPropertyBytes = jsonStringByteLength(id) + 1 + bounded.serializedBytes + (previous ? (retained.size > 1 ? 1 : 0) : (retained.size ? 1 : 0));
    const nextCount = retained.size + (previous ? 0 : 1);
    const nextSerialized = entrySerializedBytes - previousPropertyBytes + nextPropertyBytes;
    const nextText = textBytes - (previous?.textBytes || 0) + bounded.textBytes;
    const nextSegments = segmentBytes - (previous?.segmentBytes || 0) + bounded.segmentBytes;
    if (nextCount > limits.entries) return { accepted: false, reason: "entry-count-budget" };
    if (nextSerialized > serializedEntryBudget) return { accepted: false, reason: "serialized-byte-budget" };
    if (nextText > limits.aggregateTextBytes) return { accepted: false, reason: "text-byte-budget" };
    if (nextSegments > limits.aggregateSegmentBytes) return { accepted: false, reason: "segment-byte-budget" };
    retained.set(id, bounded);
    entrySerializedBytes = nextSerialized;
    textBytes = nextText;
    segmentBytes = nextSegments;
    return { accepted: true, entry: bounded.entry };
  };

  const snapshot = () => {
    const entriesById = Object.fromEntries([...retained.entries()].map(([id, value]) => [id, value.entry]));
    const bounded = boundContentIndex({ version: 3, scope, entriesById }, { scope, preferredIds: [...retained.keys()] });
    return {
      ...bounded,
      pendingEntries: retained.size,
      pendingSerializedBytes: entrySerializedBytes,
      pendingTextBytes: textBytes,
      pendingSegmentBytes: segmentBytes
    };
  };

  return {
    admit,
    snapshot,
    get size() { return retained.size; },
    get serializedBytes() { return entrySerializedBytes; },
    get textBytes() { return textBytes; },
    get segmentBytes() { return segmentBytes; }
  };
}

function compareContentIds(entries, leftId, rightId) {
  const left = entries[leftId];
  const right = entries[rightId];
  const modified = String(right?.item?.lastModifiedDateTime || right?.lastModifiedDateTime || "")
    .localeCompare(String(left?.item?.lastModifiedDateTime || left?.lastModifiedDateTime || ""));
  if (modified) return modified;
  const indexed = String(right?.indexedAt || "").localeCompare(String(left?.indexedAt || ""));
  return indexed || leftId.localeCompare(rightId);
}

export function boundContentIndex(index = {}, { scope = index.scope ?? null, preferredIds = [] } = {}) {
  const limits = LOCAL_STATE_LIMITS.contentIndex;
  const sourceEntries = index?.entriesById && typeof index.entriesById === "object" ? index.entriesById : {};
  const preferred = [...new Set(preferredIds.filter((id) => typeof id === "string" && Object.hasOwn(sourceEntries, id)))].sort();
  const preferredSet = new Set(preferred);
  const remaining = Object.keys(sourceEntries).filter((id) => !preferredSet.has(id)).sort((a, b) => compareContentIds(sourceEntries, a, b));
  const base = {
    version: 3,
    scope,
    createdAt: boundedScalarString(index.createdAt, 128) || new Date(0).toISOString(),
    updatedAt: boundedScalarString(index.updatedAt, 128),
    itemCount: 0,
    entriesById: {}
  };
  const envelopeBytes = boundedJsonByteLength(base, limits.serializedBytes);
  if (envelopeBytes === null) throw new Error("Content-index envelope exceeds its local-state budget.");
  let entrySerializedBytes = 0;
  let textBytes = 0;
  let segmentBytes = 0;
  let evicted = 0;
  const orderedIds = [...preferred, ...remaining];
  for (let index = 0; index < orderedIds.length; index += 1) {
    const id = orderedIds[index];
    if (base.itemCount >= limits.entries) {
      evicted += orderedIds.length - index;
      break;
    }
    if (textBytes >= limits.aggregateTextBytes && String(sourceEntries[id]?.text ?? "").length) {
      evicted += orderedIds.length - index;
      break;
    }
    const bounded = boundContentIndexEntry(id, sourceEntries[id], limits);
    if (!bounded
      || textBytes + bounded.textBytes > limits.aggregateTextBytes
      || segmentBytes + bounded.segmentBytes > limits.aggregateSegmentBytes) {
      evicted += 1;
      continue;
    }
    const propertyBytes = jsonStringByteLength(id) + 1 + bounded.serializedBytes + (base.itemCount ? 1 : 0);
    // Keep a conservative envelope/encoding margin. This also ensures admission
    // never depends on a later full-object serialization to discover overflow.
    if (envelopeBytes + entrySerializedBytes + propertyBytes + mebibyte > limits.serializedBytes) {
      evicted += 1;
      continue;
    }
    base.entriesById[id] = bounded.entry;
    base.itemCount += 1;
    entrySerializedBytes += propertyBytes;
    textBytes += bounded.textBytes;
    segmentBytes += bounded.segmentBytes;
  }
  // JSON replaces the initially-empty entries object, so account for its
  // opening/closing braces exactly rather than adding properties to `{}`.
  if (base.itemCount) entrySerializedBytes -= 2;
  const exactBytes = boundedJsonByteLength(base, limits.serializedBytes);
  if (exactBytes === null) throw new Error("Content-index admission produced an oversized local-state object.");
  return { index: base, serializedBytes: exactBytes, textBytes, segmentBytes, evicted };
}

function compareMetadataIds(items, leftId, rightId) {
  const modified = String(items[rightId]?.lastModifiedDateTime || "").localeCompare(String(items[leftId]?.lastModifiedDateTime || ""));
  return modified || leftId.localeCompare(rightId);
}

function boundedEnvelopeValue(value, maximumBytes) {
  return boundedJsonClone(value, maximumBytes);
}

export function boundMetadataCache(cache = {}, { scope = cache.scope ?? null, preferredIds = [] } = {}) {
  const limits = LOCAL_STATE_LIMITS.metadataCache;
  const sourceItems = cache?.itemsById && typeof cache.itemsById === "object" ? cache.itemsById : {};
  const preferred = [...new Set(preferredIds.filter((id) => typeof id === "string" && Object.hasOwn(sourceItems, id)))].sort();
  const preferredSet = new Set(preferred);
  const remaining = Object.keys(sourceItems).filter((id) => !preferredSet.has(id)).sort((a, b) => compareMetadataIds(sourceItems, a, b));
  const pathRootsById = {};
  for (const id of Object.keys(cache.pathRootsById || {}).sort().slice(0, limits.pathRoots)) {
    const root = boundedScalarString(cache.pathRootsById[id], 4096);
    if (root !== null && Buffer.byteLength(id, "utf8") <= 1024) pathRootsById[id] = root;
  }
  const tombstones = [];
  for (const value of Array.isArray(cache.searchTombstones) ? cache.searchTombstones.slice(-limits.tombstones) : []) {
    const bounded = boundedJsonClone(value, 8192);
    if (bounded?.id) tombstones.push(bounded);
  }
  const result = {
    version: 4,
    scope,
    createdAt: boundedScalarString(cache.createdAt, 128) || new Date(0).toISOString(),
    updatedAt: boundedScalarString(cache.updatedAt, 128),
    deltaLink: boundedScalarString(cache.deltaLink, 16 * kibibyte),
    deltaNextLink: boundedScalarString(cache.deltaNextLink, 16 * kibibyte),
    deltaTarget: boundedEnvelopeValue(cache.deltaTarget, limits.envelopeValueBytes),
    scanRoot: boundedEnvelopeValue(cache.scanRoot, limits.envelopeValueBytes),
    pathRootsById,
    itemCount: 0,
    itemsById: {},
    pathsByLower: {},
    searchTombstones: tombstones
  };
  let serializedBytes = boundedJsonByteLength(result, limits.serializedBytes);
  if (serializedBytes === null) {
    result.searchTombstones = [];
    result.deltaLink = null;
    result.deltaNextLink = null;
    result.deltaTarget = null;
    result.scanRoot = null;
    result.pathRootsById = {};
    serializedBytes = boundedJsonByteLength(result, limits.serializedBytes);
  }
  if (serializedBytes === null) throw new Error("Metadata-cache envelope exceeds its local-state budget.");
  let evicted = 0;
  let pathCount = 0;
  for (const id of [...preferred, ...remaining]) {
    if (result.itemCount >= limits.entries || Buffer.byteLength(id, "utf8") > 1024) {
      evicted += 1;
      continue;
    }
    const item = boundedJsonClone(sourceItems[id], limits.entrySerializedBytes);
    if (!item || item.id !== id) {
      evicted += 1;
      continue;
    }
    const remotePath = typeof item.remotePath === "string" ? item.remotePath.toLowerCase() : null;
    const itemBytes = boundedJsonByteLength(item, limits.entrySerializedBytes);
    const itemPropertyBytes = jsonStringByteLength(id) + 1 + itemBytes + (result.itemCount ? 1 : 0);
    const addPath = Boolean(remotePath
      && Buffer.byteLength(remotePath, "utf8") <= 8192
      && !Object.hasOwn(result.pathsByLower, remotePath));
    const pathPropertyBytes = addPath
      ? jsonStringByteLength(remotePath) + 1 + jsonStringByteLength(id) + (pathCount ? 1 : 0)
      : 0;
    // Reserve a few bytes for itemCount growing from one to five digits.
    if (serializedBytes + itemPropertyBytes + pathPropertyBytes + 16 > limits.serializedBytes) {
      evicted += 1;
      continue;
    }
    result.itemsById[id] = item;
    if (addPath) {
      result.pathsByLower[remotePath] = id;
      pathCount += 1;
    }
    result.itemCount += 1;
    serializedBytes += itemPropertyBytes + pathPropertyBytes;
  }
  while (boundedJsonByteLength(result, limits.serializedBytes) === null && result.searchTombstones.length) result.searchTombstones.shift();
  if (boundedJsonByteLength(result, limits.serializedBytes) === null) {
    result.deltaLink = null;
    result.deltaNextLink = null;
    result.deltaTarget = null;
    result.scanRoot = null;
    result.pathRootsById = {};
  }
  const exactBytes = boundedJsonByteLength(result, limits.serializedBytes);
  if (exactBytes === null) throw new Error("Metadata-cache admission produced an oversized local-state object.");
  return { cache: result, serializedBytes: exactBytes, evicted };
}
