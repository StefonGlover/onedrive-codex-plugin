#!/usr/bin/env node

import assert from "node:assert/strict";
import { mkdtempSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  LOCAL_STATE_LIMITS,
  boundContentIndex,
  boundMetadataCache,
  boundedJsonByteLength,
  createContentIndexAdmissionLedger,
  scopedStateDirectoriesToEvict
} from "../mcp/bounded-local-state.mjs";

let checks = 0;
const check = (condition, message) => {
  assert.ok(condition, message);
  checks += 1;
};

const scope = { authContextId: "subject-a", driveId: "drive-a" };
const largeText = "A😀".repeat(350_000);
const largeTokens = Array.from({ length: 5000 }, (_, tokenIndex) => `token-${tokenIndex}`);
const entriesById = {};
for (let index = 0; index < 1000; index += 1) {
  const id = `large-${String(index).padStart(4, "0")}`;
  entriesById[id] = {
    item: { id, name: `${id}.txt`, file: { mimeType: "text/plain" }, lastModifiedDateTime: new Date(index * 1000).toISOString() },
    text: largeText,
    normalizedText: largeText,
    tokens: largeTokens,
    indexedAt: new Date(index * 1000).toISOString(),
    source: "text-read",
    segments: [{ text: largeText, anchor: { type: "paragraph", index } }]
  };
}

const bounded = boundContentIndex({
  version: 3,
  scope,
  createdAt: new Date(0).toISOString(),
  entriesById
}, { scope, preferredIds: ["large-0999"] });
check(bounded.index.itemCount <= LOCAL_STATE_LIMITS.contentIndex.entries, "content entry count exceeded its cap");
check(bounded.textBytes <= LOCAL_STATE_LIMITS.contentIndex.aggregateTextBytes, "aggregate content text exceeded its cap");
check(bounded.segmentBytes <= LOCAL_STATE_LIMITS.contentIndex.aggregateSegmentBytes, "aggregate segment content exceeded its cap");
check(bounded.serializedBytes <= LOCAL_STATE_LIMITS.contentIndex.serializedBytes, "content index exceeded its serialized cap");
check(Boolean(bounded.index.entriesById["large-0999"]), "preferred focused-warm entry was not retained");
check(bounded.index.entriesById["large-0999"].truncated === true, "bounded focused-warm text was not marked truncated");

const admission = createContentIndexAdmissionLedger(scope);
let admissionRejected = 0;
for (let index = 0; index < 1000; index += 1) {
  const id = `pending-${String(index).padStart(4, "0")}`;
  const result = admission.admit(id, {
    item: { id, name: `${id}.txt`, file: { mimeType: "text/plain" } },
    text: largeText,
    normalizedText: largeText,
    tokens: largeTokens,
    indexedAt: new Date(index * 1000).toISOString(),
    source: "text-read",
    segments: [{ text: largeText, anchor: { type: "paragraph", index } }]
  });
  if (!result.accepted) admissionRejected += 1;
  check(admission.serializedBytes <= LOCAL_STATE_LIMITS.contentIndex.serializedBytes - 1024 * 1024, "incremental pending serialized bytes exceeded budget");
  check(admission.textBytes <= LOCAL_STATE_LIMITS.contentIndex.aggregateTextBytes, "incremental pending text exceeded budget");
  check(admission.segmentBytes <= LOCAL_STATE_LIMITS.contentIndex.aggregateSegmentBytes, "incremental pending segments exceeded budget");
}
const admissionSnapshot = admission.snapshot();
check(admissionRejected > 0, "incremental admission did not reject inputs after reaching its budget");
check(admissionSnapshot.serializedBytes <= LOCAL_STATE_LIMITS.contentIndex.serializedBytes, "incremental admission snapshot exceeded disk budget");

const testRoot = mkdtempSync(join(tmpdir(), "onedrive-bounded-state-"));
try {
  const path = join(testRoot, "content-index.json");
  writeFileSync(path, JSON.stringify(bounded.index), { mode: 0o600 });
  check(statSync(path).size <= LOCAL_STATE_LIMITS.contentIndex.serializedBytes, "persisted content index exceeded its disk cap");
} finally {
  rmSync(testRoot, { recursive: true, force: true });
}

const warmed = boundContentIndex({
  version: 3,
  scope,
  createdAt: new Date(0).toISOString(),
  entriesById: {
    focused: {
      item: { id: "focused", name: "Focused.txt", file: { mimeType: "text/plain" } },
      text: "Focused fetch warming stays searchable.",
      normalizedText: "focused fetch warming stays searchable",
      tokens: ["focused", "fetch", "warming", "stays", "searchable"],
      indexedAt: new Date(0).toISOString(),
      source: "graph-text-export",
      segments: []
    }
  }
}, { scope, preferredIds: ["focused"] }).index;
check(warmed.itemCount === 1 && warmed.entriesById.focused?.tokens.includes("warming"), "focused fetch warming did not remain functional");

const malformed = boundContentIndex({
  version: 3,
  scope,
  entriesById: {
    expected: { item: { id: "foreign" }, text: "must not survive", normalizedText: "must not survive", tokens: ["survive"] }
  }
}, { scope }).index;
check(malformed.itemCount === 0, "foreign entry identity did not fail closed");

const metadataItems = {};
const metadataPadding = "m".repeat(24 * 1024);
for (let index = 0; index < 1000; index += 1) {
  const id = `metadata-${String(index).padStart(4, "0")}`;
  metadataItems[id] = { id, name: `${id}.txt`, remotePath: `Folder/${id}.txt`, lastModifiedDateTime: new Date(index * 1000).toISOString(), padding: metadataPadding };
}
const metadata = boundMetadataCache({ version: 4, scope, itemsById: metadataItems }, { scope, preferredIds: ["metadata-0999"] });
check(metadata.serializedBytes <= LOCAL_STATE_LIMITS.metadataCache.serializedBytes, "metadata cache exceeded its serialized cap");
check(metadata.cache.itemCount <= LOCAL_STATE_LIMITS.metadataCache.entries, "metadata cache exceeded its entry cap");
check(Boolean(metadata.cache.itemsById["metadata-0999"]), "preferred new metadata entry was evicted");
check(boundedJsonByteLength(metadata.cache, LOCAL_STATE_LIMITS.metadataCache.serializedBytes) !== null, "metadata cache cannot be serialized within its budget");

const scopeDirectories = Array.from({ length: 1000 }, (_, index) => ({
  name: index.toString(16).padStart(64, "0"),
  modifiedAt: index
}));
const currentScopeId = scopeDirectories.at(-1).name;
const evictedScopes = scopedStateDirectoriesToEvict(scopeDirectories, currentScopeId);
check(evictedScopes.length === 1000 - LOCAL_STATE_LIMITS.diskScopes, "cross-scope quota retained too many scope directories");
check(!evictedScopes.some((entry) => entry.name === currentScopeId), "cross-scope quota selected the active scope for eviction");

console.log(JSON.stringify({
  ok: true,
  checks,
  contentEntriesRetained: bounded.index.itemCount,
  contentSerializedBytes: bounded.serializedBytes,
  contentTextBytes: bounded.textBytes,
  metadataEntriesRetained: metadata.cache.itemCount,
  metadataSerializedBytes: metadata.serializedBytes,
  scopeDirectoriesRetained: scopeDirectories.length - evictedScopes.length
}, null, 2));
