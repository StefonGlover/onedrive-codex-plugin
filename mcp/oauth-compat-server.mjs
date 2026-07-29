#!/usr/bin/env node

import {
  createCipheriv,
  createDecipheriv,
  createHash,
  createHmac,
  randomBytes,
  timingSafeEqual
} from "node:crypto";
import {
  chmodSync,
  closeSync,
  existsSync,
  fsyncSync,
  lstatSync,
  openSync,
  readFileSync,
  renameSync,
  unlinkSync,
  writeFileSync
} from "node:fs";
import { createServer } from "node:http";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  facadeJwks,
  issueFacadeAccessToken,
  validateFacadeAccessTokenKeyFile
} from "./oauth-facade-token.mjs";

const ENTRA_CONSUMERS_AUTHORIZE_URL =
  "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize";
const ENTRA_CONSUMERS_TOKEN_URL =
  "https://login.microsoftonline.com/consumers/oauth2/v2.0/token";
const CHATGPT_ORIGIN = "https://chatgpt.com";
const MICROSOFT_LOGIN_ORIGIN = "https://login.microsoftonline.com";
const MAX_URL_BYTES = 16 * 1024;
const MAX_FORM_BYTES = 32 * 1024;
const MAX_REGISTRATION_BYTES = 16 * 1024;
const MAX_PROVIDER_BYTES = 256 * 1024;
const MAX_CIMD_BYTES = 64 * 1024;
const MAX_STATE_BYTES = 512;
const MAX_PROVIDER_CODE_BYTES = 3 * 1024;
const UPSTREAM_TIMEOUT_MS = 10_000;
const HEADERS_TIMEOUT_MS = 10_000;
const REQUEST_TIMEOUT_MS = 15_000;
const KEEP_ALIVE_TIMEOUT_MS = 5_000;
const MAX_REQUESTS_PER_SOCKET = 100;
const MAX_CONNECTIONS = 256;
const DEFAULT_TOKEN_CONCURRENCY = 16;
const RATE_WINDOW_MS = 60_000;
const MAX_RATE_KEYS = 4096;
const DEFAULT_AUTHORIZE_ROUTE_BURST = 32;
const DEFAULT_AUTHORIZE_ROUTE_REFILL_PER_SECOND = 0.2;
const DEFAULT_CALLBACK_ROUTE_BURST = 16;
const DEFAULT_CALLBACK_ROUTE_REFILL_PER_SECOND = 0.05;
const REFRESH_STORE_VERSION = 1;
const MAX_REFRESH_STORE_BYTES = 1024 * 1024;
const MAX_REFRESH_PARTITION_BYTES = 320 * 1024;
const MAX_EPHEMERAL_PARTITION_BYTES = 640 * 1024;
const EPHEMERAL_PROJECTION_HEADROOM_BYTES = 64 * 1024;
const MAX_REFRESH_STORE_ENTRIES = 4096;
const MAX_REFRESH_STORE_FAMILIES = 1024;
const MAX_AUTHORIZATION_TRANSACTIONS = 2048;
const MAX_AUTHORIZATION_CODES = 2048;
const REFRESH_HANDLE_TTL_MS = 90 * 24 * 60 * 60 * 1000;
const AUTHORIZATION_TRANSACTION_TTL_MS = 10 * 60 * 1000;
const AUTHORIZATION_CODE_TTL_MS = 5 * 60 * 1000;
const REFRESH_HANDLE_PATTERN = /^odrh_[A-Za-z0-9_-]{43}$/;
const UPSTREAM_STATE_PATTERN = /^odst_[A-Za-z0-9_-]{43}$/;
const FACADE_CODE_PATTERN = /^odac_[A-Za-z0-9_-]{43}$/;
const DCR_CLIENT_ID_PATTERN = /^oddc_[A-Za-z0-9_-]{22}\.[A-Za-z0-9_-]{43}$/;
const REFRESH_HANDLE_HASH_PATTERN = /^[a-f0-9]{64}$/;
const REFRESH_FAMILY_PATTERN = /^[A-Za-z0-9_-]{22}$/;
const BASE64URL_PATTERN = /^[A-Za-z0-9_-]+$/;
const UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const PKCE_CHALLENGE_PATTERN = /^[A-Za-z0-9_-]{43}$/;
const PKCE_VERIFIER_PATTERN = /^[A-Za-z0-9._~-]{43,128}$/;
const PLACEHOLDER_PATTERN = /(?:REPLACE[_ -]?WITH|YOUR[_ -]?|EXAMPLE[_ -]?ONLY|CHANGEME)/i;
const AUTHORIZATION_PARAMETERS = new Set([
  "client_id",
  "redirect_uri",
  "response_type",
  "response_mode",
  "scope",
  "state",
  "code_challenge",
  "code_challenge_method",
  "nonce",
  "display",
  "prompt",
  "max_age",
  "ui_locales",
  "id_token_hint",
  "login_hint",
  "acr_values",
  "claims",
  "domain_hint",
  "resource"
]);
const AUTHORIZATION_CALLBACK_PARAMETERS = new Set([
  "state",
  "code",
  "error",
  "error_description"
]);
const CODE_TOKEN_PARAMETERS = new Set([
  "grant_type",
  "client_id",
  "client_secret",
  "code",
  "redirect_uri",
  "code_verifier",
  "scope",
  "resource"
]);
const PUBLIC_CODE_TOKEN_PARAMETERS = new Set(
  [...CODE_TOKEN_PARAMETERS].filter((name) => name !== "client_secret")
);
const REFRESH_TOKEN_PARAMETERS = new Set([
  "grant_type",
  "client_id",
  "client_secret",
  "refresh_token",
  "scope",
  "resource"
]);
const PUBLIC_REFRESH_TOKEN_PARAMETERS = new Set(
  [...REFRESH_TOKEN_PARAMETERS].filter((name) => name !== "client_secret")
);
const DIAGNOSTIC_PARAMETER_NAMES = new Set([
  ...AUTHORIZATION_PARAMETERS,
  ...CODE_TOKEN_PARAMETERS,
  ...REFRESH_TOKEN_PARAMETERS
]);
const DEFAULT_RATE_LIMITS = Object.freeze({
  authorize: 60,
  callback: 60,
  register: 15,
  token: 30,
});

class RequestError extends Error {
  constructor(status, code, message, headers = {}, reason = code) {
    super(message);
    this.name = "RequestError";
    this.status = status;
    this.code = code;
    this.headers = headers;
    this.reason = reason;
  }
}

function genericInvalidRefreshGrant() {
  return new RequestError(
    400,
    "invalid_grant",
    "Invalid refresh grant.",
    {},
    "invalid_refresh_handle"
  );
}

function genericRefreshStoreFailure() {
  return new RequestError(
    503,
    "temporarily_unavailable",
    "Refresh service is temporarily unavailable.",
    {},
    "refresh_store_unavailable"
  );
}

function refreshStoreState() {
  return {
    version: REFRESH_STORE_VERSION,
    families: [],
    entries: [],
    authorizations: [],
    codes: []
  };
}

function exactObjectKeys(value, allowed) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  return Object.keys(value).every((key) => allowed.has(key));
}

function boundedVaultLimit(value, maximum, name) {
  if (value === undefined) return maximum;
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw new Error(`OAuth refresh-store test limit ${name} is invalid.`);
  }
  return value;
}

function normalizeRefreshStoreLimits(limits = {}) {
  if (!limits || typeof limits !== "object" || Array.isArray(limits)) {
    throw new Error("OAuth refresh-store test limits are invalid.");
  }
  const maxBytes = boundedVaultLimit(
    limits.maxBytes,
    MAX_REFRESH_STORE_BYTES,
    "maxBytes"
  );
  const normalized = {
    maxBytes,
    maxRefreshBytes: boundedVaultLimit(
      limits.maxRefreshBytes,
      Math.min(MAX_REFRESH_PARTITION_BYTES, maxBytes),
      "maxRefreshBytes"
    ),
    maxEphemeralBytes: boundedVaultLimit(
      limits.maxEphemeralBytes,
      Math.min(MAX_EPHEMERAL_PARTITION_BYTES, maxBytes),
      "maxEphemeralBytes"
    ),
    maxEntries: boundedVaultLimit(
      limits.maxEntries,
      MAX_REFRESH_STORE_ENTRIES,
      "maxEntries"
    ),
    maxFamilies: boundedVaultLimit(
      limits.maxFamilies,
      MAX_REFRESH_STORE_FAMILIES,
      "maxFamilies"
    ),
    maxAuthorizations: boundedVaultLimit(
      limits.maxAuthorizations,
      MAX_AUTHORIZATION_TRANSACTIONS,
      "maxAuthorizations"
    ),
    maxCodes: boundedVaultLimit(
      limits.maxCodes,
      MAX_AUTHORIZATION_CODES,
      "maxCodes"
    ),
    refreshTtlMs: boundedVaultLimit(
      limits.refreshTtlMs,
      REFRESH_HANDLE_TTL_MS,
      "refreshTtlMs"
    ),
    authorizationTtlMs: boundedVaultLimit(
      limits.authorizationTtlMs,
      AUTHORIZATION_TRANSACTION_TTL_MS,
      "authorizationTtlMs"
    ),
    codeTtlMs: boundedVaultLimit(
      limits.codeTtlMs,
      AUTHORIZATION_CODE_TTL_MS,
      "codeTtlMs"
    )
  };
  if (normalized.maxRefreshBytes + normalized.maxEphemeralBytes > normalized.maxBytes) {
    throw new Error("OAuth refresh-store partition byte limits exceed the total store limit.");
  }
  return normalized;
}

function base64urlEncodedLength(byteLength) {
  return Math.ceil(byteLength * 4 / 3);
}

function maximumEncryptedRecordBytes(kind) {
  const payload = kind === "authorization"
    ? {
        outerState: "\\".repeat(MAX_STATE_BYTES),
        codeChallenge: "A".repeat(43),
        responseMode: "form_post",
        providerCodeVerifier: "v".repeat(128)
      }
    : {
        providerCode: "\\".repeat(MAX_PROVIDER_CODE_BYTES),
        codeChallenge: "A".repeat(43),
        providerCodeVerifier: "v".repeat(128)
      };
  const plaintextBytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
  const record = {
    hash: "f".repeat(64),
    status: "active",
    expiresAt: Number.MAX_SAFE_INTEGER,
    iv: "A".repeat(16),
    ciphertext: "A".repeat(base64urlEncodedLength(plaintextBytes)),
    tag: "A".repeat(22)
  };
  return Buffer.byteLength(JSON.stringify(record), "utf8") + 1;
}

const MAX_AUTHORIZATION_RECORD_BYTES = maximumEncryptedRecordBytes("authorization");
const MAX_AUTHORIZATION_CODE_RECORD_BYTES =
  maximumEncryptedRecordBytes("authorization_code");

function refreshPartitionBytes(state) {
  return Buffer.byteLength(
    JSON.stringify({
      version: state.version,
      families: state.families,
      entries: state.entries
    }),
    "utf8"
  );
}

function ephemeralPartitionBytes(state) {
  return Buffer.byteLength(
    JSON.stringify({
      version: state.version,
      authorizations: state.authorizations,
      codes: state.codes
    }),
    "utf8"
  );
}

class RefreshHandleVault {
  constructor({
    file,
    encryptionKey,
    now,
    limits = {},
    authorizationAdmissionLimit,
    codeAdmissionLimit
  }) {
    const normalizedLimits = normalizeRefreshStoreLimits(limits);
    this.file = file;
    this.now = now;
    this.maxBytes = normalizedLimits.maxBytes;
    this.maxRefreshBytes = normalizedLimits.maxRefreshBytes;
    this.maxEphemeralBytes = normalizedLimits.maxEphemeralBytes;
    this.maxEntries = normalizedLimits.maxEntries;
    this.maxFamilies = normalizedLimits.maxFamilies;
    this.maxAuthorizations = normalizedLimits.maxAuthorizations;
    this.maxCodes = normalizedLimits.maxCodes;
    this.refreshTtlMs = normalizedLimits.refreshTtlMs;
    this.authorizationTtlMs = normalizedLimits.authorizationTtlMs;
    this.codeTtlMs = normalizedLimits.codeTtlMs;
    this.authorizationAdmissionLimit = boundedVaultLimit(
      authorizationAdmissionLimit,
      this.maxAuthorizations,
      "authorizationAdmissionLimit"
    );
    this.codeAdmissionLimit = boundedVaultLimit(
      codeAdmissionLimit,
      this.maxCodes,
      "codeAdmissionLimit"
    );
    this.key = createHash("sha256")
      .update("onedrive-oauth-compat-refresh-store-v1\u0000", "utf8")
      .update(encryptionKey)
      .digest();
    this.mutationTail = Promise.resolve();
    try {
      this.state = this.loadOrInitialize();
    } catch {
      throw new Error(
        "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE could not be initialized securely."
      );
    }
  }

  serializeMutation(callback) {
    const operation = this.mutationTail.then(callback);
    this.mutationTail = operation.catch(() => {});
    return operation;
  }

  handleHash(handle) {
    return createHash("sha256")
      .update("onedrive-oauth-compat-refresh-handle-v1\u0000", "utf8")
      .update(handle, "utf8")
      .digest("hex");
  }

  aad(entry) {
    return Buffer.from(
      [
        REFRESH_STORE_VERSION,
        entry.handleHash,
        entry.familyId,
        entry.expiresAt
      ].join("|"),
      "utf8"
    );
  }

  encrypt(upstreamRefreshToken, entry) {
    const iv = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", this.key, iv);
    cipher.setAAD(this.aad(entry));
    const ciphertext = Buffer.concat([
      cipher.update(upstreamRefreshToken, "utf8"),
      cipher.final()
    ]);
    return {
      iv: iv.toString("base64url"),
      ciphertext: ciphertext.toString("base64url"),
      tag: cipher.getAuthTag().toString("base64url")
    };
  }

  recordAad(kind, hash, expiresAt) {
    return Buffer.from(
      [REFRESH_STORE_VERSION, kind, hash, expiresAt].join("|"),
      "utf8"
    );
  }

  encryptRecord(kind, hash, expiresAt, payload) {
    const iv = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", this.key, iv);
    cipher.setAAD(this.recordAad(kind, hash, expiresAt));
    const ciphertext = Buffer.concat([
      cipher.update(JSON.stringify(payload), "utf8"),
      cipher.final()
    ]);
    return {
      iv: iv.toString("base64url"),
      ciphertext: ciphertext.toString("base64url"),
      tag: cipher.getAuthTag().toString("base64url")
    };
  }

  decryptRecord(kind, record) {
    try {
      const decipher = createDecipheriv(
        "aes-256-gcm",
        this.key,
        Buffer.from(record.iv, "base64url")
      );
      decipher.setAAD(this.recordAad(kind, record.hash, record.expiresAt));
      decipher.setAuthTag(Buffer.from(record.tag, "base64url"));
      const plaintext = Buffer.concat([
        decipher.update(Buffer.from(record.ciphertext, "base64url")),
        decipher.final()
      ]).toString("utf8");
      const payload = JSON.parse(plaintext);
      if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
        throw new Error("invalid record");
      }
      return payload;
    } catch {
      throw genericRefreshStoreFailure();
    }
  }

  decrypt(entry) {
    try {
      const decipher = createDecipheriv(
        "aes-256-gcm",
        this.key,
        Buffer.from(entry.iv, "base64url")
      );
      decipher.setAAD(this.aad(entry));
      decipher.setAuthTag(Buffer.from(entry.tag, "base64url"));
      const plaintext = Buffer.concat([
        decipher.update(Buffer.from(entry.ciphertext, "base64url")),
        decipher.final()
      ]).toString("utf8");
      if (!requireOpaque(plaintext, { maximum: 32768 })) throw new Error("invalid token");
      return plaintext;
    } catch {
      throw genericRefreshStoreFailure();
    }
  }

  validateState(state) {
    if (
      !exactObjectKeys(
        state,
        new Set(["version", "families", "entries", "authorizations", "codes"])
      )
      || state.version !== REFRESH_STORE_VERSION
      || !Array.isArray(state.families)
      || !Array.isArray(state.entries)
      || !Array.isArray(state.authorizations)
      || !Array.isArray(state.codes)
      || state.families.length > this.maxFamilies
      || state.entries.length > this.maxEntries
      || state.authorizations.length > this.maxAuthorizations
      || state.codes.length > this.maxCodes
    ) {
      throw new Error("invalid refresh store");
    }

    const familyIds = new Set();
    for (const family of state.families) {
      if (
        !exactObjectKeys(family, new Set(["id", "expiresAt", "revokedAt"]))
        || !REFRESH_FAMILY_PATTERN.test(family.id)
        || !Number.isSafeInteger(family.expiresAt)
        || family.expiresAt < 1
        || !(
          family.revokedAt === null
          || (Number.isSafeInteger(family.revokedAt) && family.revokedAt >= 1)
        )
        || familyIds.has(family.id)
      ) {
        throw new Error("invalid refresh store");
      }
      familyIds.add(family.id);
    }

    const handleHashes = new Set();
    const activeKeys = new Set(["handleHash", "familyId", "status", "expiresAt", "iv", "ciphertext", "tag"]);
    const terminalKeys = new Set(["handleHash", "familyId", "status", "expiresAt", "terminalAt"]);
    for (const entry of state.entries) {
      const active = entry?.status === "active";
      if (
        !exactObjectKeys(entry, active ? activeKeys : terminalKeys)
        || !REFRESH_HANDLE_HASH_PATTERN.test(entry.handleHash)
        || !familyIds.has(entry.familyId)
        || !["active", "consumed", "revoked"].includes(entry.status)
        || !Number.isSafeInteger(entry.expiresAt)
        || entry.expiresAt < 1
        || handleHashes.has(entry.handleHash)
      ) {
        throw new Error("invalid refresh store");
      }
      handleHashes.add(entry.handleHash);
      const family = state.families.find((candidate) => candidate.id === entry.familyId);
      if (!family || entry.expiresAt !== family.expiresAt) {
        throw new Error("invalid refresh store");
      }
      if (active) {
        if (
          family.revokedAt !== null
          || !BASE64URL_PATTERN.test(entry.iv)
          || Buffer.from(entry.iv, "base64url").length !== 12
          || !BASE64URL_PATTERN.test(entry.tag)
          || Buffer.from(entry.tag, "base64url").length !== 16
          || !BASE64URL_PATTERN.test(entry.ciphertext)
          || Buffer.from(entry.ciphertext, "base64url").length > 32768
        ) {
          throw new Error("invalid refresh store");
        }
        this.decrypt(entry);
      } else if (
        !Number.isSafeInteger(entry.terminalAt)
        || entry.terminalAt < 1
      ) {
        throw new Error("invalid refresh store");
      }
    }

    const validateEncryptedRecords = (records, kind, payloadValidator) => {
      const hashes = new Set();
      const activeKeys = new Set(["hash", "status", "expiresAt", "iv", "ciphertext", "tag"]);
      const terminalKeys = new Set(["hash", "status", "expiresAt", "terminalAt"]);
      for (const record of records) {
        const active = record?.status === "active";
        if (
          !exactObjectKeys(record, active ? activeKeys : terminalKeys)
          || !REFRESH_HANDLE_HASH_PATTERN.test(record.hash)
          || !["active", "consumed"].includes(record.status)
          || !Number.isSafeInteger(record.expiresAt)
          || record.expiresAt < 1
          || hashes.has(record.hash)
        ) {
          throw new Error("invalid refresh store");
        }
        hashes.add(record.hash);
        if (active) {
          if (
            !BASE64URL_PATTERN.test(record.iv)
            || Buffer.from(record.iv, "base64url").length !== 12
            || !BASE64URL_PATTERN.test(record.tag)
            || Buffer.from(record.tag, "base64url").length !== 16
            || !BASE64URL_PATTERN.test(record.ciphertext)
            || Buffer.from(record.ciphertext, "base64url").length > 32768
          ) {
            throw new Error("invalid refresh store");
          }
          if (!payloadValidator(this.decryptRecord(kind, record))) {
            throw new Error("invalid refresh store");
          }
        } else if (
          !Number.isSafeInteger(record.terminalAt)
          || record.terminalAt < 1
        ) {
          throw new Error("invalid refresh store");
        }
      }
    };

    validateEncryptedRecords(
      state.authorizations,
      "authorization",
      (payload) =>
        exactObjectKeys(
          payload,
          new Set([
            "outerState",
            "codeChallenge",
            "responseMode",
            "providerCodeVerifier"
          ])
        )
        && validOpaqueState(payload.outerState)
        && (
          payload.codeChallenge === null
          || PKCE_CHALLENGE_PATTERN.test(payload.codeChallenge)
        )
        && (
          payload.providerCodeVerifier === undefined
          || PKCE_VERIFIER_PATTERN.test(payload.providerCodeVerifier)
        )
        && (
          payload.codeChallenge !== null
          || PKCE_VERIFIER_PATTERN.test(payload.providerCodeVerifier || "")
        )
        && ["query", "form_post"].includes(payload.responseMode)
    );
    validateEncryptedRecords(
      state.codes,
      "authorization_code",
      (payload) =>
        exactObjectKeys(
          payload,
          new Set(["providerCode", "codeChallenge", "providerCodeVerifier"])
        )
        && validProviderCode(payload.providerCode)
        && (
          payload.codeChallenge === null
          || PKCE_CHALLENGE_PATTERN.test(payload.codeChallenge)
        )
        && (
          payload.providerCodeVerifier === undefined
          || PKCE_VERIFIER_PATTERN.test(payload.providerCodeVerifier)
        )
        && (
          payload.codeChallenge !== null
          || PKCE_VERIFIER_PATTERN.test(payload.providerCodeVerifier || "")
        )
    );
    if (!this.storeBoundsValid(state)) throw new Error("invalid refresh store");
    return state;
  }

  cleanupExpired(state) {
    const currentTime = this.now();
    const liveFamilies = new Set(
      state.families
        .filter((family) => family.expiresAt > currentTime)
        .map((family) => family.id)
    );
    const families = state.families.filter((family) => liveFamilies.has(family.id));
    const entries = state.entries.filter((entry) => liveFamilies.has(entry.familyId));
    const authorizations = state.authorizations.filter(
      (record) => record.expiresAt > currentTime
    );
    const codes = state.codes.filter((record) => record.expiresAt > currentTime);
    const changed =
      families.length !== state.families.length
      || entries.length !== state.entries.length
      || authorizations.length !== state.authorizations.length
      || codes.length !== state.codes.length;
    state.families = families;
    state.entries = entries;
    state.authorizations = authorizations;
    state.codes = codes;
    return changed;
  }

  loadOrInitialize() {
    const parent = lstatSync(dirname(this.file));
    const currentUid = typeof process.getuid === "function" ? process.getuid() : parent.uid;
    if (
      !parent.isDirectory()
      || parent.isSymbolicLink()
      || parent.uid !== currentUid
      || (parent.mode & 0o077) !== 0
    ) {
      throw new Error("invalid refresh store directory");
    }
    if (!existsSync(this.file)) {
      const state = refreshStoreState();
      this.atomicWrite(state);
      return state;
    }
    const metadata = lstatSync(this.file);
    if (
      !metadata.isFile()
      || metadata.isSymbolicLink()
      || metadata.uid !== currentUid
      || metadata.size > this.maxBytes
      || (metadata.mode & 0o077) !== 0
    ) {
      throw new Error("invalid refresh store");
    }
    const body = readFileSync(this.file, "utf8");
    if (Buffer.byteLength(body, "utf8") > this.maxBytes) {
      throw new Error("invalid refresh store");
    }
    const state = this.validateState(JSON.parse(body));
    if (this.cleanupExpired(state)) this.atomicWrite(state);
    return state;
  }

  storeBoundsValid(state) {
    const serialized = `${JSON.stringify(state)}\n`;
    return (
      Buffer.byteLength(serialized, "utf8") <= this.maxBytes
      && refreshPartitionBytes(state) <= this.maxRefreshBytes
      && ephemeralPartitionBytes(state) <= this.maxEphemeralBytes
      && state.entries.length <= this.maxEntries
      && state.families.length <= this.maxFamilies
      && state.authorizations.length <= this.maxAuthorizations
      && state.codes.length <= this.maxCodes
    );
  }

  atomicWrite(state) {
    const serialized = `${JSON.stringify(state)}\n`;
    if (!this.storeBoundsValid(state)) {
      throw genericRefreshStoreFailure();
    }
    const temporary = `${this.file}.tmp-${process.pid}-${randomBytes(8).toString("hex")}`;
    let fileDescriptor;
    try {
      fileDescriptor = openSync(temporary, "wx", 0o600);
      writeFileSync(fileDescriptor, serialized, "utf8");
      fsyncSync(fileDescriptor);
      closeSync(fileDescriptor);
      fileDescriptor = undefined;
      renameSync(temporary, this.file);
      chmodSync(this.file, 0o600);
      const directoryDescriptor = openSync(dirname(this.file), "r");
      try {
        fsyncSync(directoryDescriptor);
      } finally {
        closeSync(directoryDescriptor);
      }
    } catch {
      if (fileDescriptor !== undefined) {
        try {
          closeSync(fileDescriptor);
        } catch {
          // Best-effort cleanup after a failed atomic write.
        }
      }
      try {
        unlinkSync(temporary);
      } catch {
        // The temporary file may already have been renamed or never created.
      }
      throw genericRefreshStoreFailure();
    }
  }

  assertCapacity({ newFamily = false } = {}) {
    if (
      this.state.entries.length >= this.maxEntries
      || (
        newFamily
        && this.state.families.length >= this.maxFamilies
      )
    ) {
      throw genericRefreshStoreFailure();
    }
  }

  newHandle() {
    for (let attempt = 0; attempt < 8; attempt += 1) {
      const handle = `odrh_${randomBytes(32).toString("base64url")}`;
      const handleHash = this.handleHash(handle);
      if (!this.state.entries.some((entry) => entry.handleHash === handleHash)) {
        return { handle, handleHash };
      }
    }
    throw genericRefreshStoreFailure();
  }

  newOpaque(prefix, records) {
    for (let attempt = 0; attempt < 8; attempt += 1) {
      const value = `${prefix}${randomBytes(32).toString("base64url")}`;
      const hash = this.handleHash(value);
      if (!records.some((record) => record.hash === hash)) return { value, hash };
    }
    throw genericRefreshStoreFailure();
  }

  beginAuthorization({
    outerState,
    codeChallenge,
    responseMode,
    providerCodeVerifier = undefined
  }) {
    return this.serializeMutation(() => {
      if (
        !validOpaqueState(outerState)
        || !(
          codeChallenge === null
          || PKCE_CHALLENGE_PATTERN.test(codeChallenge)
        )
        || !(
          providerCodeVerifier === undefined
          || PKCE_VERIFIER_PATTERN.test(providerCodeVerifier)
        )
        || (
          codeChallenge === null
          && !PKCE_VERIFIER_PATTERN.test(providerCodeVerifier || "")
        )
        || !["query", "form_post"].includes(responseMode)
      ) {
        throw new RequestError(400, "invalid_request", "Invalid authorization request.");
      }
      const candidateState = structuredClone(this.state);
      this.cleanupExpired(candidateState);
      if (candidateState.authorizations.length >= this.authorizationAdmissionLimit) {
        throw new RequestError(
          503,
          "temporarily_unavailable",
          "Authorization service is temporarily unavailable.",
          { "Retry-After": "5" },
          "authorization_admission_limit"
        );
      }
      const { value: upstreamState, hash } = this.newOpaque(
        "odst_",
        candidateState.authorizations
      );
      const expiresAt = this.now() + this.authorizationTtlMs;
      const record = {
        hash,
        status: "active",
        expiresAt
      };
      Object.assign(
        record,
        this.encryptRecord("authorization", hash, expiresAt, {
          outerState,
          codeChallenge,
          responseMode,
          ...(providerCodeVerifier === undefined ? {} : { providerCodeVerifier })
        })
      );
      candidateState.authorizations.push(record);
      this.atomicWrite(candidateState);
      this.state = candidateState;
      return upstreamState;
    });
  }

  consumeAuthorization(upstreamState, providerCode = null, admitProviderCode = () => {}) {
    return this.serializeMutation(() => {
      if (!UPSTREAM_STATE_PATTERN.test(upstreamState || "")) {
        throw new RequestError(400, "invalid_request", "Invalid authorization response.");
      }
      const candidateState = structuredClone(this.state);
      const changed = this.cleanupExpired(candidateState);
      const hash = this.handleHash(upstreamState);
      const record = candidateState.authorizations.find(
        (candidate) => candidate.hash === hash
      );
      if (!record || record.status !== "active") {
        if (changed) {
          this.atomicWrite(candidateState);
          this.state = candidateState;
        }
        throw new RequestError(400, "invalid_request", "Invalid authorization response.");
      }
      if (providerCode !== null && !validProviderCode(providerCode)) {
        throw new RequestError(400, "invalid_request", "Invalid authorization response.");
      }
      if (
        providerCode !== null
        && candidateState.codes.length >= this.codeAdmissionLimit
      ) {
        throw new RequestError(
          503,
          "temporarily_unavailable",
          "Authorization service is temporarily unavailable.",
          { "Retry-After": "20" },
          "authorization_code_admission_limit"
        );
      }
      if (providerCode !== null) admitProviderCode();
      const transaction = this.decryptRecord("authorization", record);
      record.status = "consumed";
      record.terminalAt = this.now();
      delete record.iv;
      delete record.ciphertext;
      delete record.tag;

      let facadeCode = null;
      if (providerCode !== null) {
        const generated = this.newOpaque("odac_", candidateState.codes);
        facadeCode = generated.value;
        const expiresAt = this.now() + this.codeTtlMs;
        const codeRecord = {
          hash: generated.hash,
          status: "active",
          expiresAt
        };
        Object.assign(
          codeRecord,
          this.encryptRecord("authorization_code", generated.hash, expiresAt, {
            providerCode,
            codeChallenge: transaction.codeChallenge,
            ...(transaction.providerCodeVerifier === undefined
              ? {}
              : { providerCodeVerifier: transaction.providerCodeVerifier })
          })
        );
        candidateState.codes.push(codeRecord);
      }

      this.atomicWrite(candidateState);
      this.state = candidateState;
      return {
        outerState: transaction.outerState,
        responseMode: transaction.responseMode,
        facadeCode
      };
    });
  }

  consumeAuthorizationCode(facadeCode, codeVerifier) {
    return this.serializeMutation(() => {
      if (
        !FACADE_CODE_PATTERN.test(facadeCode || "")
        || !(
          codeVerifier === null
          || PKCE_VERIFIER_PATTERN.test(codeVerifier)
        )
      ) {
        throw new RequestError(400, "invalid_grant", "Invalid authorization grant.");
      }
      const candidateState = structuredClone(this.state);
      const changed = this.cleanupExpired(candidateState);
      const hash = this.handleHash(facadeCode);
      const record = candidateState.codes.find((candidate) => candidate.hash === hash);
      if (!record || record.status !== "active") {
        if (changed) {
          this.atomicWrite(candidateState);
          this.state = candidateState;
        }
        throw new RequestError(400, "invalid_grant", "Invalid authorization grant.");
      }
      const mapping = this.decryptRecord("authorization_code", record);
      record.status = "consumed";
      record.terminalAt = this.now();
      delete record.iv;
      delete record.ciphertext;
      delete record.tag;
      this.atomicWrite(candidateState);
      this.state = candidateState;

      if (mapping.codeChallenge === null) {
        if (
          codeVerifier !== null
          || !PKCE_VERIFIER_PATTERN.test(mapping.providerCodeVerifier || "")
        ) {
          throw new RequestError(400, "invalid_grant", "Invalid authorization grant.");
        }
      } else {
        if (!PKCE_VERIFIER_PATTERN.test(codeVerifier || "")) {
          throw new RequestError(400, "invalid_grant", "Invalid authorization grant.");
        }
        const derivedChallenge = createHash("sha256")
          .update(codeVerifier, "ascii")
          .digest("base64url");
        const expected = Buffer.from(mapping.codeChallenge, "ascii");
        const actual = Buffer.from(derivedChallenge, "ascii");
        if (
          expected.length !== actual.length
          || !timingSafeEqual(expected, actual)
        ) {
          throw new RequestError(400, "invalid_grant", "Invalid authorization grant.");
        }
      }
      return {
        providerCode: mapping.providerCode,
        providerCodeVerifier: mapping.providerCodeVerifier || null
      };
    });
  }

  issueInitial(upstreamRefreshToken) {
    return this.serializeMutation(() => {
      if (!requireOpaque(upstreamRefreshToken, { maximum: 32768 })) {
        throw genericRefreshStoreFailure();
      }
      const changed = this.cleanupExpired(this.state);
      this.assertCapacity({ newFamily: true });
      const familyId = randomBytes(16).toString("base64url");
      const expiresAt = this.now() + this.refreshTtlMs;
      const { handle, handleHash } = this.newHandle();
      const family = { id: familyId, expiresAt, revokedAt: null };
      const entry = {
        handleHash,
        familyId,
        status: "active",
        expiresAt
      };
      Object.assign(entry, this.encrypt(upstreamRefreshToken, entry));
      this.state.families.push(family);
      this.state.entries.push(entry);
      try {
        this.atomicWrite(this.state);
      } catch (error) {
        this.state.families.pop();
        this.state.entries.pop();
        if (changed) {
          // The failed write leaves the prior on-disk state authoritative.
          this.state = this.loadOrInitialize();
        }
        throw error;
      }
      return handle;
    });
  }

  consume(handle) {
    return this.serializeMutation(() => {
      if (!REFRESH_HANDLE_PATTERN.test(handle || "")) throw genericInvalidRefreshGrant();
      const changed = this.cleanupExpired(this.state);
      const handleHash = this.handleHash(handle);
      const entry = this.state.entries.find((candidate) => candidate.handleHash === handleHash);
      if (!entry) {
        if (changed) this.atomicWrite(this.state);
        throw genericInvalidRefreshGrant();
      }
      const family = this.state.families.find((candidate) => candidate.id === entry.familyId);
      if (!family || family.revokedAt !== null || entry.status === "revoked") {
        if (changed) this.atomicWrite(this.state);
        throw genericInvalidRefreshGrant();
      }
      if (entry.status === "consumed") {
        const currentTime = this.now();
        family.revokedAt = currentTime;
        for (const descendant of this.state.entries) {
          if (descendant.familyId === family.id && descendant.status === "active") {
            descendant.status = "revoked";
            descendant.terminalAt = currentTime;
            delete descendant.iv;
            delete descendant.ciphertext;
            delete descendant.tag;
          }
        }
        this.atomicWrite(this.state);
        throw genericInvalidRefreshGrant();
      }

      const upstreamRefreshToken = this.decrypt(entry);
      entry.status = "consumed";
      entry.terminalAt = this.now();
      delete entry.iv;
      delete entry.ciphertext;
      delete entry.tag;
      this.atomicWrite(this.state);
      return {
        familyId: family.id,
        upstreamRefreshToken
      };
    });
  }

  issueRotated(familyId, upstreamRefreshToken) {
    return this.serializeMutation(() => {
      if (!requireOpaque(upstreamRefreshToken, { maximum: 32768 })) {
        throw genericRefreshStoreFailure();
      }
      this.cleanupExpired(this.state);
      const family = this.state.families.find((candidate) => candidate.id === familyId);
      if (!family || family.revokedAt !== null || family.expiresAt <= this.now()) {
        throw genericInvalidRefreshGrant();
      }
      this.assertCapacity();
      const { handle, handleHash } = this.newHandle();
      const entry = {
        handleHash,
        familyId,
        status: "active",
        expiresAt: family.expiresAt
      };
      Object.assign(entry, this.encrypt(upstreamRefreshToken, entry));
      this.state.entries.push(entry);
      try {
        this.atomicWrite(this.state);
      } catch (error) {
        this.state.entries.pop();
        throw error;
      }
      return handle;
    });
  }
}

function defaultDiagnosticsLogger(entry) {
  process.stderr.write(`${JSON.stringify(entry)}\n`);
}

function sampledDiagnosticsLogger(logger, now) {
  const rateSamples = new Map();
  return (entry) => {
    if (entry.route === "/healthz" && entry.status < 400) return;
    const capacityDiagnostic =
      entry.status === 429
      || entry.reason === "authorize_route_capacity"
      || entry.reason === "callback_route_capacity"
      || entry.reason === "temporarily_unavailable";
    if (capacityDiagnostic) {
      const key = `${entry.route}:${entry.reason}`;
      const currentTime = now();
      const lastSample = rateSamples.get(key);
      if (
        lastSample !== undefined
        && currentTime - lastSample < RATE_WINDOW_MS
      ) {
        return;
      }
      rateSamples.set(key, currentTime);
    }
    logger(entry);
  };
}

function redactedParameterNames(parameters) {
  const names = new Set();
  for (const [name] of parameters) {
    names.add(DIAGNOSTIC_PARAMETER_NAMES.has(name) ? name : "<unsupported>");
  }
  return [...names].sort();
}

function attachDiagnostics(response, diagnostic, logger) {
  response.once("finish", () => {
    const status = response.statusCode;
    const entry = {
      route: diagnostic.route,
      status,
      code: diagnostic.code || (status < 400 ? "ok" : "http_error"),
      reason: diagnostic.reason || (status < 400 ? "request_completed" : "request_failed"),
      parameterNames: [...diagnostic.parameterNames],
      ...(diagnostic.tokenResponse ? { tokenResponse: diagnostic.tokenResponse } : {})
    };
    try {
      logger(entry);
    } catch {
      // Diagnostics must never change OAuth response behavior.
    }
  });
}

function safeTokenResponseDiagnostic(payload, normalizedScope) {
  const accessToken = typeof payload.access_token === "string"
    ? payload.access_token
    : "";
  const parts = accessToken.split(".");
  let claims = null;
  if (parts.length === 3 && parts.every(Boolean) && parts[1].length <= 16_384) {
    try {
      const decoded = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
      if (decoded && typeof decoded === "object" && !Array.isArray(decoded)) {
        claims = Object.fromEntries(
          [
            "aud", "iss", "iat", "nbf", "exp", "scp", "azp",
            "appid", "tid", "ver"
          ]
            .filter((name) => ["string", "number"].includes(typeof decoded[name]))
            .map((name) => [name, decoded[name]])
        );
      }
    } catch {
      claims = null;
    }
  }
  return {
    tokenType: typeof payload.token_type === "string" ? payload.token_type : null,
    expiresIn: Number.isFinite(Number(payload.expires_in))
      ? Number(payload.expires_in)
      : null,
    expiresInType: typeof payload.expires_in,
    extExpiresIn: Number.isFinite(Number(payload.ext_expires_in))
      ? Number(payload.ext_expires_in)
      : null,
    providerScope: typeof payload.scope === "string" ? payload.scope : null,
    normalizedScope,
    hasRefreshToken: typeof payload.refresh_token === "string"
      && payload.refresh_token.length > 0,
    hasIdToken: typeof payload.id_token === "string"
      && payload.id_token.length > 0,
    accessTokenFormat: parts.length === 3 && parts.every(Boolean) ? "jwt" : "opaque",
    claims
  };
}

function required(value, name) {
  const normalized = String(value || "").trim();
  if (!normalized) throw new Error(`${name} is required.`);
  if (PLACEHOLDER_PATTERN.test(normalized)) {
    throw new Error(`${name} contains a deployment placeholder.`);
  }
  return normalized;
}

function secretSetting(env) {
  const direct = String(env.ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET || "");
  if (direct) {
    if (PLACEHOLDER_PATTERN.test(direct)) {
      throw new Error("ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET contains a deployment placeholder.");
    }
    if (direct.length < 16 || direct.length > 4096 || /[\u0000-\u001f\u007f]/.test(direct)) {
      throw new Error("ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET has an invalid length or format.");
    }
    return direct;
  }
  const file = String(env.ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE || "").trim();
  if (file) {
    const value = readFileSync(file, "utf8").trim();
    if (!value) throw new Error("ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE is empty.");
    if (PLACEHOLDER_PATTERN.test(value)) {
      throw new Error("ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE contains a deployment placeholder.");
    }
    if (value.length < 16 || value.length > 4096 || /[\u0000-\u001f\u007f]/.test(value)) {
      throw new Error("ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE has an invalid length or format.");
    }
    return value;
  }
  throw new Error(
    "ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET or ONEDRIVE_OAUTH_COMPAT_CLIENT_SECRET_FILE is required."
  );
}

function strictHttpsUrl(value, name, { rootOnly = false } = {}) {
  const raw = required(value, name);
  let parsed;
  try {
    parsed = new URL(raw);
  } catch {
    throw new Error(`${name} must be an absolute HTTPS URL.`);
  }
  if (parsed.protocol !== "https:") throw new Error(`${name} must use HTTPS.`);
  if (parsed.username || parsed.password || parsed.search || parsed.hash) {
    throw new Error(`${name} must not contain credentials, a query, or a fragment.`);
  }
  if (rootOnly && parsed.pathname !== "/") {
    throw new Error(`${name} must use the origin root with no path.`);
  }
  return parsed;
}

function protectedResourceAliases(value, protectedResource) {
  const raw = String(value || "").trim();
  if (!raw) return Object.freeze([]);
  const aliases = raw.split(/\s+/).filter(Boolean);
  if (!aliases.length || new Set(aliases).size !== aliases.length) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES must contain a unique, non-empty URL set."
    );
  }
  const canonical = new URL(protectedResource);
  const tunnelMatch = /^\/v1\/mcp\/(tunnel_[a-f0-9]{32})$/.exec(
    canonical.pathname
  );
  if (
    canonical.origin !== "https://api.openai.com"
    || !tunnelMatch
  ) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES may be used only with "
        + "an https://api.openai.com/v1/mcp/tunnel_... protected resource."
    );
  }
  for (const alias of aliases) {
    const parsed = strictHttpsUrl(
      alias,
      "ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES"
    );
    if (
      !/^tunnel-service\.gateway\.unified-\d+\.internal\.api\.openai\.org$/.test(
        parsed.hostname
      )
      || parsed.port
      || parsed.pathname !== canonical.pathname
      || parsed.toString() !== alias
    ) {
      throw new Error(
        "ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES must contain only exact "
          + "ChatGPT tunnel-gateway identifiers for the configured tunnel."
      );
    }
  }
  return Object.freeze([...aliases]);
}

function parseExactScopes(value, name = "ONEDRIVE_OAUTH_COMPAT_SCOPES") {
  const raw = required(value, name);
  const scopes = raw.split(/\s+/).filter(Boolean);
  if (!scopes.length || new Set(scopes).size !== scopes.length) {
    throw new Error(`${name} must contain a unique, non-empty scope set.`);
  }
  for (const scope of scopes) {
    if (scope.length > 512 || /[\u0000-\u0020\u007f]/.test(scope)) {
      throw new Error(`${name} contains an invalid scope.`);
    }
  }
  if (scopes.includes("openid") || scopes.includes("profile")) {
    throw new Error(
      `${name} must not include OIDC scopes because this OAuth-only service does not issue ID tokens.`
    );
  }
  if (!scopes.includes("offline_access")) {
    throw new Error(`${name} must include offline_access for refresh-token support.`);
  }
  return Object.freeze([...scopes]);
}

function upstreamScopes(value, outerScopes) {
  if (value === undefined) return outerScopes;
  const scopes = parseExactScopes(
    value,
    "ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES"
  );
  if (
    scopes.length === outerScopes.length
    && scopes.every((scope) => outerScopes.includes(scope))
  ) {
    return scopes;
  }
  const outerResourceScopes = outerScopes.filter(
    (scope) => scope !== "offline_access"
  );
  const upstreamResourceScopes = scopes.filter(
    (scope) => scope !== "offline_access"
  );
  const outerSeparator = outerResourceScopes[0]?.lastIndexOf("/") ?? -1;
  const outerResource = outerResourceScopes[0]?.slice(0, outerSeparator) ?? "";
  const outerClientId = outerResource.startsWith("api://")
    ? outerResource.slice("api://".length)
    : "";
  if (
    outerResourceScopes.length !== 1
    || upstreamResourceScopes.length !== 1
    || outerSeparator < 1
    || !UUID_PATTERN.test(outerClientId)
    || !/^[^/\s]+$/.test(outerResourceScopes[0].slice(outerSeparator + 1))
    || upstreamResourceScopes[0]
      !== `${outerResource}/.default`
  ) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES may differ only by replacing "
        + "the single outer API delegated scope with that same API's /.default scope."
    );
  }
  return scopes;
}

function strictBoolean(value, name, defaultValue = false) {
  if (value === undefined) return defaultValue;
  if (value === "true") return true;
  if (value === "false") return false;
  throw new Error(`${name} must be exactly "true" or "false".`);
}

function outerTokenAuthMethod(value) {
  const method = value === undefined ? "client_secret_post" : value;
  if (method === "client_secret_post" || method === "none") return method;
  throw new Error(
    "ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD must be exactly "
      + '"client_secret_post" or "none".'
  );
}

function accessTokenMode(value) {
  const mode = value === undefined ? "facade" : value;
  if (mode === "facade" || mode === "provider") return mode;
  throw new Error(
    'ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_MODE must be exactly "facade" or "provider".'
  );
}

function refreshStoreFile(value, tokenAuthMethod) {
  if (tokenAuthMethod !== "none") return null;
  const raw = required(
    value,
    "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE"
  );
  if (
    raw === "/"
    || resolve(raw) !== raw
    || /[\u0000-\u001f\u007f]/.test(raw)
  ) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE must be a normalized absolute file path."
    );
  }
  return raw;
}

function refreshStoreEncryptionKey(value, tokenAuthMethod) {
  if (tokenAuthMethod !== "none") return null;
  const raw = required(
    value,
    "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE"
  );
  if (raw === "/" || resolve(raw) !== raw || /[\u0000-\u001f\u007f]/.test(raw)) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE must be a normalized absolute file path."
    );
  }
  let encoded;
  try {
    const metadata = lstatSync(raw);
    const currentUid = typeof process.getuid === "function" ? process.getuid() : metadata.uid;
    if (
      !metadata.isFile()
      || metadata.isSymbolicLink()
      || metadata.uid !== currentUid
      || metadata.size > 4096
      || (metadata.mode & 0o077) !== 0
    ) {
      throw new Error("invalid key file");
    }
    encoded = readFileSync(raw, "utf8").trim();
  } catch {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE must be an owner-only regular file."
    );
  }
  const key = /^[0-9a-f]{64}$/i.test(encoded)
    ? Buffer.from(encoded, "hex")
    : Buffer.from(encoded, "base64");
  if (key.length !== 32) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE must decode to exactly 32 bytes."
    );
  }
  return key;
}

function providerCallbackUri(value, tokenAuthMethod, issuer) {
  if (tokenAuthMethod !== "none") return null;
  const parsed = strictHttpsUrl(
    value,
    "ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI"
  );
  const expected = `${issuer}/callback`;
  if (parsed.toString() !== expected) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI must exactly equal "
        + `${expected}.`
    );
  }
  return parsed.toString();
}

export function oauthCompatSettings(env = process.env) {
  const issuerUrl = strictHttpsUrl(
    env.ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER,
    "ONEDRIVE_OAUTH_COMPAT_PUBLIC_ISSUER",
    { rootOnly: true }
  );
  const resourceUrl = strictHttpsUrl(
    env.ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE,
    "ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE"
  );
  const redirectUrl = strictHttpsUrl(
    env.ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI,
    "ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI"
  );
  if (
    redirectUrl.origin !== CHATGPT_ORIGIN
    || !/^\/connector\/oauth\/[A-Za-z0-9_-]+$/.test(redirectUrl.pathname)
  ) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_REDIRECT_URI must be an exact https://chatgpt.com/connector/oauth/... callback."
    );
  }

  const clientId = required(
    env.ONEDRIVE_OAUTH_COMPAT_CLIENT_ID,
    "ONEDRIVE_OAUTH_COMPAT_CLIENT_ID"
  ).toLowerCase();
  if (!UUID_PATTERN.test(clientId)) {
    throw new Error("ONEDRIVE_OAUTH_COMPAT_CLIENT_ID must be a Microsoft Entra application UUID.");
  }

  const authorizeUrl = required(
    env.ONEDRIVE_OAUTH_COMPAT_UPSTREAM_AUTHORIZE_URL || ENTRA_CONSUMERS_AUTHORIZE_URL,
    "ONEDRIVE_OAUTH_COMPAT_UPSTREAM_AUTHORIZE_URL"
  );
  const tokenUrl = required(
    env.ONEDRIVE_OAUTH_COMPAT_UPSTREAM_TOKEN_URL || ENTRA_CONSUMERS_TOKEN_URL,
    "ONEDRIVE_OAUTH_COMPAT_UPSTREAM_TOKEN_URL"
  );
  if (authorizeUrl !== ENTRA_CONSUMERS_AUTHORIZE_URL) {
    throw new Error(
      `ONEDRIVE_OAUTH_COMPAT_UPSTREAM_AUTHORIZE_URL must equal ${ENTRA_CONSUMERS_AUTHORIZE_URL}.`
    );
  }
  if (tokenUrl !== ENTRA_CONSUMERS_TOKEN_URL) {
    throw new Error(
      `ONEDRIVE_OAUTH_COMPAT_UPSTREAM_TOKEN_URL must equal ${ENTRA_CONSUMERS_TOKEN_URL}.`
    );
  }

  const port = Number(env.ONEDRIVE_OAUTH_COMPAT_PORT || 3010);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("ONEDRIVE_OAUTH_COMPAT_PORT must be an integer from 1 through 65535.");
  }

  const issuer = issuerUrl.origin;
  const configuredOuterTokenAuthMethod = outerTokenAuthMethod(
    env.ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD
  );
  const allowConfidentialNoPkce = strictBoolean(
    env.ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE,
    "ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE"
  );
  const allowPublicNoPkce = strictBoolean(
    env.ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE,
    "ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE"
  );
  const enableCimd = strictBoolean(
    env.ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD,
    "ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD"
  );
  const enableDcr = strictBoolean(
    env.ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR,
    "ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR"
  );
  if (enableCimd && configuredOuterTokenAuthMethod !== "none") {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_ENABLE_CIMD may be true only for a public OAuth client."
    );
  }
  if (enableDcr && configuredOuterTokenAuthMethod !== "none") {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_ENABLE_DCR may be true only for a public OAuth client."
    );
  }
  if (configuredOuterTokenAuthMethod === "none" && allowConfidentialNoPkce) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_ALLOW_CONFIDENTIAL_NO_PKCE must be false "
        + 'when ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD is "none".'
    );
  }
  if (configuredOuterTokenAuthMethod !== "none" && allowPublicNoPkce) {
    throw new Error(
      "ONEDRIVE_OAUTH_COMPAT_ALLOW_PUBLIC_NO_PKCE must be false "
        + 'unless ONEDRIVE_OAUTH_COMPAT_OUTER_TOKEN_AUTH_METHOD is "none".'
    );
  }
  const scopes = parseExactScopes(env.ONEDRIVE_OAUTH_COMPAT_SCOPES);
  const protectedResourceValue = resourceUrl.toString();
  const configuredAccessTokenMode = accessTokenMode(
    env.ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_MODE
  );
  return Object.freeze({
    issuer,
    protectedResource: protectedResourceValue,
    protectedResourceAliases: protectedResourceAliases(
      env.ONEDRIVE_OAUTH_COMPAT_PROTECTED_RESOURCE_ALIASES,
      protectedResourceValue
    ),
    clientId,
    clientSecret: secretSetting(env),
    accessTokenKeyFile: validateFacadeAccessTokenKeyFile(
      env.ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_KEY_FILE
        || env.ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE,
      "ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_KEY_FILE"
    ),
    redirectUri: redirectUrl.toString(),
    scopes,
    upstreamScopes: upstreamScopes(
      env.ONEDRIVE_OAUTH_COMPAT_UPSTREAM_SCOPES,
      scopes
    ),
    authorizeUrl,
    tokenUrl,
    accessTokenMode: configuredAccessTokenMode,
    outerTokenAuthMethod: configuredOuterTokenAuthMethod,
    allowConfidentialNoPkce,
    allowPublicNoPkce,
    enableCimd,
    enableDcr,
    refreshStoreFile: refreshStoreFile(
      env.ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_FILE,
      configuredOuterTokenAuthMethod
    ),
    refreshStoreEncryptionKey: refreshStoreEncryptionKey(
      env.ONEDRIVE_OAUTH_COMPAT_REFRESH_STORE_KEY_FILE,
      configuredOuterTokenAuthMethod
    ),
    providerCallbackUri: providerCallbackUri(
      env.ONEDRIVE_OAUTH_COMPAT_PROVIDER_CALLBACK_URI,
      configuredOuterTokenAuthMethod,
      issuer
    ),
    host: String(env.ONEDRIVE_OAUTH_COMPAT_HOST || "127.0.0.1").trim(),
    port
  });
}

export function validateOAuthCompatConfiguration(env = process.env) {
  const settings = oauthCompatSettings(env);
  if (!settings.host || /[\u0000-\u0020\u007f/]/.test(settings.host)) {
    throw new Error("ONEDRIVE_OAUTH_COMPAT_HOST is invalid.");
  }
  return settings;
}

function commonHeaders(response, request, { allowMethods = "GET" } = {}) {
  response.setHeader("Cache-Control", "no-store");
  response.setHeader("Pragma", "no-cache");
  response.setHeader("X-Content-Type-Options", "nosniff");
  response.setHeader("X-Frame-Options", "DENY");
  response.setHeader("Referrer-Policy", "no-referrer");
  response.setHeader("Content-Security-Policy", "default-src 'none'; frame-ancestors 'none'");
  const origin = request.headers.origin;
  if (origin === CHATGPT_ORIGIN) {
    response.setHeader("Access-Control-Allow-Origin", CHATGPT_ORIGIN);
    response.setHeader("Vary", "Origin");
    response.setHeader("Access-Control-Allow-Methods", allowMethods);
    response.setHeader("Access-Control-Allow-Headers", "Content-Type");
    response.setHeader("Access-Control-Max-Age", "600");
  }
}

function sendJson(response, request, status, payload, headers = {}) {
  commonHeaders(response, request, { allowMethods: "GET, POST, OPTIONS" });
  for (const [name, value] of Object.entries(headers)) response.setHeader(name, value);
  const body = `${JSON.stringify(payload)}\n`;
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  response.end(body);
}

function sendEmpty(response, request, status, headers = {}) {
  commonHeaders(response, request, { allowMethods: "GET, POST, OPTIONS" });
  for (const [name, value] of Object.entries(headers)) response.setHeader(name, value);
  response.writeHead(status);
  response.end();
}

function sendHtml(response, request, status, body, headers = {}) {
  commonHeaders(response, request, { allowMethods: "GET, POST, OPTIONS" });
  for (const [name, value] of Object.entries(headers)) response.setHeader(name, value);
  response.writeHead(status, {
    "Content-Type": "text/html; charset=utf-8",
    "Content-Length": Buffer.byteLength(body)
  });
  response.end(body);
}

function rejectBadOrigin(request, { allowMicrosoft = false } = {}) {
  const origin = request.headers.origin;
  if (
    origin
    && origin !== CHATGPT_ORIGIN
    && !(allowMicrosoft && origin === MICROSOFT_LOGIN_ORIGIN)
  ) {
    throw new RequestError(403, "access_denied", "Cross-origin request denied.", {}, "origin_rejected");
  }
}

function rejectDuplicateParameters(parameters) {
  const seen = new Set();
  for (const [name] of parameters) {
    if (seen.has(name)) {
      throw new RequestError(400, "invalid_request", "Duplicate OAuth parameter.", {}, "duplicate_parameter");
    }
    seen.add(name);
  }
}

function enforceAllowedParameters(parameters, allowed) {
  for (const [name] of parameters) {
    if (!allowed.has(name)) {
      throw new RequestError(400, "invalid_request", "Unsupported OAuth parameter.", {}, "unsupported_parameter");
    }
  }
}

function requireExact(parameters, name, expected, reason = "registered_value_mismatch") {
  if (parameters.get(name) !== expected) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth request does not match the registered client.",
      {},
      reason
    );
  }
}

function requireOpaque(value, { minimum = 1, maximum = 8192 } = {}) {
  return (
    typeof value === "string"
    && value.length >= minimum
    && value.length <= maximum
    && !/[\u0000-\u001f\u007f]/.test(value)
  );
}

function validOpaqueState(value) {
  return requireOpaque(value, { minimum: 1, maximum: MAX_STATE_BYTES })
    && Buffer.byteLength(value, "utf8") <= MAX_STATE_BYTES;
}

function validProviderCode(value) {
  return requireOpaque(value, { minimum: 1, maximum: MAX_PROVIDER_CODE_BYTES })
    && Buffer.byteLength(value, "utf8") <= MAX_PROVIDER_CODE_BYTES;
}

function authorizationRateMaterial(state, codeChallenge) {
  return createHash("sha256")
    .update("onedrive-oauth-compat-authorization-rate-v1\u0000", "utf8")
    .update(state, "utf8")
    .update("\u0000", "utf8")
    .update(codeChallenge || "", "utf8")
    .digest("base64url");
}

function canonicalAllowedScopeSet(value, configuredScopes, { allowMissing = false } = {}) {
  const resourceScopes = configuredScopes.filter((scope) => scope !== "offline_access");
  if (value === null || value === undefined) {
    return allowMissing ? resourceScopes : null;
  }
  if (typeof value !== "string" || !value.trim()) return null;
  const requested = value.trim().split(/\s+/);
  if (new Set(requested).size !== requested.length) return null;
  const configured = new Set(configuredScopes);
  if (!requested.every((scope) => configured.has(scope))) return null;
  const requestedSet = new Set(requested);
  const exactFullSet =
    requested.length === configuredScopes.length
    && configuredScopes.every((scope) => requestedSet.has(scope));
  const exactResourceSet =
    requested.length === resourceScopes.length
    && resourceScopes.every((scope) => requestedSet.has(scope));
  if (!exactFullSet && !exactResourceSet) return null;
  return configuredScopes.filter((scope) => requestedSet.has(scope));
}

function hasAllowedScopeSet(value, configuredScopes) {
  return canonicalAllowedScopeSet(value, configuredScopes) !== null;
}

function validateOptionalAuthorizationParameters(parameters) {
  const responseMode = parameters.get("response_mode");
  if (responseMode && !new Set(["query", "form_post"]).has(responseMode)) {
    throw new RequestError(
      400,
      "invalid_request",
      "Unsupported OAuth response mode.",
      {},
      "unsupported_response_mode"
    );
  }
  const display = parameters.get("display");
  if (display && !new Set(["page", "popup", "touch", "wap"]).has(display)) {
    throw new RequestError(400, "invalid_request", "Unsupported OIDC display value.", {}, "unsupported_display");
  }
  const prompt = parameters.get("prompt");
  if (prompt) {
    const values = prompt.trim().split(/\s+/);
    const supported = new Set(["none", "login", "consent", "select_account"]);
    if (!values.length || new Set(values).size !== values.length || values.some((value) => !supported.has(value))) {
      throw new RequestError(400, "invalid_request", "Unsupported OIDC prompt value.", {}, "unsupported_prompt");
    }
    if (values.includes("none") && values.length !== 1) {
      throw new RequestError(400, "invalid_request", "OIDC prompt none cannot be combined.", {}, "invalid_prompt");
    }
  }
  const maxAge = parameters.get("max_age");
  if (maxAge && !/^(?:0|[1-9][0-9]{0,9})$/.test(maxAge)) {
    throw new RequestError(400, "invalid_request", "Invalid OIDC max_age value.", {}, "invalid_max_age");
  }
  const claims = parameters.get("claims");
  if (claims) {
    if (claims.length > 8192) {
      throw new RequestError(400, "invalid_request", "OIDC claims value is too large.", {}, "invalid_claims");
    }
    try {
      const parsed = JSON.parse(claims);
      if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("not an object");
    } catch {
      throw new RequestError(400, "invalid_request", "Invalid OIDC claims value.", {}, "invalid_claims");
    }
  }
  for (const name of [
    "nonce",
    "ui_locales",
    "id_token_hint",
    "login_hint",
    "acr_values",
    "domain_hint"
  ]) {
    const value = parameters.get(name);
    if (value && !requireOpaque(value, { maximum: name === "id_token_hint" ? 16384 : 2048 })) {
      throw new RequestError(
        400,
        "invalid_request",
        "Invalid optional OAuth parameter.",
        {},
        `invalid_${name}`
      );
    }
  }
}

function authorizationMetadata(settings) {
  const metadata = {
    issuer: settings.issuer,
    authorization_endpoint: `${settings.issuer}/authorize`,
    token_endpoint: `${settings.issuer}/token`,
    scopes_supported: settings.scopes,
    response_types_supported: ["code"],
    response_modes_supported: ["query", "form_post"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    token_endpoint_auth_methods_supported: [settings.outerTokenAuthMethod]
  };
  if (settings.accessTokenMode === "facade") {
    metadata.jwks_uri = `${settings.issuer}/jwks.json`;
  }
  if (!settings.allowPublicNoPkce) {
    metadata.code_challenge_methods_supported = ["S256"];
  }
  if (settings.enableCimd) {
    metadata.client_id_metadata_document_supported = true;
  }
  if (settings.enableDcr) {
    metadata.registration_endpoint = `${settings.issuer}/register`;
  }
  return metadata;
}

function protectedResourceMetadata(settings) {
  const authorizationServerOnlyScopes = new Set([
    "offline_access",
    "openid",
    "profile",
    "email"
  ]);
  return {
    resource: settings.protectedResource,
    authorization_servers: [settings.issuer],
    scopes_supported: settings.scopes.filter(
      (scope) => !authorizationServerOnlyScopes.has(scope)
    ),
    bearer_methods_supported: ["header"]
  };
}

async function readRegistrationJson(request) {
  const contentType = String(request.headers["content-type"] || "").toLowerCase();
  if (!/^application\/json(?:\s*;\s*charset=utf-8)?$/.test(contentType)) {
    throw new RequestError(
      415,
      "invalid_client_metadata",
      "Client registration must use JSON.",
      {},
      "registration_content_type"
    );
  }
  const declared = Number(request.headers["content-length"]);
  if (Number.isFinite(declared) && declared > MAX_REGISTRATION_BYTES) {
    throw new RequestError(
      413,
      "invalid_client_metadata",
      "Client registration is too large.",
      {},
      "registration_too_large"
    );
  }
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_REGISTRATION_BYTES) {
      throw new RequestError(
        413,
        "invalid_client_metadata",
        "Client registration is too large.",
        {},
        "registration_too_large"
      );
    }
    chunks.push(chunk);
  }
  if (!chunks.length) {
    throw new RequestError(
      400,
      "invalid_client_metadata",
      "Client registration is empty.",
      {},
      "registration_empty"
    );
  }
  try {
    const value = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("invalid");
    return value;
  } catch {
    throw new RequestError(
      400,
      "invalid_client_metadata",
      "Client registration JSON is invalid.",
      {},
      "registration_invalid_json"
    );
  }
}

async function handleRegistration(
  request,
  response,
  settings,
  enforceRate,
  now,
  diagnostic
) {
  if (!settings.enableDcr) {
    throw new RequestError(404, "not_found", "Not found.", {}, "dcr_disabled");
  }
  if (request.method !== "POST") {
    diagnostic.code = "invalid_request";
    diagnostic.reason = "method_not_allowed";
    sendJson(
      response,
      request,
      405,
      { error: "invalid_request" },
      { Allow: "POST, OPTIONS" }
    );
    return;
  }
  if (request.headers.authorization) {
    throw new RequestError(
      401,
      "invalid_client",
      "Client registration authentication is not supported.",
      {},
      "registration_authentication_rejected"
    );
  }
  enforceRate("register", "chatgpt-dynamic-client-registration");
  const metadata = await readRegistrationJson(request);
  const allowedKeys = new Set([
    "client_name",
    "client_uri",
    "contacts",
    "grant_types",
    "jwks",
    "jwks_uri",
    "logo_uri",
    "policy_uri",
    "redirect_uris",
    "response_types",
    "scope",
    "software_id",
    "software_version",
    "token_endpoint_auth_method",
    "tos_uri"
  ]);
  if (!exactObjectKeys(metadata, allowedKeys)) {
    throw new RequestError(
      400,
      "invalid_client_metadata",
      "Client registration metadata is unsupported.",
      {},
      "registration_unknown_metadata"
    );
  }
  if (
    !Array.isArray(metadata.redirect_uris)
    || metadata.redirect_uris.length !== 1
    || metadata.redirect_uris[0] !== settings.redirectUri
    || metadata.token_endpoint_auth_method !== "none"
    || (
      metadata.response_types !== undefined
      && (
        !Array.isArray(metadata.response_types)
        || metadata.response_types.length !== 1
        || metadata.response_types[0] !== "code"
      )
    )
    || (
      metadata.grant_types !== undefined
      && (
        !Array.isArray(metadata.grant_types)
        || !metadata.grant_types.includes("authorization_code")
        || metadata.grant_types.some(
          (value) => !new Set(["authorization_code", "refresh_token"]).has(value)
        )
      )
    )
    || (
      metadata.scope !== undefined
      && !hasAllowedScopeSet(metadata.scope, settings.scopes)
    )
  ) {
    throw new RequestError(
      400,
      "invalid_client_metadata",
      "Client registration metadata does not match this server.",
      {},
      "registration_metadata_mismatch"
    );
  }
  diagnostic.code = "created";
  diagnostic.reason = "public_pkce_client_registered";
  const nonce = randomBytes(16).toString("base64url");
  const clientId = `oddc_${nonce}.${createHmac(
    "sha256",
    settings.refreshStoreEncryptionKey
  ).update(`onedrive-dcr-client-v1\u0000${nonce}`, "utf8").digest("base64url")}`;
  sendJson(response, request, 201, {
    client_id: clientId,
    client_id_issued_at: Math.floor(now() / 1000),
    redirect_uris: [settings.redirectUri],
    token_endpoint_auth_method: "none",
    grant_types: ["authorization_code", "refresh_token"],
    response_types: ["code"],
    scope: settings.scopes.join(" "),
    ...(typeof metadata.client_name === "string"
      && requireOpaque(metadata.client_name, { maximum: 256 })
      ? { client_name: metadata.client_name }
      : {})
  });
}

function exactSingleValue(parameters, name, expected) {
  const values = parameters.getAll(name);
  return values.length === 1 && values[0] === expected;
}

function requireProtectedResource(parameters, settings) {
  const values = parameters.getAll("resource");
  if (
    values.length !== 1
    || !new Set([
      settings.protectedResource,
      ...settings.protectedResourceAliases
    ]).has(values[0])
  ) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth request does not match the registered client.",
      {},
      "resource_mismatch"
    );
  }
}

function cimdClientUrl(value) {
  if (typeof value !== "string" || Buffer.byteLength(value, "utf8") > 2048) return null;
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    return null;
  }
  if (
    parsed.origin !== CHATGPT_ORIGIN
    || parsed.username
    || parsed.password
    || parsed.search
    || parsed.hash
    || !/^\/oauth(?:\/[A-Za-z0-9._~-]+)+\/client\.json$/.test(parsed.pathname)
    || parsed.toString() !== value
  ) {
    return null;
  }
  return parsed.toString();
}

async function readBoundedCimd(response) {
  const declared = Number(response.headers.get("content-length"));
  if (Number.isFinite(declared) && declared > MAX_CIMD_BYTES) {
    throw new RequestError(400, "invalid_request", "OAuth client metadata is invalid.", {}, "cimd_too_large");
  }
  if (!response.body) {
    throw new RequestError(400, "invalid_request", "OAuth client metadata is invalid.", {}, "cimd_missing_body");
  }
  const reader = response.body.getReader();
  const chunks = [];
  let size = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > MAX_CIMD_BYTES) {
        await reader.cancel();
        throw new RequestError(400, "invalid_request", "OAuth client metadata is invalid.", {}, "cimd_too_large");
      }
      chunks.push(Buffer.from(value));
    }
  } finally {
    reader.releaseLock();
  }
  try {
    const value = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("invalid");
    return value;
  } catch {
    throw new RequestError(400, "invalid_request", "OAuth client metadata is invalid.", {}, "cimd_invalid_json");
  }
}

async function validateCimdClient(clientId, settings, fetchImpl) {
  const url = settings.enableCimd ? cimdClientUrl(clientId) : null;
  if (!url) return false;
  let response;
  try {
    response = await fetchImpl(url, {
      method: "GET",
      headers: { Accept: "application/json" },
      redirect: "error",
      signal: AbortSignal.timeout(5_000)
    });
  } catch {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth client metadata could not be verified.",
      {},
      "cimd_fetch_failed"
    );
  }
  if (!response.ok) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth client metadata could not be verified.",
      {},
      "cimd_fetch_rejected"
    );
  }
  const metadata = await readBoundedCimd(response);
  const tokenMethods = Array.isArray(metadata.token_endpoint_auth_methods_supported)
    ? metadata.token_endpoint_auth_methods_supported
    : metadata.token_endpoint_auth_method
      ? [metadata.token_endpoint_auth_method]
      : [];
  if (
    (metadata.client_id !== undefined && metadata.client_id !== url)
    || !Array.isArray(metadata.redirect_uris)
    || !metadata.redirect_uris.includes(settings.redirectUri)
    || !tokenMethods.includes("none")
    || (Array.isArray(metadata.response_types) && !metadata.response_types.includes("code"))
    || (
      Array.isArray(metadata.grant_types)
      && !metadata.grant_types.includes("authorization_code")
    )
  ) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth client metadata is not compatible with this server.",
      {},
      "cimd_metadata_mismatch"
    );
  }
  return true;
}

function validDcrClientId(value, settings) {
  if (
    !settings.enableDcr
    || typeof value !== "string"
    || !DCR_CLIENT_ID_PATTERN.test(value)
  ) {
    return false;
  }
  const [nonce, signature] = value.slice("oddc_".length).split(".");
  const expected = createHmac("sha256", settings.refreshStoreEncryptionKey)
    .update(`onedrive-dcr-client-v1\u0000${nonce}`, "utf8")
    .digest("base64url");
  return constantTimeEqual(signature, expected);
}

function trustedClientIdValue(value, settings) {
  return (
    value === settings.clientId
    || validDcrClientId(value, settings)
    || Boolean(settings.enableCimd && cimdClientUrl(value))
  );
}

async function requireTrustedAuthorizationTarget(parameters, settings, fetchImpl) {
  const clientIds = parameters.getAll("client_id");
  if (clientIds.length !== 1) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth client is not trusted.",
      {},
      "untrusted_client_id"
    );
  }
  if (
    clientIds[0] !== settings.clientId
    && !validDcrClientId(clientIds[0], settings)
    && !await validateCimdClient(clientIds[0], settings, fetchImpl)
  ) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth client is not trusted.",
      {},
      "untrusted_client_id"
    );
  }
  if (!exactSingleValue(parameters, "redirect_uri", settings.redirectUri)) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth redirect URI is not trusted.",
      {},
      "untrusted_redirect_uri"
    );
  }
}

function safeAuthorizationErrorCode(value) {
  const code = String(value || "");
  return /^[A-Za-z0-9._~-]{1,64}$/.test(code) ? code : "invalid_request";
}

function safeAuthorizationErrorDescription(value) {
  return String(value || "")
    .replace(/[^\x20-\x21\x23-\x5B\x5D-\x7E]/g, " ")
    .slice(0, 180);
}

function authorizationErrorParameters(parameters, error) {
  const output = new URLSearchParams();
  output.set("error", safeAuthorizationErrorCode(error.code));
  const description = safeAuthorizationErrorDescription(error.message);
  if (description) output.set("error_description", description);
  const states = parameters.getAll("state");
  if (states.length === 1 && validOpaqueState(states[0])) {
    output.set("state", states[0]);
  }
  return output;
}

function htmlAttribute(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function sendOuterAuthorizationResponse(
  request,
  response,
  settings,
  responseMode,
  output
) {
  if (responseMode === "form_post") {
    const nonce = randomBytes(18).toString("base64url");
    const inputs = [...output].map(([name, value]) =>
      `<input type="hidden" name="${htmlAttribute(name)}" value="${htmlAttribute(value)}">`
    ).join("");
    const body = [
      "<!doctype html><html><head><meta charset=\"utf-8\"><title>OAuth response</title></head><body>",
      `<form method="post" action="${htmlAttribute(settings.redirectUri)}">${inputs}`,
      "<noscript><button type=\"submit\">Continue</button></noscript></form>",
      `<script nonce="${nonce}">document.forms[0].submit();</script>`,
      "</body></html>"
    ].join("");
    sendHtml(response, request, 200, body, {
      "Content-Security-Policy":
        `default-src 'none'; base-uri 'none'; form-action ${CHATGPT_ORIGIN}; script-src 'nonce-${nonce}'; frame-ancestors 'none'`
    });
    return;
  }

  const callback = new URL(settings.redirectUri);
  for (const [name, value] of output) callback.searchParams.set(name, value);
  sendEmpty(response, request, 302, { Location: callback.toString() });
}

function sendAuthorizationError(request, response, parameters, settings, error, diagnostic) {
  const output = authorizationErrorParameters(parameters, error);
  const responseModes = parameters.getAll("response_mode");
  const responseMode =
    responseModes.length === 1 && responseModes[0] === "form_post"
      ? "form_post"
      : "query";
  diagnostic.code = safeAuthorizationErrorCode(error.code);
  diagnostic.reason = error.reason || "authorization_validation_failed";
  sendOuterAuthorizationResponse(
    request,
    response,
    settings,
    responseMode,
    output
  );
}

async function handleAuthorize(
  request,
  response,
  url,
  settings,
  refreshVault,
  fetchImpl,
  enforceRate,
  consumeAuthorizeCapacity,
  diagnostic
) {
  if (request.method !== "GET") {
    diagnostic.code = "invalid_request";
    diagnostic.reason = "method_not_allowed";
    sendJson(response, request, 405, { error: "invalid_request" }, { Allow: "GET, OPTIONS" });
    return;
  }
  const parameters = url.searchParams;
  const hasCodeChallenge = parameters.has("code_challenge");
  const hasCodeChallengeMethod = parameters.has("code_challenge_method");
  await requireTrustedAuthorizationTarget(parameters, settings, fetchImpl);
  try {
    rejectDuplicateParameters(parameters);
    enforceAllowedParameters(parameters, AUTHORIZATION_PARAMETERS);
    requireProtectedResource(parameters, settings);
    requireExact(parameters, "response_type", "code", "response_type_mismatch");
    if (hasCodeChallenge !== hasCodeChallengeMethod) {
      throw new RequestError(
        400,
        "invalid_request",
        "Incomplete PKCE parameters.",
        {},
        "incomplete_pkce"
      );
    }
    const publicOuterClient = settings.outerTokenAuthMethod === "none";
    if (
      !hasCodeChallenge
      && (
        publicOuterClient
          ? !settings.allowPublicNoPkce
          : !settings.allowConfidentialNoPkce
      )
    ) {
      throw new RequestError(
        400,
        "invalid_request",
        "PKCE is required.",
        {},
        "pkce_required"
      );
    }
    if (hasCodeChallenge) {
      requireExact(parameters, "code_challenge_method", "S256", "pkce_method_mismatch");
      if (!PKCE_CHALLENGE_PATTERN.test(parameters.get("code_challenge") || "")) {
        throw new RequestError(
          400,
          "invalid_request",
          "Invalid PKCE code challenge.",
          {},
          "invalid_code_challenge"
        );
      }
    }
    if (!validOpaqueState(parameters.get("state"))) {
      throw new RequestError(400, "invalid_request", "Invalid OAuth state.", {}, "invalid_state");
    }
    if (!hasAllowedScopeSet(parameters.get("scope"), settings.scopes)) {
      throw new RequestError(
        400,
        "invalid_scope",
        "OAuth scope set does not match the registered client.",
        {},
        "scope_mismatch"
      );
    }
    validateOptionalAuthorizationParameters(parameters);
    enforceRate(
      "authorize",
      authorizationRateMaterial(
        parameters.get("state"),
        parameters.get("code_challenge")
      )
    );
    if (settings.outerTokenAuthMethod === "none") consumeAuthorizeCapacity();
  } catch (error) {
    if (error instanceof RequestError) {
      sendAuthorizationError(request, response, parameters, settings, error, diagnostic);
      return;
    }
    throw error;
  }

  const upstream = new URL(settings.authorizeUrl);
  if (settings.outerTokenAuthMethod === "none") {
    let providerCodeVerifier;
    let providerCodeChallenge = parameters.get("code_challenge");
    if (!hasCodeChallenge) {
      providerCodeVerifier = randomBytes(32).toString("base64url");
      providerCodeChallenge = createHash("sha256")
        .update(providerCodeVerifier, "ascii")
        .digest("base64url");
    }
    const upstreamState = await refreshVault.beginAuthorization({
      outerState: parameters.get("state"),
      codeChallenge: parameters.get("code_challenge"),
      responseMode: parameters.get("response_mode") || "query",
      ...(providerCodeVerifier === undefined ? {} : { providerCodeVerifier })
    });
    for (const [name, value] of parameters) {
      if (["client_id", "resource", "redirect_uri", "state", "scope"].includes(name)) continue;
      upstream.searchParams.append(name, value);
    }
    if (!hasCodeChallenge) {
      upstream.searchParams.set("code_challenge", providerCodeChallenge);
      upstream.searchParams.set("code_challenge_method", "S256");
    }
    upstream.searchParams.set("redirect_uri", settings.providerCallbackUri);
    upstream.searchParams.set("client_id", settings.clientId);
    upstream.searchParams.set("state", upstreamState);
    upstream.searchParams.set("response_mode", "query");
    upstream.searchParams.set("scope", settings.upstreamScopes.join(" "));
  } else {
    for (const [name, value] of parameters) {
      if (!["resource", "scope"].includes(name)) {
        upstream.searchParams.append(name, value);
      }
    }
    upstream.searchParams.set("scope", settings.upstreamScopes.join(" "));
  }
  diagnostic.code = "redirect";
  diagnostic.reason = "microsoft_authorization_redirect";
  commonHeaders(response, request);
  response.writeHead(302, { Location: upstream.toString() });
  response.end();
}

async function handleAuthorizationCallback(
  request,
  response,
  url,
  settings,
  refreshVault,
  enforceRate,
  consumeCallbackCapacity,
  diagnostic
) {
  if (settings.outerTokenAuthMethod !== "none") {
    throw new RequestError(404, "not_found", "Not found.", {}, "route_not_found");
  }
  if (request.method !== "GET") {
    diagnostic.code = "invalid_request";
    diagnostic.reason = "method_not_allowed";
    sendJson(
      response,
      request,
      405,
      { error: "invalid_request" },
      { Allow: "GET" }
    );
    return;
  }
  const parameters = url.searchParams;

  diagnostic.parameterNames = redactedParameterNames(parameters);
  rejectDuplicateParameters(parameters);
  enforceAllowedParameters(parameters, AUTHORIZATION_CALLBACK_PARAMETERS);
  const state = parameters.get("state");
  const code = parameters.get("code");
  const providerError = parameters.get("error");
  if (
    !UPSTREAM_STATE_PATTERN.test(state || "")
    || Boolean(code) === Boolean(providerError)
    || (parameters.has("error_description") && !providerError)
  ) {
    throw new RequestError(400, "invalid_request", "Invalid authorization response.");
  }
  enforceRate("callback", state);

  const transaction = await refreshVault.consumeAuthorization(
    state,
    code || null,
    code ? consumeCallbackCapacity : () => {}
  );
  const output = new URLSearchParams();
  if (code) {
    output.set("code", transaction.facadeCode);
    diagnostic.code = "redirect";
    diagnostic.reason = "facade_authorization_code_issued";
  } else {
    output.set("error", safeAuthorizationErrorCode(providerError));
    const rawDescription = parameters.get("error_description") || "";
    const description = rawDescription.includes(settings.clientSecret)
      ? ""
      : safeAuthorizationErrorDescription(rawDescription);
    if (description) output.set("error_description", description);
    diagnostic.code = safeAuthorizationErrorCode(providerError);
    diagnostic.reason = "provider_authorization_error";
  }
  output.set("state", transaction.outerState);
  sendOuterAuthorizationResponse(
    request,
    response,
    settings,
    transaction.responseMode,
    output
  );
}

async function readForm(request) {
  const contentType = String(request.headers["content-type"] || "").toLowerCase();
  if (!/^application\/x-www-form-urlencoded(?:\s*;\s*charset=utf-8)?$/.test(contentType)) {
    throw new RequestError(415, "invalid_request", "Token request must use form encoding.");
  }
  const declared = Number(request.headers["content-length"]);
  if (Number.isFinite(declared) && declared > MAX_FORM_BYTES) {
    throw new RequestError(413, "invalid_request", "Token request is too large.");
  }
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_FORM_BYTES) {
      throw new RequestError(413, "invalid_request", "Token request is too large.");
    }
    chunks.push(chunk);
  }
  if (!chunks.length) throw new RequestError(400, "invalid_request", "Token request is empty.");
  const parameters = new URLSearchParams(Buffer.concat(chunks).toString("utf8"));
  rejectDuplicateParameters(parameters);
  return parameters;
}

function constantTimeEqual(left, right) {
  const leftDigest = createHash("sha256").update(String(left || ""), "utf8").digest();
  const rightDigest = createHash("sha256").update(String(right || ""), "utf8").digest();
  return timingSafeEqual(leftDigest, rightDigest);
}

function validateTokenParameters(request, parameters, settings) {
  if (request.headers.authorization) {
    throw new RequestError(
      401,
      "invalid_client",
      "Unsupported token client authentication.",
      {},
      "unsupported_client_authentication"
    );
  }
  const grantType = parameters.get("grant_type");
  const publicOuterClient = settings.outerTokenAuthMethod === "none";
  const allowed = grantType === "authorization_code"
    ? publicOuterClient
      ? PUBLIC_CODE_TOKEN_PARAMETERS
      : CODE_TOKEN_PARAMETERS
    : grantType === "refresh_token"
      ? publicOuterClient
        ? PUBLIC_REFRESH_TOKEN_PARAMETERS
        : REFRESH_TOKEN_PARAMETERS
      : null;
  if (!allowed) {
    throw new RequestError(
      400,
      "unsupported_grant_type",
      "Unsupported OAuth grant.",
      {},
      "unsupported_grant_type"
    );
  }
  if (publicOuterClient && parameters.has("client_secret")) {
    throw new RequestError(
      400,
      "invalid_request",
      "Client secret is not accepted for this public OAuth client.",
      {},
      "outer_client_secret_forbidden"
    );
  }
  enforceAllowedParameters(parameters, allowed);
  requireProtectedResource(parameters, settings);
  if (!trustedClientIdValue(parameters.get("client_id"), settings)) {
    throw new RequestError(
      400,
      "invalid_request",
      "OAuth client does not match the registered client.",
      {},
      "client_id_mismatch"
    );
  }
  if (
    !publicOuterClient
    && !constantTimeEqual(parameters.get("client_secret"), settings.clientSecret)
  ) {
    throw new RequestError(401, "invalid_client", "Invalid OAuth client.", {}, "invalid_client_secret");
  }
  const scope = parameters.get("scope");
  if (scope !== null && !hasAllowedScopeSet(scope, settings.scopes)) {
    throw new RequestError(
      400,
      "invalid_scope",
      "OAuth scope set does not match the registered client.",
      {},
      "scope_mismatch"
    );
  }

  if (grantType === "authorization_code") {
    requireExact(parameters, "redirect_uri", settings.redirectUri, "redirect_uri_mismatch");
    if (
      publicOuterClient
        ? !FACADE_CODE_PATTERN.test(parameters.get("code") || "")
        : !requireOpaque(parameters.get("code"), { maximum: 16384 })
    ) {
      throw new RequestError(400, "invalid_grant", "Invalid authorization grant.", {}, "invalid_code");
    }
    const codeVerifier = parameters.get("code_verifier");
    if (
      codeVerifier === null
      && (
        publicOuterClient
          ? !settings.allowPublicNoPkce
          : !settings.allowConfidentialNoPkce
      )
    ) {
      throw new RequestError(400, "invalid_grant", "PKCE verifier is required.", {}, "pkce_required");
    }
    if (codeVerifier !== null && !PKCE_VERIFIER_PATTERN.test(codeVerifier)) {
      throw new RequestError(400, "invalid_grant", "Invalid PKCE verifier.", {}, "invalid_code_verifier");
    }
  } else {
    const refreshToken = parameters.get("refresh_token");
    if (
      publicOuterClient
        ? !REFRESH_HANDLE_PATTERN.test(refreshToken || "")
        : !requireOpaque(refreshToken, { maximum: 32768 })
    ) {
      throw new RequestError(400, "invalid_grant", "Invalid refresh grant.", {}, "invalid_refresh_token");
    }
  }
}

async function readBoundedProviderJson(providerResponse) {
  const declared = Number(providerResponse.headers.get("content-length"));
  if (Number.isFinite(declared) && declared > MAX_PROVIDER_BYTES) {
    throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
  }
  if (!providerResponse.body) {
    throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
  }
  const reader = providerResponse.body.getReader();
  const chunks = [];
  let size = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > MAX_PROVIDER_BYTES) {
        await reader.cancel();
        throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
      }
      chunks.push(Buffer.from(value));
    }
  } finally {
    reader.releaseLock();
  }
  let payload;
  try {
    payload = JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
  }
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
  }
  return payload;
}

async function proxyToken(
  request,
  response,
  settings,
  refreshVault,
  fetchImpl,
  enforceRate,
  acquireTokenSlot,
  diagnostic
) {
  if (request.method !== "POST") {
    diagnostic.code = "invalid_request";
    diagnostic.reason = "method_not_allowed";
    sendJson(response, request, 405, { error: "invalid_request" }, { Allow: "POST, OPTIONS" });
    return;
  }
  const releaseTokenSlot = acquireTokenSlot();
  try {
    const publicOuterClient = settings.outerTokenAuthMethod === "none";
    const parameters = await readForm(request);
    diagnostic.parameterNames = redactedParameterNames(parameters);
    validateTokenParameters(request, parameters, settings);
    const responseScopes = canonicalAllowedScopeSet(
      parameters.get("scope"),
      settings.scopes,
      { allowMissing: true }
    );
    if (!responseScopes) {
      throw new RequestError(
        400,
        "invalid_scope",
        "OAuth scope set does not match the registered client.",
        {},
        "scope_mismatch"
      );
    }
    const responseScope = responseScopes.join(" ");

    const grantType = parameters.get("grant_type");
    enforceRate(
      "token",
      publicOuterClient
        ? grantType === "authorization_code"
          ? parameters.get("code")
          : parameters.get("refresh_token")
        : parameters.get("client_id")
    );
    let refreshContext = null;
    let mappedProviderAuthorization = null;
    if (publicOuterClient && grantType === "authorization_code") {
      mappedProviderAuthorization = await refreshVault.consumeAuthorizationCode(
        parameters.get("code"),
        parameters.get("code_verifier")
      );
    }
    if (publicOuterClient && grantType === "refresh_token") {
      refreshContext = await refreshVault.consume(parameters.get("refresh_token"));
    }

    const upstreamParameters = new URLSearchParams();
    for (const [name, value] of parameters) {
      if (name === "resource" || name === "scope") continue;
      if (publicOuterClient && name === "client_id") continue;
      if (publicOuterClient && grantType === "authorization_code" && name === "code") {
        upstreamParameters.append(name, mappedProviderAuthorization.providerCode);
        continue;
      }
      if (
        publicOuterClient
        && grantType === "authorization_code"
        && name === "code_verifier"
        && mappedProviderAuthorization.providerCodeVerifier
      ) {
        continue;
      }
      if (
        publicOuterClient
        && grantType === "authorization_code"
        && name === "redirect_uri"
      ) {
        upstreamParameters.append(name, settings.providerCallbackUri);
        continue;
      }
      if (publicOuterClient && grantType === "refresh_token" && name === "refresh_token") {
        upstreamParameters.append(name, refreshContext.upstreamRefreshToken);
        continue;
      }
      upstreamParameters.append(name, value);
    }
    upstreamParameters.set("scope", settings.upstreamScopes.join(" "));
    if (publicOuterClient) {
      upstreamParameters.set("client_id", settings.clientId);
      upstreamParameters.set("client_secret", settings.clientSecret);
      if (
        grantType === "authorization_code"
        && mappedProviderAuthorization.providerCodeVerifier
      ) {
        upstreamParameters.set(
          "code_verifier",
          mappedProviderAuthorization.providerCodeVerifier
        );
      }
    }

    let providerResponse;
    try {
      providerResponse = await fetchImpl(settings.tokenUrl, {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/x-www-form-urlencoded"
        },
        body: upstreamParameters.toString(),
        redirect: "error",
        signal: AbortSignal.timeout(UPSTREAM_TIMEOUT_MS)
      });
    } catch {
      throw new RequestError(502, "server_error", "OAuth provider request failed.");
    }
    if (providerResponse.status < 200 || providerResponse.status > 599) {
      throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
    }
    const payload = await readBoundedProviderJson(providerResponse);
    if (JSON.stringify(payload).includes(settings.clientSecret)) {
      throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
    }
    let responsePayload = providerResponse.ok
      ? { ...payload, scope: responseScope }
      : payload;
    if (providerResponse.ok) {
      if (
        String(payload.token_type || "").toLowerCase() !== "bearer"
        || !requireOpaque(payload.access_token, { maximum: 32_768 })
        || !Number.isSafeInteger(payload.expires_in)
        || payload.expires_in < 1
      ) {
        throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
      }
      diagnostic.tokenResponse = safeTokenResponseDiagnostic(
        payload,
        responseScope
      );
      if (settings.accessTokenMode === "facade") {
        const requestedResource = parameters.get("resource");
        const facade = issueFacadeAccessToken({
          providerAccessToken: payload.access_token,
          issuer: settings.issuer,
          audience: requestedResource,
          clientId: settings.clientId,
          scope: responseScope,
          expiresIn: payload.expires_in,
          keyFile: settings.accessTokenKeyFile
        });
        responsePayload.access_token = facade.accessToken;
        responsePayload.expires_in = facade.expiresIn;
        diagnostic.tokenResponse.outerAudienceClass =
          requestedResource === settings.protectedResource
            ? "canonical"
            : settings.protectedResourceAliases.includes(requestedResource)
              ? "same_tunnel_alias"
              : "invalid";
      }
      diagnostic.tokenResponse.outerAccessTokenBytes = Buffer.byteLength(
        responsePayload.access_token,
        "utf8"
      );
      diagnostic.tokenResponse.outerAccessTokenFormat = settings.accessTokenMode === "facade"
        ? "jwt"
        : "provider-jwt";
      if (Object.hasOwn(responsePayload, "ext_expires_in")) {
        responsePayload.ext_expires_in = responsePayload.expires_in;
      }
    }
    if (publicOuterClient) {
      const providerRefreshToken = responsePayload.refresh_token;
      delete responsePayload.refresh_token;
      if (
        typeof providerRefreshToken === "string"
        && JSON.stringify(responsePayload).includes(providerRefreshToken)
      ) {
        throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
      }
      if (providerResponse.ok) {
        if (
          String(responsePayload.token_type || "").toLowerCase() !== "bearer"
          || !requireOpaque(responsePayload.access_token, { maximum: 32768 })
          || !Number.isSafeInteger(responsePayload.expires_in)
          || responsePayload.expires_in < 1
        ) {
          throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
        }
        let outerRefreshHandle;
        if (grantType === "authorization_code") {
          if (!requireOpaque(providerRefreshToken, { maximum: 32768 })) {
            throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
          }
          outerRefreshHandle = await refreshVault.issueInitial(providerRefreshToken);
        } else {
          if (
            providerRefreshToken !== undefined
            && !requireOpaque(providerRefreshToken, { maximum: 32768 })
          ) {
            throw new RequestError(502, "server_error", "OAuth provider response was invalid.");
          }
          outerRefreshHandle = await refreshVault.issueRotated(
            refreshContext.familyId,
            providerRefreshToken || refreshContext.upstreamRefreshToken
          );
        }
        responsePayload.refresh_token = outerRefreshHandle;
        responsePayload.scope = responseScope;
      }
    }
    const headers = {};
    const retryAfter = providerResponse.headers.get("retry-after");
    if (retryAfter && /^(?:[1-9][0-9]{0,4})$/.test(retryAfter)) headers["Retry-After"] = retryAfter;
    diagnostic.code = providerResponse.ok ? "ok" : "provider_error";
    diagnostic.reason = "microsoft_token_response";
    sendJson(response, request, providerResponse.status, responsePayload, headers);
  } finally {
    releaseTokenSlot();
  }
}

function createRateLimiter(rateLimits, now, capacity) {
  const buckets = new Map();
  return (category, keyMaterial) => {
    const limit = rateLimits[category];
    if (
      !Number.isInteger(limit)
      || typeof keyMaterial !== "string"
      || keyMaterial.length < 1
      || Buffer.byteLength(keyMaterial, "utf8") > MAX_STATE_BYTES
    ) {
      throw new Error("OAuth rate-limit key is invalid.");
    }
    const key = createHash("sha256")
      .update("onedrive-oauth-compat-rate-limit-v1\u0000", "utf8")
      .update(category, "utf8")
      .update("\u0000", "utf8")
      .update(keyMaterial, "utf8")
      .digest("hex");
    const currentTime = now();
    let bucket = buckets.get(key);
    if (
      bucket
      && (
        currentTime < bucket.startedAt
        || currentTime >= bucket.resetAt
      )
    ) {
      buckets.delete(key);
      bucket = null;
    }
    if (!bucket) {
      bucket = {
        count: 0,
        startedAt: currentTime,
        resetAt: currentTime + RATE_WINDOW_MS
      };
      if (buckets.size >= capacity) {
        for (const [entryKey, entry] of buckets) {
          if (entry.resetAt <= currentTime) buckets.delete(entryKey);
        }
      }
      while (buckets.size >= capacity) {
        const oldest = buckets.keys().next().value;
        if (oldest === undefined) break;
        buckets.delete(oldest);
      }
      buckets.set(key, bucket);
    } else {
      buckets.delete(key);
      buckets.set(key, bucket);
    }
    bucket.count += 1;
    if (bucket.count > limit) {
      const seconds = Math.max(1, Math.ceil((bucket.resetAt - currentTime) / 1000));
      throw new RequestError(429, "temporarily_unavailable", "Request rate limit exceeded.", {
        "Retry-After": String(seconds)
      });
    }
  };
}

function createConcurrencyLimiter(limit) {
  let active = 0;
  return () => {
    if (active >= limit) {
      throw new RequestError(
        503,
        "temporarily_unavailable",
        "Token service is temporarily busy.",
        { "Retry-After": "1", Connection: "close" },
        "token_concurrency_exceeded"
      );
    }
    active += 1;
    let released = false;
    return () => {
      if (released) return;
      released = true;
      active -= 1;
    };
  };
}

function createRefillingTokenBucket({
  burst,
  refillPerSecond,
  now,
  retryReason
}) {
  let tokens = burst;
  let lastRefill = now();
  return () => {
    const currentTime = now();
    if (currentTime < lastRefill) {
      lastRefill = currentTime;
    } else if (currentTime > lastRefill) {
      tokens = Math.min(
        burst,
        tokens + (currentTime - lastRefill) * refillPerSecond / 1000
      );
      lastRefill = currentTime;
    }
    if (tokens < 1) {
      const retryAfter = Math.max(
        1,
        Math.ceil((1 - tokens) / refillPerSecond)
      );
      throw new RequestError(
        429,
        "temporarily_unavailable",
        "Authorization service is temporarily busy.",
        { "Retry-After": String(retryAfter) },
        retryReason
      );
    }
    tokens -= 1;
  };
}

function routeAdmissionEnvelope(burst, refillPerSecond, ttlMs) {
  return burst + refillPerSecond * ttlMs / 1000;
}

function safeMcpHealthProjection(value) {
  if (!value || typeof value !== "object" || value.ok !== true) return null;
  const projection = {
    ok: true,
    authMode: String(value.authMode || "").slice(0, 32)
  };
  if (value.lastAuthFailure && typeof value.lastAuthFailure === "object") {
    projection.lastAuthFailure = {
      at: String(value.lastAuthFailure.at || "").slice(0, 64),
      authMode: String(value.lastAuthFailure.authMode || "").slice(0, 32),
      code: String(value.lastAuthFailure.code || "").slice(0, 128),
      status: Number.isInteger(value.lastAuthFailure.status)
        ? value.lastAuthFailure.status
        : 500
    };
  }
  if (value.lastToolFailure && typeof value.lastToolFailure === "object") {
    projection.lastToolFailure = {
      at: String(value.lastToolFailure.at || "").slice(0, 64),
      tool: String(value.lastToolFailure.tool || "").slice(0, 128),
      code: String(value.lastToolFailure.code || "").slice(0, 128),
      ...(Number.isInteger(value.lastToolFailure.graphStatus)
        ? { graphStatus: value.lastToolFailure.graphStatus }
        : {})
    };
  }
  if (value.lastToolCall && typeof value.lastToolCall === "object") {
    projection.lastToolCall = {
      at: String(value.lastToolCall.at || "").slice(0, 64),
      tool: String(value.lastToolCall.tool || "").slice(0, 128),
      isError: value.lastToolCall.isError === true
    };
  }
  return projection;
}

async function readMcpHealth(fetchImpl, env) {
  const configuredHost = String(env.ONEDRIVE_MCP_HTTP_HOST || "127.0.0.1").trim();
  const host = new Set(["0.0.0.0", "::", "[::]"]).has(configuredHost)
    ? "127.0.0.1"
    : configuredHost;
  const port = Number(env.ONEDRIVE_MCP_HTTP_PORT || 3001);
  if (!Number.isInteger(port) || port < 1 || port > 65535) return null;
  try {
    const response = await fetchImpl(`http://${host}:${port}/healthz`, {
      method: "GET",
      redirect: "error",
      signal: AbortSignal.timeout(750)
    });
    if (!response.ok) return null;
    return safeMcpHealthProjection(await response.json());
  } catch {
    return null;
  }
}

export function createOAuthCompatServer(
  env = process.env,
  {
    fetchImpl = globalThis.fetch,
    now = Date.now,
    rateLimits = DEFAULT_RATE_LIMITS,
    rateKeyCapacity = MAX_RATE_KEYS,
    tokenConcurrency = DEFAULT_TOKEN_CONCURRENCY,
    authorizeRouteBurst = DEFAULT_AUTHORIZE_ROUTE_BURST,
    authorizeRouteRefillPerSecond = DEFAULT_AUTHORIZE_ROUTE_REFILL_PER_SECOND,
    callbackRouteBurst = DEFAULT_CALLBACK_ROUTE_BURST,
    callbackRouteRefillPerSecond = DEFAULT_CALLBACK_ROUTE_REFILL_PER_SECOND,
    diagnostics = defaultDiagnosticsLogger,
    refreshStoreLimits = {}
  } = {}
) {
  const settings = validateOAuthCompatConfiguration(env);
  if (typeof fetchImpl !== "function") throw new Error("A Fetch API implementation is required.");
  if (typeof diagnostics !== "function") throw new Error("OAuth diagnostics must be a function.");
  const sampledDiagnosticLogger = sampledDiagnosticsLogger(diagnostics, now);
  let lastOAuthRequest = null;
  let lastTokenResponse = null;
  const diagnosticLogger = (entry) => {
    if (new Set(["/authorize", "/callback", "/register", "/token"]).has(entry.route)) {
      lastOAuthRequest = {
        at: new Date(now()).toISOString(),
        route: entry.route,
        status: entry.status,
        code: entry.code,
        reason: entry.reason,
        parameterNames: [...entry.parameterNames]
      };
    }
    if (entry.route === "/token" && entry.tokenResponse) {
      lastTokenResponse = entry.tokenResponse;
    }
    sampledDiagnosticLogger(entry);
  };
  const normalizedRateLimits = { ...DEFAULT_RATE_LIMITS, ...rateLimits };
  for (const [name, value] of Object.entries(normalizedRateLimits)) {
    if (!Number.isInteger(value) || value < 1 || value > 10_000) {
      throw new Error(`OAuth compatibility rate limit ${name} is invalid.`);
    }
  }
  if (!Number.isSafeInteger(rateKeyCapacity) || rateKeyCapacity < 1 || rateKeyCapacity > MAX_RATE_KEYS) {
    throw new Error("OAuth compatibility rate-key capacity is invalid.");
  }
  if (!Number.isSafeInteger(tokenConcurrency) || tokenConcurrency < 1 || tokenConcurrency > MAX_CONNECTIONS) {
    throw new Error("OAuth compatibility token concurrency is invalid.");
  }
  for (const [name, burst] of [
    ["authorize", authorizeRouteBurst],
    ["callback", callbackRouteBurst]
  ]) {
    if (!Number.isSafeInteger(burst) || burst < 1 || burst > MAX_AUTHORIZATION_TRANSACTIONS) {
      throw new Error(`OAuth compatibility ${name} route burst is invalid.`);
    }
  }
  for (const [name, refill] of [
    ["authorize", authorizeRouteRefillPerSecond],
    ["callback", callbackRouteRefillPerSecond]
  ]) {
    if (!Number.isFinite(refill) || refill <= 0 || refill > 100) {
      throw new Error(`OAuth compatibility ${name} route refill is invalid.`);
    }
  }

  const enforceRate = createRateLimiter(
    normalizedRateLimits,
    now,
    rateKeyCapacity
  );
  const acquireTokenSlot = createConcurrencyLimiter(tokenConcurrency);
  const consumeAuthorizeCapacity = createRefillingTokenBucket({
    burst: authorizeRouteBurst,
    refillPerSecond: authorizeRouteRefillPerSecond,
    now,
    retryReason: "authorize_route_capacity"
  });
  const consumeCallbackCapacity = createRefillingTokenBucket({
    burst: callbackRouteBurst,
    refillPerSecond: callbackRouteRefillPerSecond,
    now,
    retryReason: "callback_route_capacity"
  });

  let refreshVault = null;
  if (settings.outerTokenAuthMethod === "none") {
    const normalizedStoreLimits = normalizeRefreshStoreLimits(refreshStoreLimits);
    const authorizationEnvelope = routeAdmissionEnvelope(
      authorizeRouteBurst,
      authorizeRouteRefillPerSecond,
      normalizedStoreLimits.authorizationTtlMs
    );
    const codeEnvelope = routeAdmissionEnvelope(
      callbackRouteBurst,
      callbackRouteRefillPerSecond,
      normalizedStoreLimits.codeTtlMs
    );
    const authorizationAdmissionLimit = Math.ceil(authorizationEnvelope);
    const codeAdmissionLimit = Math.ceil(codeEnvelope);
    if (authorizationAdmissionLimit >= normalizedStoreLimits.maxAuthorizations) {
      throw new Error(
        "OAuth authorize route capacity can exhaust the authorization transaction count."
      );
    }
    if (codeAdmissionLimit >= normalizedStoreLimits.maxCodes) {
      throw new Error(
        "OAuth callback route capacity can exhaust the authorization-code count."
      );
    }
    const projectedEphemeralBytes =
      ephemeralPartitionBytes(refreshStoreState())
      + authorizationAdmissionLimit * MAX_AUTHORIZATION_RECORD_BYTES
      + codeAdmissionLimit * MAX_AUTHORIZATION_CODE_RECORD_BYTES;
    if (
      projectedEphemeralBytes + EPHEMERAL_PROJECTION_HEADROOM_BYTES
      > normalizedStoreLimits.maxEphemeralBytes
    ) {
      throw new Error(
        "OAuth route capacity can exhaust the ephemeral refresh-store byte budget."
      );
    }
    refreshVault = new RefreshHandleVault({
        file: settings.refreshStoreFile,
        encryptionKey: settings.refreshStoreEncryptionKey,
        now,
        limits: normalizedStoreLimits,
        authorizationAdmissionLimit,
        codeAdmissionLimit
      });
  }

  const server = createServer(async (request, response) => {
    const diagnostic = {
      route: "<unparsed>",
      code: "",
      reason: "",
      parameterNames: []
    };
    attachDiagnostics(response, diagnostic, diagnosticLogger);
    try {
      if (Buffer.byteLength(request.url || "", "utf8") > MAX_URL_BYTES) {
        throw new RequestError(414, "invalid_request", "Request URL is too large.", {}, "url_too_large");
      }
      const url = new URL(request.url || "/", "http://localhost");
      diagnostic.route = url.pathname;
      diagnostic.parameterNames = redactedParameterNames(url.searchParams);
      rejectBadOrigin(request, { allowMicrosoft: url.pathname === "/callback" });

      if (request.method === "OPTIONS") {
        if (request.headers.origin !== CHATGPT_ORIGIN) {
          throw new RequestError(
            403,
            "access_denied",
            "Cross-origin request denied.",
            {},
            "origin_rejected"
          );
        }
        diagnostic.code = "ok";
        diagnostic.reason = "cors_preflight";
        sendEmpty(response, request, 204);
        return;
      }
      if (url.pathname === "/healthz") {
        if (request.method !== "GET") {
          diagnostic.code = "invalid_request";
          diagnostic.reason = "method_not_allowed";
          sendJson(response, request, 405, { error: "invalid_request" }, { Allow: "GET, OPTIONS" });
          return;
        }
        diagnostic.code = "ok";
        diagnostic.reason = "health_check";
        const mcp = await readMcpHealth(fetchImpl, env);
        sendJson(response, request, 200, {
          ok: true,
          service: "onedrive-oauth-compat",
          provider: "microsoft-entra-consumers",
          ...(lastOAuthRequest ? { lastOAuthRequest } : {}),
          ...(lastTokenResponse ? { lastTokenResponse } : {}),
          ...(mcp ? { mcp } : {})
        });
        return;
      }
      const protectedResourceMetadataRoutes = new Set([
        "/.well-known/oauth-protected-resource",
        "/.well-known/oauth-protected-resource/mcp",
        `/.well-known/oauth-protected-resource${
          new URL(settings.protectedResource).pathname.replace(/\/+$/, "")
        }`
      ]);
      if (protectedResourceMetadataRoutes.has(url.pathname)) {
        if (request.method !== "GET") {
          diagnostic.code = "invalid_request";
          diagnostic.reason = "method_not_allowed";
          sendJson(response, request, 405, { error: "invalid_request" }, { Allow: "GET, OPTIONS" });
          return;
        }
        if (url.search) {
          throw new RequestError(
            400,
            "invalid_request",
            "Metadata query is not supported.",
            {},
            "metadata_query_rejected"
          );
        }
        diagnostic.code = "ok";
        diagnostic.reason = "protected_resource_metadata";
        sendJson(response, request, 200, protectedResourceMetadata(settings));
        return;
      }
      if (
        url.pathname === "/.well-known/oauth-authorization-server"
        || url.pathname === "/.well-known/openid-configuration"
      ) {
        if (request.method !== "GET") {
          diagnostic.code = "invalid_request";
          diagnostic.reason = "method_not_allowed";
          sendJson(response, request, 405, { error: "invalid_request" }, { Allow: "GET, OPTIONS" });
          return;
        }
        if (url.search) {
          throw new RequestError(
            400,
            "invalid_request",
            "Metadata query is not supported.",
            {},
            "metadata_query_rejected"
          );
        }
        diagnostic.code = "ok";
        diagnostic.reason = "authorization_metadata";
        sendJson(response, request, 200, authorizationMetadata(settings));
        return;
      }
      if (url.pathname === "/jwks.json") {
        if (request.method !== "GET") {
          diagnostic.code = "invalid_request";
          diagnostic.reason = "method_not_allowed";
          sendJson(response, request, 405, { error: "invalid_request" }, { Allow: "GET, OPTIONS" });
          return;
        }
        if (url.search) {
          throw new RequestError(
            400,
            "invalid_request",
            "JWKS query is not supported.",
            {},
            "jwks_query_rejected"
          );
        }
        diagnostic.code = "ok";
        diagnostic.reason = "jwks";
        sendJson(response, request, 200, facadeJwks(settings.accessTokenKeyFile));
        return;
      }
      if (url.pathname === "/register") {
        await handleRegistration(
          request,
          response,
          settings,
          enforceRate,
          now,
          diagnostic
        );
        return;
      }
      if (url.pathname === "/authorize") {
        await handleAuthorize(
          request,
          response,
          url,
          settings,
          refreshVault,
          fetchImpl,
          enforceRate,
          consumeAuthorizeCapacity,
          diagnostic
        );
        return;
      }
      if (url.pathname === "/callback") {
        await handleAuthorizationCallback(
          request,
          response,
          url,
          settings,
          refreshVault,
          enforceRate,
          consumeCallbackCapacity,
          diagnostic
        );
        return;
      }
      if (url.pathname === "/token") {
        await proxyToken(
          request,
          response,
          settings,
          refreshVault,
          fetchImpl,
          enforceRate,
          acquireTokenSlot,
          diagnostic
        );
        return;
      }
      diagnostic.code = "not_found";
      diagnostic.reason = "route_not_found";
      sendJson(response, request, 404, { error: "not_found" });
    } catch (error) {
      const safe = error instanceof RequestError
        ? error
        : new RequestError(500, "server_error", "OAuth compatibility service failed.");
      diagnostic.code = safe.code;
      diagnostic.reason = safe.reason || safe.code;
      sendJson(response, request, safe.status, { error: safe.code }, safe.headers);
    }
  });
  server.headersTimeout = HEADERS_TIMEOUT_MS;
  server.requestTimeout = REQUEST_TIMEOUT_MS;
  server.keepAliveTimeout = KEEP_ALIVE_TIMEOUT_MS;
  server.maxRequestsPerSocket = MAX_REQUESTS_PER_SOCKET;
  server.maxConnections = MAX_CONNECTIONS;
  return server;
}

async function main() {
  const settings = validateOAuthCompatConfiguration();
  const server = createOAuthCompatServer();
  await new Promise((resolvePromise, reject) => {
    server.once("error", reject);
    server.listen(settings.port, settings.host, resolvePromise);
  });
  process.stderr.write(
    `OneDrive OAuth compatibility service listening on ${settings.host}:${settings.port}\n`
  );
  const stop = () => server.close(() => process.exit(0));
  process.once("SIGINT", stop);
  process.once("SIGTERM", stop);
}

const isMainModule =
  process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMainModule) {
  main().catch(() => {
    process.stderr.write("OneDrive OAuth compatibility service failed to start.\n");
    process.exit(1);
  });
}
