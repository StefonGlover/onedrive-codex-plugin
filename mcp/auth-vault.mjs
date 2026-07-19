import { execFileSync } from "node:child_process";
import {
  chmodSync,
  closeSync,
  existsSync,
  fsyncSync,
  lstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  unlinkSync,
  writeFileSync
} from "node:fs";
import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import { dirname, join, resolve } from "node:path";

const encryptionAlgorithm = "aes-256-gcm";
const envelopeVersion = 1;
const maximumEnvelopeBytes = 256 * 1024;
const additionalAuthenticatedData = Buffer.from("Codex OneDrive auth vault v1", "utf8");

function requirePrivateRegularFile(path, label) {
  const metadata = lstatSync(path);
  if (!metadata.isFile() || metadata.isSymbolicLink()) {
    throw new Error(`${label} must be a regular file and must not be a symbolic link.`);
  }
  if ((metadata.mode & 0o077) !== 0) {
    throw new Error(`${label} must not be readable or writable by group or other users.`);
  }
  return metadata;
}

function decodeEncryptionKey(value) {
  const normalized = String(value || "").trim();
  if (!normalized) throw new Error("The OneDrive auth-vault encryption key is not configured.");
  const key = /^[0-9a-f]{64}$/i.test(normalized)
    ? Buffer.from(normalized, "hex")
    : Buffer.from(normalized, "base64");
  if (key.length !== 32) {
    throw new Error("The OneDrive auth-vault encryption key must decode to exactly 32 bytes.");
  }
  return key;
}

function loadEncryptionKey(environment) {
  const keyFile = environment.ONEDRIVE_TOKEN_ENCRYPTION_KEY_FILE;
  if (keyFile) {
    const resolved = resolve(keyFile);
    requirePrivateRegularFile(resolved, "The OneDrive auth-vault key file");
    return decodeEncryptionKey(readFileSync(resolved, "utf8"));
  }
  return decodeEncryptionKey(environment.ONEDRIVE_TOKEN_ENCRYPTION_KEY);
}

function validateStoredToken(token) {
  if (!token || typeof token !== "object" || Array.isArray(token)) {
    throw new Error("The stored OneDrive authentication payload is invalid.");
  }
  return token;
}

function encryptedFileVault({ environment, storageRoot }) {
  const path = resolve(environment.ONEDRIVE_TOKEN_FILE || join(storageRoot, "auth", "tokens.enc"));

  function read() {
    if (!existsSync(path)) return null;
    const metadata = requirePrivateRegularFile(path, "The encrypted OneDrive auth-vault file");
    if (metadata.size > maximumEnvelopeBytes) throw new Error("The encrypted OneDrive auth-vault file is unexpectedly large.");
    const envelope = JSON.parse(readFileSync(path, "utf8"));
    if (envelope?.version !== envelopeVersion || envelope?.algorithm !== encryptionAlgorithm) {
      throw new Error("The encrypted OneDrive auth-vault file uses an unsupported format.");
    }
    const key = loadEncryptionKey(environment);
    const iv = Buffer.from(envelope.iv || "", "base64");
    const tag = Buffer.from(envelope.tag || "", "base64");
    const ciphertext = Buffer.from(envelope.ciphertext || "", "base64");
    if (iv.length !== 12 || tag.length !== 16 || ciphertext.length === 0) {
      throw new Error("The encrypted OneDrive auth-vault file is malformed.");
    }
    const decipher = createDecipheriv(encryptionAlgorithm, key, iv);
    decipher.setAAD(additionalAuthenticatedData);
    decipher.setAuthTag(tag);
    const plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    return validateStoredToken(JSON.parse(plaintext.toString("utf8")));
  }

  function write(token) {
    validateStoredToken(token);
    const key = loadEncryptionKey(environment);
    const iv = randomBytes(12);
    const cipher = createCipheriv(encryptionAlgorithm, key, iv);
    cipher.setAAD(additionalAuthenticatedData);
    const ciphertext = Buffer.concat([cipher.update(JSON.stringify(token), "utf8"), cipher.final()]);
    const envelope = JSON.stringify({
      version: envelopeVersion,
      algorithm: encryptionAlgorithm,
      iv: iv.toString("base64"),
      tag: cipher.getAuthTag().toString("base64"),
      ciphertext: ciphertext.toString("base64")
    });
    const parent = dirname(path);
    mkdirSync(parent, { recursive: true, mode: 0o700 });
    chmodSync(parent, 0o700);
    const temporaryPath = `${path}.${process.pid}.${randomBytes(8).toString("hex")}.tmp`;
    let descriptor;
    try {
      descriptor = openSync(temporaryPath, "wx", 0o600);
      writeFileSync(descriptor, envelope, "utf8");
      fsyncSync(descriptor);
      closeSync(descriptor);
      descriptor = undefined;
      renameSync(temporaryPath, path);
      chmodSync(path, 0o600);
    } catch (error) {
      if (descriptor !== undefined) {
        try { closeSync(descriptor); } catch {}
      }
      try { unlinkSync(temporaryPath); } catch {}
      throw new Error(`Could not persist the encrypted OneDrive authentication payload: ${error.message}`);
    }
  }

  function remove() {
    if (!existsSync(path)) return false;
    requirePrivateRegularFile(path, "The encrypted OneDrive auth-vault file");
    unlinkSync(path);
    return true;
  }

  return { mode: "encrypted-file", read, write, remove };
}

function macosKeychainVault({ account, service }) {
  return {
    mode: "keychain",
    read() {
      try {
        const raw = execFileSync("security", [
          "find-generic-password", "-a", account, "-s", service, "-w"
        ], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
        return validateStoredToken(JSON.parse(raw));
      } catch {
        return null;
      }
    },
    write(token) {
      validateStoredToken(token);
      try {
        execFileSync("security", [
          "add-generic-password", "-U", "-a", account, "-s", service, "-w", JSON.stringify(token)
        ], { stdio: "ignore" });
      } catch (error) {
        throw new Error(`Could not store OneDrive authentication in macOS Keychain: ${error.message}`);
      }
    },
    remove() {
      try {
        execFileSync("security", ["delete-generic-password", "-a", account, "-s", service], { stdio: "ignore" });
        return true;
      } catch {
        return false;
      }
    }
  };
}

export function createAuthVault({
  account = "tokens",
  environment = process.env,
  platform = process.platform,
  service = "Codex OneDrive",
  storageRoot
} = {}) {
  if (!storageRoot) throw new Error("A storage root is required for OneDrive authentication storage.");
  const requestedMode = String(environment.ONEDRIVE_TOKEN_STORE || "").trim().toLowerCase();
  const mode = requestedMode || (platform === "darwin" ? "keychain" : "encrypted-file");
  if (mode === "keychain") {
    if (platform !== "darwin") throw new Error("The keychain authentication store is available only on macOS.");
    return macosKeychainVault({ account, service });
  }
  if (mode === "encrypted-file") return encryptedFileVault({ environment, storageRoot });
  throw new Error(`Unsupported OneDrive authentication store: ${mode}`);
}
