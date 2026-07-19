#!/usr/bin/env node

import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { chmodSync, lstatSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createAuthVault } from "../mcp/auth-vault.mjs";

const root = mkdtempSync(join(tmpdir(), "onedrive-auth-vault-test-"));
const keyPath = join(root, "vault.key");
const vaultPath = join(root, "tokens.enc");
const first = {
  access_token: "access-value-that-must-not-appear-on-disk",
  refresh_token: "refresh-value-that-must-not-appear-on-disk",
  expires_at: 12345
};
const second = { ...first, refresh_token: "rotated-refresh-value", expires_at: 67890 };

function makeVault(encryptionKey = randomBytes(32)) {
  writeFileSync(keyPath, `${encryptionKey.toString("base64")}\n`, { mode: 0o600 });
  chmodSync(keyPath, 0o600);
  return createAuthVault({
    platform: "linux",
    storageRoot: root,
    environment: {
      ONEDRIVE_TOKEN_STORE: "encrypted-file",
      ONEDRIVE_TOKEN_FILE: vaultPath,
      ONEDRIVE_TOKEN_ENCRYPTION_KEY_FILE: keyPath
    }
  });
}

try {
  const vault = makeVault();
  assert.equal(vault.mode, "encrypted-file");
  assert.equal(vault.read(), null);
  vault.write(first);
  assert.deepEqual(vault.read(), first);
  assert.equal(lstatSync(vaultPath).mode & 0o777, 0o600);
  const storedEnvelope = readFileSync(vaultPath, "utf8");
  assert.equal(storedEnvelope.includes(first.access_token), false);
  assert.equal(storedEnvelope.includes(first.refresh_token), false);

  vault.write(second);
  assert.deepEqual(vault.read(), second);

  const originalKey = readFileSync(keyPath, "utf8");
  makeVault(randomBytes(32));
  assert.throws(() => vault.read(), /authenticate|authentic|unsupported state/i);
  writeFileSync(keyPath, originalKey, { mode: 0o600 });
  chmodSync(keyPath, 0o600);
  assert.deepEqual(vault.read(), second);

  chmodSync(keyPath, 0o644);
  assert.throws(() => vault.read(), /group or other users/i);
  chmodSync(keyPath, 0o600);

  assert.equal(vault.remove(), true);
  assert.equal(vault.remove(), false);

  const symlinkTarget = join(root, "symlink-target.enc");
  writeFileSync(symlinkTarget, "{}", { mode: 0o600 });
  symlinkSync(symlinkTarget, vaultPath);
  assert.throws(() => vault.read(), /symbolic link/i);

  console.log(JSON.stringify({ ok: true, checks: 12 }));
} finally {
  rmSync(root, { recursive: true, force: true });
}
