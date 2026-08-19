import {
  createCipheriv,
  createDecipheriv,
  createECDH,
  createHash,
  createPrivateKey,
  createPublicKey,
  sign,
  verify,
  randomBytes
} from "node:crypto";
import { lstatSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { deflateRawSync, inflateRawSync } from "node:zlib";

const TOKEN_TYPE = "at+jwt";
const TOKEN_AAD = Buffer.from("onedrive-oauth-facade-upstream-token-v1", "utf8");
const ENCRYPTION_KEY_DOMAIN = Buffer.from("onedrive-oauth-facade-encryption-key-v1\0", "utf8");
const SIGNING_KEY_DOMAIN = Buffer.from("onedrive-oauth-facade-signing-key-v1\0", "utf8");
const MAX_TOKEN_BYTES = 65_536;
const CLOCK_SKEW_SECONDS = 60;

function ownerOnlyKey(file, name, domain = ENCRYPTION_KEY_DOMAIN) {
  const path = String(file || "").trim();
  if (!path || path === "/" || resolve(path) !== path) {
    throw new Error(`${name} must be a normalized absolute file path.`);
  }
  let encoded;
  try {
    const metadata = lstatSync(path);
    const currentUid = typeof process.getuid === "function"
      ? process.getuid()
      : metadata.uid;
    if (
      !metadata.isFile()
      || metadata.isSymbolicLink()
      || metadata.uid !== currentUid
      || metadata.size > 4096
      || (metadata.mode & 0o077) !== 0
    ) {
      throw new Error("invalid key file");
    }
    encoded = readFileSync(path, "utf8").trim();
  } catch {
    throw new Error(`${name} must be an owner-only regular file.`);
  }
  const source = /^[0-9a-f]{64}$/i.test(encoded)
    ? Buffer.from(encoded, "hex")
    : Buffer.from(encoded, "base64");
  if (source.length !== 32) {
    throw new Error(`${name} must decode to exactly 32 bytes.`);
  }
  return createHash("sha256").update(domain).update(source).digest();
}

function jwtClaims(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3 || parts.some((part) => !part) || parts[1].length > 16_384) {
    throw new Error("The Microsoft access token is not a bounded JWT.");
  }
  try {
    const claims = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
    if (!claims || typeof claims !== "object" || Array.isArray(claims)) {
      throw new Error("invalid claims");
    }
    return claims;
  } catch {
    throw new Error("The Microsoft access token has invalid claims.");
  }
}

function opaqueString(value, maximum = 32_768) {
  return typeof value === "string"
    && value.length > 0
    && value.length <= maximum
    && !/[\u0000-\u001f\u007f]/.test(value);
}

function signingKeys(file, name) {
  let privateValue = ownerOnlyKey(file, name, SIGNING_KEY_DOMAIN);
  const ecdh = createECDH("prime256v1");
  for (let attempt = 0; attempt < 16; attempt += 1) {
    try {
      ecdh.setPrivateKey(privateValue);
      break;
    } catch {
      privateValue = createHash("sha256")
        .update(SIGNING_KEY_DOMAIN)
        .update(privateValue)
        .update(Buffer.from([attempt + 1]))
        .digest();
    }
  }
  const publicValue = ecdh.getPublicKey(undefined, "uncompressed");
  if (publicValue.length !== 65 || publicValue[0] !== 4) {
    throw new Error(`${name} could not derive an ES256 signing key.`);
  }
  const x = publicValue.subarray(1, 33).toString("base64url");
  const y = publicValue.subarray(33, 65).toString("base64url");
  const d = privateValue.toString("base64url");
  const kid = createHash("sha256")
    .update(publicValue)
    .digest()
    .subarray(0, 16)
    .toString("base64url");
  const publicJwk = {
    kty: "EC",
    crv: "P-256",
    x,
    y,
    use: "sig",
    alg: "ES256",
    kid
  };
  const privateKey = createPrivateKey({
    key: { ...publicJwk, d },
    format: "jwk"
  });
  return {
    kid,
    privateKey,
    publicKey: createPublicKey(privateKey),
    publicJwk
  };
}

export function isFacadeAccessToken(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3 || !parts.every(Boolean) || parts[0].length > 1024) {
    return false;
  }
  try {
    const header = JSON.parse(Buffer.from(parts[0], "base64url").toString("utf8"));
    return header?.alg === "ES256" && header?.typ === TOKEN_TYPE;
  } catch {
    return false;
  }
}

export function validateFacadeAccessTokenKeyFile(file, name) {
  ownerOnlyKey(file, name);
  signingKeys(file, name);
  return String(file).trim();
}

export function facadeJwks(keyFile) {
  const { publicJwk } = signingKeys(
    keyFile,
    "ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_KEY_FILE"
  );
  return { keys: [publicJwk] };
}

export function issueFacadeAccessToken({
  providerAccessToken,
  issuer,
  audience,
  clientId,
  scope,
  expiresIn,
  keyFile,
  now = Date.now
}) {
  if (!opaqueString(providerAccessToken)) {
    throw new Error("A bounded Microsoft access token is required.");
  }
  const providerClaims = jwtClaims(providerAccessToken);
  const nowSeconds = Math.floor(now() / 1000);
  const providerExpiry = Number(providerClaims.exp);
  const requestedLifetime = Number(expiresIn);
  const lifetime = Number.isFinite(requestedLifetime) && requestedLifetime > 0
    ? Math.floor(requestedLifetime)
    : 300;
  const expiresAt = Math.min(
    Number.isFinite(providerExpiry) ? providerExpiry : nowSeconds + lifetime,
    nowSeconds + lifetime
  );
  if (expiresAt <= nowSeconds) {
    throw new Error("The Microsoft access token is already expired.");
  }
  const subject = String(providerClaims.oid || providerClaims.sub || "");
  const tenant = String(providerClaims.tid || "");
  if (!opaqueString(subject, 512) || !opaqueString(tenant, 128)) {
    throw new Error("The Microsoft access token is missing a stable subject or tenant.");
  }
  const payload = {
    v: 1,
    iss: String(issuer),
    aud: String(audience),
    azp: String(clientId).toLowerCase(),
    client_id: String(clientId).toLowerCase(),
    scp: String(scope),
    scope: String(scope),
    iat: nowSeconds,
    nbf: nowSeconds - 5,
    exp: expiresAt,
    tid: tenant,
    sub: subject
  };
  const encryptionKey = ownerOnlyKey(
    keyFile,
    "ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_KEY_FILE",
    ENCRYPTION_KEY_DOMAIN
  );
  const nonce = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", encryptionKey, nonce);
  cipher.setAAD(TOKEN_AAD);
  const compressedUpstream = deflateRawSync(
    Buffer.from(providerAccessToken, "utf8"),
    { level: 9 }
  );
  const ciphertext = Buffer.concat([
    cipher.update(compressedUpstream),
    cipher.final()
  ]);
  const tag = cipher.getAuthTag();
  payload.enc = [
    nonce.toString("base64url"),
    ciphertext.toString("base64url"),
    tag.toString("base64url")
  ].join(".");
  payload.zip = "DEF";
  const signingMaterial = signingKeys(
    keyFile,
    "ONEDRIVE_OAUTH_COMPAT_ACCESS_TOKEN_KEY_FILE"
  );
  const headerValue = Buffer.from(JSON.stringify({
    alg: "ES256",
    typ: TOKEN_TYPE,
    kid: signingMaterial.kid
  }), "utf8").toString("base64url");
  const payloadValue = Buffer.from(JSON.stringify(payload), "utf8").toString("base64url");
  const signingInput = `${headerValue}.${payloadValue}`;
  if (Buffer.byteLength(signingInput, "utf8") > MAX_TOKEN_BYTES) {
    throw new Error("The facade access token payload is too large.");
  }
  return {
    accessToken: `${signingInput}.${sign(
      "sha256",
      Buffer.from(signingInput, "utf8"),
      { key: signingMaterial.privateKey, dsaEncoding: "ieee-p1363" }
    ).toString("base64url")}`,
    expiresIn: expiresAt - nowSeconds
  };
}

export function verifyFacadeAccessToken(token, {
  issuer,
  audience,
  clientId,
  requiredScope,
  keyFile,
  now = Date.now
}) {
  const encoded = String(token || "");
  if (Buffer.byteLength(encoded, "utf8") > MAX_TOKEN_BYTES * 2) {
    throw new Error("The facade access token is too large.");
  }
  const [headerValue, payloadValue, signatureValue, ...extra] = encoded.split(".");
  if (
    extra.length
    || !headerValue
    || !payloadValue
    || !signatureValue
  ) {
    throw new Error("The facade access token is malformed.");
  }
  let header;
  let payload;
  let signature;
  try {
    header = JSON.parse(Buffer.from(headerValue, "base64url").toString("utf8"));
    payload = JSON.parse(Buffer.from(payloadValue, "base64url").toString("utf8"));
    signature = Buffer.from(signatureValue, "base64url");
  } catch {
    throw new Error("The facade access token is malformed.");
  }
  if (
    header?.alg !== "ES256"
    || header?.typ !== TOKEN_TYPE
    || !payload
    || typeof payload !== "object"
    || Array.isArray(payload)
  ) {
    throw new Error("The facade access token is malformed.");
  }
  const signingInput = `${headerValue}.${payloadValue}`;
  const signingMaterial = signingKeys(
    keyFile,
    "ONEDRIVE_MCP_OAUTH_FACADE_ACCESS_TOKEN_KEY_FILE"
  );
  if (
    header.kid !== signingMaterial.kid
    || signature.length !== 64
    || !verify(
      "sha256",
      Buffer.from(signingInput, "utf8"),
      { key: signingMaterial.publicKey, dsaEncoding: "ieee-p1363" },
      signature
    )
  ) {
    throw new Error("The facade access token is invalid.");
  }
  const nowSeconds = Math.floor(now() / 1000);
  const scopes = new Set(String(payload?.scp || "").split(/\s+/).filter(Boolean));
  const audiences = new Set(
    (Array.isArray(audience) ? audience : [audience])
      .map((value) => String(value || ""))
      .filter(Boolean)
  );
  if (
    payload?.v !== 1
    || payload.iss !== issuer
    || !audiences.has(payload.aud)
    || String(payload.azp || "").toLowerCase() !== String(clientId || "").toLowerCase()
    || String(payload.client_id || "").toLowerCase() !== String(clientId || "").toLowerCase()
    || payload.scope !== payload.scp
    || !Number.isFinite(payload.iat)
    || !Number.isFinite(payload.nbf)
    || !Number.isFinite(payload.exp)
    || payload.nbf > nowSeconds + CLOCK_SKEW_SECONDS
    || payload.exp <= nowSeconds - CLOCK_SKEW_SECONDS
    || !scopes.has(requiredScope)
    || !opaqueString(payload.tid, 128)
    || !opaqueString(payload.sub, 512)
    || payload.zip !== "DEF"
    || !opaqueString(payload.enc, MAX_TOKEN_BYTES)
  ) {
    throw new Error("The facade access token is invalid.");
  }
  const [nonceValue, ciphertextValue, tagValue, ...encryptedExtra] = payload.enc.split(".");
  let upstreamAccessToken;
  try {
    if (encryptedExtra.length || !nonceValue || !ciphertextValue || !tagValue) {
      throw new Error("malformed encrypted token");
    }
    const nonce = Buffer.from(nonceValue, "base64url");
    const ciphertext = Buffer.from(ciphertextValue, "base64url");
    const tag = Buffer.from(tagValue, "base64url");
    if (nonce.length !== 12 || tag.length !== 16 || !ciphertext.length) {
      throw new Error("malformed encrypted token");
    }
    const encryptionKey = ownerOnlyKey(
      keyFile,
      "ONEDRIVE_MCP_OAUTH_FACADE_ACCESS_TOKEN_KEY_FILE",
      ENCRYPTION_KEY_DOMAIN
    );
    const decipher = createDecipheriv("aes-256-gcm", encryptionKey, nonce);
    decipher.setAAD(TOKEN_AAD);
    decipher.setAuthTag(tag);
    const compressedUpstream = Buffer.concat([
      decipher.update(ciphertext),
      decipher.final()
    ]);
    upstreamAccessToken = inflateRawSync(compressedUpstream, {
      maxOutputLength: MAX_TOKEN_BYTES
    }).toString("utf8");
  } catch {
    throw new Error("The facade access token is invalid.");
  }
  if (!opaqueString(upstreamAccessToken)) {
    throw new Error("The facade access token is invalid.");
  }
  return {
    token: encoded,
    upstreamAccessToken,
    claims: {
      iss: payload.iss,
      aud: payload.aud,
      azp: payload.azp,
      client_id: payload.client_id,
      scp: payload.scp,
      scope: payload.scope,
      iat: payload.iat,
      nbf: payload.nbf,
      exp: payload.exp,
      tid: payload.tid,
      sub: payload.sub,
      ver: "facade-v1"
    }
  };
}
