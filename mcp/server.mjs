#!/usr/bin/env node

import { execFileSync } from "node:child_process";
import { createReadStream, readFileSync } from "node:fs";
import { mkdir, open, stat, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join, resolve } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const configPath = join(homedir(), ".codex", "onedrive-plugin", "config.json");
const downloadRoot = join(homedir(), ".codex", "onedrive-plugin", "downloads");
const localConfig = readLocalConfig();
const textFileLimit = 5 * 1024 * 1024;
const simpleUploadLimit = 250 * 1024 * 1024;
const uploadChunkUnit = 320 * 1024;
const defaultUploadChunkSize = 10 * 1024 * 1024;
const maxUploadChunkSize = 60 * 1024 * 1024;
const defaultSelect = "id,name,size,folder,file,webUrl,createdDateTime,lastModifiedDateTime,parentReference";
const textExtensions = new Set([
  ".bat", ".c", ".cfg", ".conf", ".cpp", ".cs", ".css", ".csv", ".env", ".go", ".h", ".hpp", ".htm",
  ".html", ".ini", ".java", ".js", ".json", ".jsx", ".log", ".md", ".mjs", ".php", ".properties",
  ".py", ".rb", ".rs", ".sh", ".sql", ".svg", ".toml", ".ts", ".tsx", ".txt", ".xml", ".yaml", ".yml"
]);
const textMimePrefixes = ["text/"];
const textMimeTypes = new Set([
  "application/csv",
  "application/ecmascript",
  "application/javascript",
  "application/json",
  "application/sql",
  "application/x-javascript",
  "application/x-ndjson",
  "application/xml",
  "image/svg+xml"
]);
const officeKinds = {
  excel: {
    label: "Excel",
    extensions: new Set([".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods"]),
    mimeTypes: new Set([
      "application/vnd.ms-excel",
      "application/vnd.ms-excel.sheet.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.oasis.opendocument.spreadsheet",
      "text/csv"
    ])
  },
  word: {
    label: "Word",
    extensions: new Set([".docx", ".doc", ".docm", ".rtf", ".odt"]),
    mimeTypes: new Set([
      "application/msword",
      "application/rtf",
      "application/vnd.ms-word.document.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/vnd.oasis.opendocument.text",
      "text/rtf"
    ])
  },
  powerpoint: {
    label: "PowerPoint",
    extensions: new Set([".pptx", ".ppt", ".pptm", ".ppsx", ".odp"]),
    mimeTypes: new Set([
      "application/vnd.ms-powerpoint",
      "application/vnd.ms-powerpoint.presentation.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
      "application/vnd.oasis.opendocument.presentation"
    ])
  }
};

let tokenCache = null;
let pendingDevice = null;

const outputFormatSchema = {
  type: "string",
  enum: ["compact", "full"],
  default: "compact",
  description: "compact returns chat-friendly item summaries; full returns richer item metadata."
};
const presetSchema = {
  type: "string",
  description: "Optional configured path preset, such as documents, desktop, pictures, or screenshots."
};
const relativePathSchema = {
  type: "string",
  description: "Path appended below the selected preset."
};
const pathTargetProperties = {
  path: { type: "string", description: "Item path relative to OneDrive root." },
  itemId: { type: "string", description: "Drive item ID." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const folderTargetProperties = {
  path: { type: "string", description: "Folder path relative to OneDrive root. Omit or use / for root." },
  itemId: { type: "string", description: "Drive item ID for the folder." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const destinationPresetProperties = {
  destinationParentPreset: presetSchema,
  destinationParentRelativePath: relativePathSchema
};
const parentPresetProperties = {
  parentPreset: presetSchema,
  parentRelativePath: relativePathSchema
};
const remotePresetProperties = {
  remotePreset: presetSchema,
  remoteRelativePath: {
    type: "string",
    description: "Destination path appended below the selected remotePreset, including filename."
  }
};
const itemTargetAnyOf = [
  { required: ["path"] },
  { required: ["itemId"] },
  { required: ["preset"] }
];
const remoteTargetAnyOf = [
  { required: ["remotePath"] },
  { required: ["remotePreset"] }
];

const tools = [
  {
    name: "onedrive_config",
    description: "Show OneDrive plugin configuration status without exposing secrets.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        checkToken: {
          type: "boolean",
          default: false,
          description: "When true, try to get an access token from Keychain or memory."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_auth_device_start",
    description: "Start Microsoft device-code login for OneDrive and return the user code and verification URL.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        tenant: { type: "string", description: "Optional tenant override. Defaults to configured tenant or common." },
        scopes: { type: "string", description: "Optional space-separated scope override." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_auth_device_poll",
    description: "Poll Microsoft token endpoint after the user completes device-code login, then store tokens in Keychain.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        deviceCode: { type: "string", description: "Optional device_code from onedrive_auth_device_start. Defaults to the latest pending code." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_logout",
    description: "Forget cached OneDrive tokens from memory and optionally delete the Keychain token.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        deleteKeychainToken: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_me",
    description: "Return the signed-in Microsoft profile.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_drive",
    description: "Return metadata for the signed-in user's default OneDrive.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_presets",
    description: "List configured OneDrive path presets and their root-relative paths.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_list",
    description: "List children in a OneDrive folder by path or item ID.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 100 },
        select: { type: "string", description: "Optional Graph $select fields." },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_list_all",
    description: "List all children in a OneDrive folder by following Microsoft Graph pagination up to a safe item cap.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...folderTargetProperties,
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        select: { type: "string", description: "Optional Graph $select fields." },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_search",
    description: "Search the signed-in user's OneDrive.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_search_all",
    description: "Search OneDrive and follow Microsoft Graph pagination up to a safe item cap.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_delta",
    description: "List recent changes from OneDrive delta sync, optionally continuing from a previous deltaLink or nextLink.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        deltaLink: { type: "string", description: "Previous @odata.deltaLink returned by onedrive_delta." },
        nextLink: { type: "string", description: "Previous @odata.nextLink returned by onedrive_delta." },
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_get_info",
    description: "Get metadata for a OneDrive item by path or item ID.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        includeDeletedItems: {
          type: "boolean",
          default: false,
          description: "When true and itemId is used, ask Graph to include deleted items. Supported by Microsoft for OneDrive Personal."
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_read_text",
    description: "Read a bounded text file from OneDrive by path or item ID.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        maxBytes: { type: "integer", minimum: 1, maximum: 10485760, default: textFileLimit },
        force: {
          type: "boolean",
          default: false,
          description: "When true, bypass MIME/extension text checks while keeping maxBytes enforcement."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download",
    description: "Download a OneDrive file to a local path.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/<filename>." },
        overwrite: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download_excel",
    description: "Download an Excel-like file (.xlsx, .xls, .xlsm, .xlsb, .csv, .ods) with extension/MIME checks.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/excel/<filename>." },
        overwrite: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download_word",
    description: "Download a Word-like document (.docx, .doc, .docm, .rtf, .odt) with extension/MIME checks.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/word/<filename>." },
        overwrite: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_download_powerpoint",
    description: "Download a PowerPoint-like file (.pptx, .ppt, .pptm, .ppsx, .odp) with extension/MIME checks.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/powerpoint/<filename>." },
        overwrite: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_upload",
    description: "Upload a local file to OneDrive. Uses simple upload for small files and upload sessions for large files.",
    inputSchema: {
      type: "object",
      required: ["localPath"],
      anyOf: remoteTargetAnyOf,
      properties: {
        localPath: { type: "string", description: "Absolute or workspace-relative local file path." },
        remotePath: { type: "string", description: "Destination path relative to OneDrive root, including filename." },
        ...remotePresetProperties,
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" },
        uploadMode: { type: "string", enum: ["auto", "simple", "session"], default: "auto" },
        chunkSize: {
          type: "integer",
          minimum: 327680,
          maximum: 62914560,
          default: defaultUploadChunkSize,
          description: "Upload-session chunk size. Must be a multiple of 320 KiB."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_write_text",
    description: "Create or replace a small text file in OneDrive.",
    inputSchema: {
      type: "object",
      required: ["content"],
      anyOf: remoteTargetAnyOf,
      properties: {
        remotePath: { type: "string", description: "Destination path relative to OneDrive root, including filename." },
        ...remotePresetProperties,
        content: { type: "string" },
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_create_folder",
    description: "Create a folder in OneDrive.",
    inputSchema: {
      type: "object",
      required: ["name"],
      properties: {
        parentPath: { type: "string", description: "Parent folder path relative to OneDrive root. Defaults to root." },
        parentItemId: { type: "string", description: "Parent folder drive item ID." },
        ...parentPresetProperties,
        name: { type: "string", minLength: 1 },
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_rename",
    description: "Rename a OneDrive item by path or item ID.",
    inputSchema: {
      type: "object",
      required: ["newName"],
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        newName: { type: "string", minLength: 1 },
        expectedName: { type: "string", description: "Optional safety check: item name must match before renaming." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match before renaming." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_move",
    description: "Move a OneDrive item to another folder, optionally renaming it. Refuses to move the OneDrive root.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        path: { type: "string", description: "Source item path relative to OneDrive root." },
        itemId: { type: "string", description: "Source drive item ID." },
        preset: presetSchema,
        relativePath: relativePathSchema,
        destinationParentPath: { type: "string", description: "Destination folder path relative to root. Omit or use / for root." },
        destinationParentItemId: { type: "string", description: "Destination folder drive item ID." },
        ...destinationPresetProperties,
        newName: { type: "string", description: "Optional new name after moving." },
        expectedName: { type: "string", description: "Optional safety check: source item name must match." },
        expectedId: { type: "string", description: "Optional safety check: source item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_copy",
    description: "Start a OneDrive copy operation to another folder. Copy is asynchronous; optionally wait for completion.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        path: { type: "string", description: "Source item path relative to OneDrive root." },
        itemId: { type: "string", description: "Source drive item ID." },
        preset: presetSchema,
        relativePath: relativePathSchema,
        destinationParentPath: { type: "string", description: "Destination folder path relative to root. Omit or use / for root." },
        destinationParentItemId: { type: "string", description: "Destination folder drive item ID." },
        ...destinationPresetProperties,
        newName: { type: "string", description: "Optional copied item name." },
        waitForCompletion: { type: "boolean", default: false },
        timeoutSeconds: { type: "integer", minimum: 1, maximum: 300, default: 60 },
        expectedName: { type: "string", description: "Optional safety check: source item name must match." },
        expectedId: { type: "string", description: "Optional safety check: source item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_create_sharing_link",
    description: "Create a OneDrive sharing link. Defaults to dry-run and requires confirmed=true plus dryRun=false to change sharing.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        type: { type: "string", enum: ["view", "edit", "embed"], default: "view" },
        scope: { type: "string", enum: ["anonymous", "organization", "users"], default: "anonymous" },
        retainInheritedPermissions: { type: "boolean" },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this can expose access to a file or folder."
        },
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_restore_deleted",
    description: "Restore a deleted OneDrive item by item ID. Defaults to dry-run and may require Files.ReadWrite.All on personal OneDrive.",
    inputSchema: {
      type: "object",
      required: ["itemId"],
      properties: {
        itemId: { type: "string", description: "Deleted drive item ID to restore." },
        destinationParentPath: { type: "string", description: "Optional restore destination folder path relative to root." },
        destinationParentItemId: { type: "string", description: "Optional restore destination folder drive item ID." },
        ...destinationPresetProperties,
        newName: { type: "string", description: "Optional new name when restoring." },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this restores a deleted item."
        },
        expectedId: { type: "string", description: "Optional safety check: itemId must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_permissions",
    description: "List sharing and permission grants on a OneDrive file or folder.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_delete",
    description: "Delete a OneDrive item by path or item ID. Defaults to dry-run.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this deletes a file or folder."
        },
        expectedName: { type: "string", description: "Optional safety check: item name must match before deleting." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match before deleting." }
      },
      additionalProperties: false
    }
  }
];

function readLocalConfig() {
  try {
    return JSON.parse(readFileSync(configPath, "utf8"));
  } catch {
    return {};
  }
}

function config(overrides = {}) {
  const clientId = overrides.clientId || process.env.ONEDRIVE_CLIENT_ID || localConfig.clientId;
  const tenant = overrides.tenant || process.env.ONEDRIVE_TENANT || localConfig.tenant || "common";
  const scopes = overrides.scopes || process.env.ONEDRIVE_SCOPES || localConfig.scopes || "offline_access User.Read Files.ReadWrite";
  const keychainService = process.env.ONEDRIVE_KEYCHAIN_SERVICE || localConfig.keychainService || "Codex OneDrive";
  return { clientId, tenant, scopes, keychainService };
}

function pathPresets() {
  return {
    documents: "Documents",
    desktop: "Desktop",
    pictures: "Pictures",
    screenshots: "Pictures/Screenshots",
    ...(localConfig.pathPresets || {})
  };
}

function requireClientId(cfg = config()) {
  if (!cfg.clientId) {
    throw new Error("ONEDRIVE_CLIENT_ID is not configured. Run scripts/configure.zsh from the OneDrive plugin directory, or set ONEDRIVE_CLIENT_ID.");
  }
}

function tokenUrl(tenant) {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
}

function deviceCodeUrl(tenant) {
  return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/devicecode`;
}

function keychainAccount() {
  return "tokens";
}

function getKeychainToken(cfg = config()) {
  try {
    const raw = execFileSync("security", [
      "find-generic-password",
      "-a", keychainAccount(),
      "-s", cfg.keychainService,
      "-w"
    ], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim();
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function setKeychainToken(token, cfg = config()) {
  const payload = JSON.stringify(token);
  try {
    execFileSync("security", [
      "add-generic-password",
      "-U",
      "-a", keychainAccount(),
      "-s", cfg.keychainService,
      "-w", payload
    ], { stdio: "ignore" });
  } catch (error) {
    throw new Error(`Could not store OneDrive token in Keychain: ${error.message}`);
  }
}

function deleteKeychainToken(cfg = config()) {
  try {
    execFileSync("security", [
      "delete-generic-password",
      "-a", keychainAccount(),
      "-s", cfg.keychainService
    ], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

function publicConfig() {
  const cfg = config();
  const stored = getKeychainToken(cfg);
  return {
    clientIdConfigured: Boolean(cfg.clientId),
    tenant: cfg.tenant,
    scopes: cfg.scopes,
    keychainService: cfg.keychainService,
    keychainTokenConfigured: Boolean(stored?.refresh_token),
    configPath,
    pathPresets: pathPresets()
  };
}

function normalizeToken(raw) {
  const now = Date.now();
  return {
    ...raw,
    expires_at: raw.expires_at || (raw.expires_in ? now + raw.expires_in * 1000 : null),
    obtained_at: raw.obtained_at || now
  };
}

async function postForm(url, params) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params)
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = body.error_description || body.error || `${response.status} ${response.statusText}`;
    const error = new Error(message);
    error.body = body;
    throw error;
  }
  return body;
}

async function startDeviceLogin(args = {}) {
  const cfg = config(args);
  requireClientId(cfg);
  const result = await postForm(deviceCodeUrl(cfg.tenant), {
    client_id: cfg.clientId,
    scope: cfg.scopes
  });
  pendingDevice = { ...result, tenant: cfg.tenant, scopes: cfg.scopes, startedAt: Date.now() };
  return {
    userCode: result.user_code,
    verificationUri: result.verification_uri,
    verificationUriComplete: result.verification_uri_complete,
    expiresIn: result.expires_in,
    interval: result.interval,
    message: result.message,
    deviceCodeStoredInMemory: true
  };
}

async function pollDeviceLogin(args = {}) {
  const deviceCode = args.deviceCode || pendingDevice?.device_code;
  if (!deviceCode) throw new Error("No pending device code. Run onedrive_auth_device_start first.");
  const cfg = config({ tenant: pendingDevice?.tenant, scopes: pendingDevice?.scopes });
  requireClientId(cfg);
  const result = await postForm(tokenUrl(cfg.tenant), {
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    client_id: cfg.clientId,
    device_code: deviceCode
  }).catch((error) => {
    if (error.body?.error === "authorization_pending") {
      return { authorizationPending: true, message: "Authorization is still pending. Try again after the user completes browser sign-in." };
    }
    if (error.body?.error === "slow_down") {
      return { authorizationPending: true, slowDown: true, message: "Microsoft asked polling to slow down. Try again in a few more seconds." };
    }
    throw error;
  });
  if (result.authorizationPending) return result;
  tokenCache = normalizeToken(result);
  setKeychainToken(tokenCache, cfg);
  pendingDevice = null;
  return {
    authenticated: true,
    tokenType: tokenCache.token_type,
    expiresAt: tokenCache.expires_at ? new Date(tokenCache.expires_at).toISOString() : null,
    refreshTokenStoredInKeychain: Boolean(tokenCache.refresh_token)
  };
}

async function refreshAccessToken(refreshToken, cfg = config()) {
  requireClientId(cfg);
  const result = await postForm(tokenUrl(cfg.tenant), {
    grant_type: "refresh_token",
    client_id: cfg.clientId,
    refresh_token: refreshToken,
    scope: cfg.scopes
  });
  tokenCache = normalizeToken({ ...result, refresh_token: result.refresh_token || refreshToken });
  setKeychainToken(tokenCache, cfg);
  return tokenCache;
}

async function getAccessToken() {
  if (process.env.ONEDRIVE_TEST_ACCESS_TOKEN) return process.env.ONEDRIVE_TEST_ACCESS_TOKEN;
  const cfg = config();
  requireClientId(cfg);
  const current = tokenCache || getKeychainToken(cfg);
  if (!current?.refresh_token && !current?.access_token) {
    throw new Error("OneDrive is not authenticated. Run onedrive_auth_device_start, complete browser login, then run onedrive_auth_device_poll.");
  }
  const expiresAt = current.expires_at || 0;
  if (current.access_token && expiresAt - Date.now() > 60_000) {
    tokenCache = current;
    return current.access_token;
  }
  if (!current.refresh_token) throw new Error("Stored token has no refresh token. Run device-code login again.");
  const refreshed = await refreshAccessToken(current.refresh_token, cfg);
  return refreshed.access_token;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function retryDelayMs(response, attempt) {
  const retryAfter = response.headers.get("retry-after");
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const dateMs = Date.parse(retryAfter);
    if (Number.isFinite(dateMs)) return Math.max(0, dateMs - Date.now());
  }
  return Math.min(1000 * 2 ** attempt, 8000);
}

function shouldRetryResponse(response) {
  return response.status === 429 || response.status === 503 || response.status === 504 || response.status === 502 || response.status === 500;
}

function isReplayableBody(body) {
  return body === undefined
    || body === null
    || typeof body === "string"
    || Buffer.isBuffer(body)
    || body instanceof ArrayBuffer
    || body instanceof URLSearchParams;
}

function retryCountForRequest(fetchOptions, explicitMaxRetries) {
  if (explicitMaxRetries === 0) return 0;
  const method = String(fetchOptions.method || "GET").toUpperCase();
  if (method === "GET" || method === "HEAD") return explicitMaxRetries ?? 3;
  if (explicitMaxRetries !== undefined && isReplayableBody(fetchOptions.body)) return explicitMaxRetries;
  return 0;
}

function shouldDefaultJsonContentType(body) {
  return body
    && typeof body === "string"
    && body.trim().startsWith("{");
}

async function parseResponseBody(response) {
  if (response.status === 204) return null;
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) return await response.json().catch(() => ({}));
  return await response.arrayBuffer();
}

async function fetchWithRetry(url, options = {}, retryOptions = {}) {
  const maxRetries = retryOptions.maxRetries ?? 3;
  for (let attempt = 0; ; attempt += 1) {
    const response = await fetch(url, options);
    if (!shouldRetryResponse(response) || attempt >= maxRetries) return response;
    await parseResponseBody(response).catch(() => null);
    await sleep(retryDelayMs(response, attempt));
  }
}

function microsoftGraphError(body, response) {
  const code = body?.error?.code ? `${body.error.code}: ` : "";
  const message = body?.error?.message || `${response.status} ${response.statusText}`;
  const requestId = body?.error?.innerError?.["request-id"] || body?.error?.innerError?.requestId;
  const suffix = requestId ? ` (request-id: ${requestId})` : "";
  return new Error(`Microsoft Graph error: ${code}${message}${suffix}`);
}

async function graph(path, options = {}) {
  const { returnResponse = false, maxRetries, skipAuth = false, ...fetchOptions } = options;
  const accessToken = skipAuth ? null : await getAccessToken();
  const graphBaseUrl = process.env.ONEDRIVE_GRAPH_BASE_URL || "https://graph.microsoft.com/v1.0";
  const url = path.startsWith("http") ? path : `${graphBaseUrl}${path}`;
  const headers = {
    ...(accessToken ? { Authorization: `Bearer ${accessToken}` } : {}),
    ...(shouldDefaultJsonContentType(fetchOptions.body) ? { "Content-Type": "application/json" } : {}),
    ...(fetchOptions.headers || {})
  };
  const retriedResponse = await fetchWithRetry(url, {
    ...fetchOptions,
    headers
  }, { maxRetries: retryCountForRequest(fetchOptions, maxRetries) });
  const body = await parseResponseBody(retriedResponse);
  if (returnResponse) {
    return { body, headers: retriedResponse.headers, status: retriedResponse.status, ok: retriedResponse.ok };
  }
  if (!retriedResponse.ok) {
    throw microsoftGraphError(body, retriedResponse);
  }
  return body;
}

function cleanPath(path = "") {
  return String(path).replace(/^\/+/, "").replace(/\/+$/, "");
}

function assertSafeRemotePath(path = "", label = "path") {
  const clean = cleanPath(path);
  for (const segment of clean.split("/").filter(Boolean)) {
    if (segment === "." || segment === ".." || segment.includes("\\")) {
      throw new Error(`${label} contains an unsafe path segment: ${segment}`);
    }
  }
  return clean;
}

function pathName(path = "") {
  return basename(cleanPath(path));
}

function encodeDrivePath(path) {
  return assertSafeRemotePath(path).split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

function joinRemotePath(...parts) {
  return assertSafeRemotePath(parts.filter((part) => part !== undefined && part !== null && String(part).trim() !== "").join("/"));
}

function resolvePresetPath(args = {}, options = {}) {
  const {
    pathField = "path",
    presetField = "preset",
    relativeField = "relativePath",
    allowEmpty = true
  } = options;
  const explicitPath = args[pathField];
  const preset = args[presetField];
  if (preset) {
    const presets = pathPresets();
    if (!Object.hasOwn(presets, preset)) {
      throw new Error(`Unknown OneDrive path preset: ${preset}. Available presets: ${Object.keys(presets).join(", ")}`);
    }
    return joinRemotePath(presets[preset], args[relativeField] || "");
  }
  const resolved = assertSafeRemotePath(explicitPath || "", pathField);
  if (!allowEmpty && !resolved) throw new Error(`${pathField} or ${presetField} is required.`);
  return resolved;
}

function itemArgsWithResolvedPath(args = {}) {
  if (args.itemId) return args;
  return { ...args, path: resolvePresetPath(args) };
}

function remotePath(args = {}) {
  return resolvePresetPath(args, {
    pathField: "remotePath",
    presetField: "remotePreset",
    relativeField: "remoteRelativePath",
    allowEmpty: false
  });
}

function formatDriveItem(item, format = "compact") {
  const simplified = simplifyItem(item);
  if (!simplified || format === "full") return simplified;
  return {
    id: simplified.id,
    name: simplified.name,
    type: simplified.folder ? "folder" : simplified.file ? "file" : "item",
    size: simplified.size,
    lastModifiedDateTime: simplified.lastModifiedDateTime,
    webUrl: simplified.webUrl
  };
}

function itemBase(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (resolved.itemId) return `/me/drive/items/${encodeURIComponent(resolved.itemId)}`;
  if (!resolved.path || cleanPath(resolved.path) === "") return "/me/drive/root";
  return `/me/drive/root:/${encodeDrivePath(resolved.path)}:`;
}

function requireNonRootTarget(args = {}, operation) {
  const resolved = itemArgsWithResolvedPath(args);
  if (!resolved.itemId && (!resolved.path || cleanPath(resolved.path) === "")) {
    throw new Error(`${operation} requires a non-root path or itemId. Refusing to operate on the OneDrive root.`);
  }
}

function assertExpectedItem(rawItem, args = {}, operation) {
  if (args.expectedId && rawItem.id !== args.expectedId) {
    throw new Error(`${operation} expected item ID ${args.expectedId}, but resolved ${rawItem.id}. Refusing to continue.`);
  }
  if (args.expectedName && rawItem.name !== args.expectedName) {
    throw new Error(`${operation} expected item named ${args.expectedName}, but resolved ${rawItem.name}. Refusing to continue.`);
  }
}

function childrenPath(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (resolved.itemId) return `/me/drive/items/${encodeURIComponent(resolved.itemId)}/children`;
  if (!resolved.path || cleanPath(resolved.path) === "") return "/me/drive/root/children";
  return `/me/drive/root:/${encodeDrivePath(resolved.path)}:/children`;
}

function contentPath(args = {}) {
  return `${itemBase(args)}/content`;
}

function uploadPath(remotePath, conflictBehavior = "fail") {
  const behavior = encodeURIComponent(conflictBehavior);
  return `/me/drive/root:/${encodeDrivePath(remotePath)}:/content?@microsoft.graph.conflictBehavior=${behavior}`;
}

function splitRemotePath(remotePath) {
  const clean = assertSafeRemotePath(remotePath, "remotePath");
  if (!clean) throw new Error("remotePath must include a filename.");
  const parts = clean.split("/").filter(Boolean);
  const name = parts.pop();
  return { parentPath: parts.join("/"), name };
}

async function uploadSessionTarget(remotePath) {
  const { parentPath, name } = splitRemotePath(remotePath);
  const parent = await resolveDestinationParent({ destinationParentPath: parentPath });
  return {
    endpoints: [
      `/me/drive/items/${parent.id}:/${encodeURIComponent(name)}:/createUploadSession`,
      `/me/drive/root:/${encodeDrivePath(remotePath)}:/createUploadSession`
    ],
    name
  };
}

function simplifyItem(item) {
  if (!item) return item;
  return {
    id: item.id,
    name: item.name,
    path: item.parentReference?.path,
    webUrl: item.webUrl,
    size: item.size,
    createdDateTime: item.createdDateTime,
    lastModifiedDateTime: item.lastModifiedDateTime,
    folder: item.folder ? { childCount: item.folder.childCount } : undefined,
    file: item.file ? { mimeType: item.file.mimeType, hashes: item.file.hashes } : undefined
  };
}

async function list(args = {}) {
  const params = new URLSearchParams();
  params.set("$top", String(args.limit ?? 100));
  params.set("$select", args.select || defaultSelect);
  const result = await graph(`${childrenPath(args)}?${params.toString()}`);
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), nextLink: result["@odata.nextLink"] || null };
}

async function collectPages(firstPath, maxItems, format = "compact", formatter = formatDriveItem) {
  const items = [];
  let nextPath = firstPath;
  let nextLink = null;
  let deltaLink = null;
  while (nextPath && items.length < maxItems) {
    const result = await graph(nextPath);
    const pageItems = result.value || [];
    const remaining = maxItems - items.length;
    items.push(...pageItems.slice(0, remaining).map((item) => formatter(item, format)));
    nextLink = result["@odata.nextLink"] || null;
    deltaLink = result["@odata.deltaLink"] || null;
    nextPath = nextLink && items.length < maxItems ? nextLink : null;
  }
  return { items, nextLink, deltaLink, truncated: Boolean(nextLink), count: items.length };
}

async function listAll(args = {}) {
  const maxItems = Math.min(args.maxItems ?? 1000, 5000);
  const params = new URLSearchParams();
  params.set("$top", String(args.pageSize ?? 200));
  params.set("$select", args.select || defaultSelect);
  return await collectPages(`${childrenPath(args)}?${params.toString()}`, maxItems, args.format);
}

async function search(args = {}) {
  const escaped = String(args.query).replace(/'/g, "''");
  const params = new URLSearchParams();
  params.set("$top", String(args.limit ?? 50));
  const result = await graph(`/me/drive/root/search(q='${encodeURIComponent(escaped)}')?${params.toString()}`);
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), nextLink: result["@odata.nextLink"] || null };
}

async function searchAll(args = {}) {
  const escaped = String(args.query).replace(/'/g, "''");
  const maxItems = Math.min(args.maxItems ?? 1000, 5000);
  const params = new URLSearchParams();
  params.set("$top", String(args.pageSize ?? 200));
  return await collectPages(`/me/drive/root/search(q='${encodeURIComponent(escaped)}')?${params.toString()}`, maxItems, args.format);
}

function formatDeltaItem(item, format = "compact") {
  const formatted = formatDriveItem(item, format);
  if (format === "full") return formatted ? { ...formatted, deleted: item.deleted } : formatted;
  return formatted ? { ...formatted, deleted: item.deleted ? item.deleted : undefined } : formatted;
}

async function delta(args = {}) {
  const maxItems = Math.min(args.maxItems ?? 1000, 5000);
  let firstPath = args.nextLink || args.deltaLink;
  let target = args.nextLink ? "nextLink" : args.deltaLink ? "deltaLink" : "root";
  if (!firstPath) {
    const params = new URLSearchParams();
    params.set("$top", String(args.pageSize ?? 200));
    if (args.itemId) {
      firstPath = `/me/drive/items/${encodeURIComponent(args.itemId)}/delta?${params.toString()}`;
      target = "itemId";
    } else {
      const resolvedPath = resolvePresetPath(args);
      if (resolvedPath) {
        const folder = await getRawInfo({ path: resolvedPath });
        if (!folder.folder && !folder.root) throw new Error(`Delta target is not a folder: ${folder.name}`);
        firstPath = `/me/drive/items/${encodeURIComponent(folder.id)}/delta?${params.toString()}`;
        target = resolvedPath;
      } else {
        firstPath = `/me/drive/root/delta?${params.toString()}`;
      }
    }
  }
  const result = await collectPages(firstPath, maxItems, args.format, formatDeltaItem);
  return {
    ...result,
    target,
    note: result.deltaLink
      ? "Save deltaLink to ask for changes since this point later."
      : "Use nextLink to continue this delta scan before saving a deltaLink."
  };
}

async function getRawInfo(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (!resolved.path && !resolved.itemId) throw new Error("Provide path, preset, or itemId.");
  const suffix = args.includeDeletedItems && resolved.itemId ? "?includeDeletedItems=true" : "";
  return await graph(`${itemBase(args)}${suffix}`);
}

async function getInfo(args = {}) {
  return formatDriveItem(await getRawInfo(args), args.format || "full");
}

function compactIdentity(identity) {
  if (!identity) return undefined;
  const user = identity.user || identity.application || identity.device;
  return user ? { id: user.id, displayName: user.displayName, email: user.email } : identity;
}

function simplifyPermission(permission, format = "compact") {
  if (format === "full") return permission;
  return {
    id: permission.id,
    roles: permission.roles,
    link: permission.link ? {
      type: permission.link.type,
      scope: permission.link.scope,
      webUrl: permission.link.webUrl,
      preventsDownload: permission.link.preventsDownload
    } : undefined,
    grantedTo: compactIdentity(permission.grantedTo),
    grantedToV2: compactIdentity(permission.grantedToV2),
    grantedToIdentities: permission.grantedToIdentities?.map(compactIdentity),
    invitation: permission.invitation ? {
      email: permission.invitation.email,
      signInRequired: permission.invitation.signInRequired
    } : undefined,
    inheritedFrom: permission.inheritedFrom ? simplifyItem(permission.inheritedFrom) : undefined,
    expirationDateTime: permission.expirationDateTime,
    hasPassword: permission.hasPassword
  };
}

async function permissions(args = {}) {
  const result = await graph(`${itemBase(args)}/permissions`);
  return {
    item: await getInfo({ ...args, format: args.format || "compact" }).catch(() => null),
    permissions: (result.value || []).map((permission) => simplifyPermission(permission, args.format)),
    count: (result.value || []).length
  };
}

function isTextMimeType(mimeType = "") {
  const lower = mimeType.toLowerCase();
  return textMimeTypes.has(lower) || textMimePrefixes.some((prefix) => lower.startsWith(prefix));
}

function isLikelyTextItem(info, args = {}) {
  const mimeType = info.file?.mimeType || "";
  if (mimeType && isTextMimeType(mimeType)) return true;
  const name = info.name || pathName(args.path || "");
  return textExtensions.has(extname(name).toLowerCase());
}

function assertTextReadable(info, args = {}) {
  if (info.folder) throw new Error(`Item is a folder, not a text file: ${info.name}`);
  if (args.force === true) return;
  if (!isLikelyTextItem(info, args)) {
    const mime = info.file?.mimeType || "unknown MIME type";
    throw new Error(`Refusing to read likely binary file as text (${mime}). Use onedrive_download, or pass force: true with maxBytes if you are sure.`);
  }
}

function assertNoBinaryNulls(buffer) {
  const sampleLength = Math.min(buffer.length, 4096);
  for (let index = 0; index < sampleLength; index += 1) {
    if (buffer[index] === 0) {
      throw new Error("Downloaded content contains NUL bytes and looks binary. Use onedrive_download instead.");
    }
  }
}

async function resolveDestinationParent(args = {}) {
  if (args.destinationParentItemId) return { id: args.destinationParentItemId };
  if (args.parentItemId) return { id: args.parentItemId };
  let parentPath = "";
  if (args.destinationParentPreset) {
    parentPath = resolvePresetPath(args, {
      pathField: "destinationParentPath",
      presetField: "destinationParentPreset",
      relativeField: "destinationParentRelativePath"
    });
  } else if (args.parentPreset) {
    parentPath = resolvePresetPath(args, {
      pathField: "parentPath",
      presetField: "parentPreset",
      relativeField: "parentRelativePath"
    });
  } else {
    parentPath = args.destinationParentPath ?? args.parentPath ?? "";
  }
  const raw = cleanPath(parentPath) === "" ? await graph("/me/drive/root") : await getRawInfo({ path: parentPath });
  if (!raw.folder && !raw.root) throw new Error(`Destination parent is not a folder: ${raw.name}`);
  return { id: raw.id, path: raw.parentReference?.path, name: raw.name };
}

async function readText(args = {}) {
  const info = await getInfo(args);
  const maxBytes = args.maxBytes ?? textFileLimit;
  assertTextReadable(info, args);
  if (info.size && info.size > maxBytes) {
    throw new Error(`File is ${info.size} bytes, above maxBytes ${maxBytes}. Use onedrive_download instead.`);
  }
  const buffer = Buffer.from(await graph(contentPath(args)));
  if (buffer.length > maxBytes) {
    throw new Error(`Downloaded content is ${buffer.length} bytes, above maxBytes ${maxBytes}.`);
  }
  if (args.force !== true) assertNoBinaryNulls(buffer);
  return { item: info, content: buffer.toString("utf8") };
}

async function download(args = {}) {
  const info = await getInfo(args);
  const target = args.localPath ? resolve(args.localPath) : join(downloadRoot, info.name || basename(cleanPath(args.path || args.itemId || "download")));
  if (args.overwrite !== true) {
    try {
      await stat(target);
      throw new Error(`Local file already exists: ${target}. Pass overwrite: true to replace it.`);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  }
  await mkdir(dirname(target), { recursive: true });
  const buffer = Buffer.from(await graph(contentPath(args)));
  await writeFile(target, buffer);
  return { item: info, localPath: target, bytesWritten: buffer.length };
}

function assertOfficeKind(info, kindName) {
  const kind = officeKinds[kindName];
  if (!kind) throw new Error(`Unknown Office helper kind: ${kindName}`);
  if (info.folder) throw new Error(`Item is a folder, not a ${kind.label} file: ${info.name}`);
  const extension = extname(info.name || "").toLowerCase();
  const mimeType = (info.file?.mimeType || "").toLowerCase();
  if (!kind.extensions.has(extension) && !kind.mimeTypes.has(mimeType)) {
    throw new Error(`Expected a ${kind.label}-compatible file. Got ${info.name || "unnamed item"} (${mimeType || "unknown MIME type"}).`);
  }
}

async function downloadOffice(args = {}, kindName) {
  const info = await getInfo(args);
  assertOfficeKind(info, kindName);
  return await download({
    ...args,
    localPath: args.localPath || join(downloadRoot, kindName, info.name || `${kindName}-download`),
    overwrite: args.overwrite
  });
}

async function upload(args = {}) {
  const localPath = resolve(args.localPath);
  const destinationPath = remotePath(args);
  const fileStat = await stat(localPath);
  if (!fileStat.isFile()) throw new Error(`Not a file: ${localPath}`);
  const uploadMode = args.uploadMode || "auto";
  if (uploadMode === "session" || (uploadMode === "auto" && fileStat.size > simpleUploadLimit)) {
    return await uploadLarge({ ...args, localPath, remotePath: destinationPath }, fileStat);
  }
  if (uploadMode === "simple" && fileStat.size > simpleUploadLimit) {
    throw new Error(`Simple upload only supports files up to ${simpleUploadLimit} bytes. Use uploadMode: "session" or "auto".`);
  }
  const stream = createReadStream(localPath);
  const result = await graph(uploadPath(destinationPath, args.conflictBehavior || "fail"), {
    method: "PUT",
    body: stream,
    duplex: "half",
    headers: { "Content-Type": "application/octet-stream" }
  });
  return { item: simplifyItem(result), localPath, bytesUploaded: fileStat.size, uploadMode: "simple" };
}

function normalizeChunkSize(value = defaultUploadChunkSize) {
  const chunkSize = Number(value);
  if (!Number.isInteger(chunkSize) || chunkSize < uploadChunkUnit || chunkSize > maxUploadChunkSize) {
    throw new Error(`chunkSize must be an integer between ${uploadChunkUnit} and ${maxUploadChunkSize} bytes.`);
  }
  if (chunkSize % uploadChunkUnit !== 0) {
    throw new Error(`chunkSize must be a multiple of ${uploadChunkUnit} bytes.`);
  }
  return chunkSize;
}

async function uploadLarge(args = {}, fileStat) {
  const chunkSize = normalizeChunkSize(args.chunkSize ?? defaultUploadChunkSize);
  const sessionTarget = await uploadSessionTarget(args.remotePath);
  const session = await createUploadSession(sessionTarget, args.conflictBehavior || "fail");
  if (!session.uploadUrl) throw new Error("Microsoft Graph did not return an uploadUrl.");

  let position = 0;
  let uploaded = 0;
  let finalItem = null;
  const handle = await open(args.localPath, "r");
  try {
    while (position < fileStat.size) {
      const length = Math.min(chunkSize, fileStat.size - position);
      const buffer = Buffer.allocUnsafe(length);
      const { bytesRead } = await handle.read(buffer, 0, length, position);
      if (bytesRead <= 0) throw new Error(`Could not read upload chunk at byte ${position}.`);
      const body = bytesRead === buffer.length ? buffer : buffer.subarray(0, bytesRead);
      const end = position + bytesRead - 1;
      let response;
      try {
        response = await graph(session.uploadUrl, {
          method: "PUT",
          skipAuth: true,
          body,
          headers: {
            "Content-Range": `bytes ${position}-${end}/${fileStat.size}`
          },
          maxRetries: 4
        });
      } catch (error) {
        throw new Error(`Upload session failed for byte range ${position}-${end}/${fileStat.size}: ${error.message}`);
      }
      uploaded = end + 1;
      if (response?.id) finalItem = response;
      position += bytesRead;
    }
  } finally {
    await handle.close();
  }

  return {
    item: simplifyItem(finalItem),
    localPath: args.localPath,
    bytesUploaded: uploaded,
    uploadMode: "session",
    chunkSize
  };
}

async function createUploadSession(sessionTarget, conflictBehavior) {
  const bodies = [
    {
      item: {
        "@odata.type": "microsoft.graph.driveItemUploadableProperties",
        "@microsoft.graph.conflictBehavior": conflictBehavior,
        name: sessionTarget.name
      }
    },
    {
      item: {
        "@microsoft.graph.conflictBehavior": conflictBehavior,
        name: sessionTarget.name
      }
    },
    null
  ];
  const errors = [];
  for (const endpoint of sessionTarget.endpoints) {
    for (const body of bodies) {
      try {
        return await graph(endpoint, {
          method: "POST",
          ...(body ? { body: JSON.stringify(body) } : {})
        });
      } catch (error) {
        errors.push(`${endpoint}: ${error.message}`);
      }
    }
  }
  throw new Error(`Could not create upload session. Tried ${errors.length} compatible request shapes. Last error: ${errors.at(-1)}`);
}

async function writeText(args = {}) {
  const destinationPath = remotePath(args);
  const result = await graph(uploadPath(destinationPath, args.conflictBehavior || "fail"), {
    method: "PUT",
    body: Buffer.from(args.content, "utf8"),
    headers: { "Content-Type": "text/plain; charset=utf-8" }
  });
  return { item: simplifyItem(result), bytesUploaded: Buffer.byteLength(args.content, "utf8") };
}

async function createFolder(args = {}) {
  const endpoint = args.parentItemId
    ? `/me/drive/items/${encodeURIComponent(args.parentItemId)}/children`
    : childrenPath({
        path: args.parentPath || "",
        preset: args.parentPreset,
        relativePath: args.parentRelativePath
      });
  const result = await graph(endpoint, {
    method: "POST",
    body: JSON.stringify({
      name: args.name,
      folder: {},
      "@microsoft.graph.conflictBehavior": args.conflictBehavior || "fail"
    })
  });
  return simplifyItem(result);
}

async function rename(args = {}) {
  requireNonRootTarget(args, "Rename");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Rename refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Rename");
  const result = await graph(itemBase(args), {
    method: "PATCH",
    body: JSON.stringify({ name: args.newName })
  });
  return simplifyItem(result);
}

async function moveItem(args = {}) {
  requireNonRootTarget(args, "Move");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Move refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Move");
  const parentReference = await resolveDestinationParent(args);
  const body = { parentReference: { id: parentReference.id } };
  if (args.newName) body.name = args.newName;
  const result = await graph(itemBase(args), {
    method: "PATCH",
    body: JSON.stringify(body)
  });
  return simplifyItem(result);
}

async function pollCopyMonitor(monitorUrl, timeoutSeconds = 60) {
  const deadline = Date.now() + timeoutSeconds * 1000;
  let last = null;
  while (Date.now() < deadline) {
    const response = await graph(monitorUrl, { skipAuth: true, returnResponse: true, maxRetries: 3, redirect: "manual" });
    last = response.body && !(response.body instanceof ArrayBuffer) ? response.body : null;
    if (response.status === 303) {
      return {
        complete: true,
        status: response.status,
        resourceLocation: response.headers.get("location"),
        monitorUrl
      };
    }
    if (response.ok && last?.status && !["notStarted", "running", "inProgress"].includes(String(last.status))) {
      return { complete: true, status: response.status, monitorUrl, monitor: last };
    }
    await sleep(2000);
  }
  return { complete: false, timeoutSeconds, monitorUrl, monitor: last };
}

async function copyItem(args = {}) {
  requireNonRootTarget(args, "Copy");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Copy refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Copy");
  const parentReference = await resolveDestinationParent(args);
  const response = await graph(`${itemBase(args)}/copy`, {
    method: "POST",
    returnResponse: true,
    body: JSON.stringify({
      parentReference: { id: parentReference.id },
      ...(args.newName ? { name: args.newName } : {})
    })
  });
  if (!response.ok) throw microsoftGraphError(response.body, { status: response.status, statusText: "Copy failed" });
  const monitorUrl = response.headers.get("location");
  const result = {
    accepted: response.status === 202 || response.ok,
    status: response.status,
    source: simplifyItem(current),
    monitorUrl
  };
  if (args.waitForCompletion && monitorUrl) {
    result.monitor = await pollCopyMonitor(monitorUrl, args.timeoutSeconds ?? 60);
  }
  return result;
}

async function createSharingLink(args = {}) {
  requireNonRootTarget(args, "Create sharing link");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Create sharing link refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Create sharing link");
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldCreate: {
      item: simplifyItem(current),
      type: args.type || "view",
      scope: args.scope || "anonymous",
      retainInheritedPermissions: args.retainInheritedPermissions
    }
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...preview,
      requiredToCreate: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  const body = {
    type: args.type || "view",
    scope: args.scope || "anonymous"
  };
  if (typeof args.retainInheritedPermissions === "boolean") {
    body.retainInheritedPermissions = args.retainInheritedPermissions;
  }
  const result = await graph(`${itemBase(args)}/createLink`, {
    method: "POST",
    body: JSON.stringify(body)
  });
  return {
    dryRun: false,
    confirmed: true,
    item: simplifyItem(current),
    permission: result
  };
}

async function deleteItem(args = {}) {
  requireNonRootTarget(args, "Delete");
  const rawItem = await getRawInfo(args);
  if (rawItem.root) throw new Error("Delete refuses to operate on the OneDrive root.");
  assertExpectedItem(rawItem, args, "Delete");
  const item = simplifyItem(rawItem);
  if (args.dryRun !== false) {
    return { dryRun: true, wouldDelete: item };
  }
  if (args.confirmed !== true) {
    return {
      dryRun: false,
      confirmed: false,
      wouldDelete: item,
      requiredToDelete: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!args.expectedName && !args.expectedId) {
    return {
      dryRun: false,
      confirmed: true,
      wouldDelete: item,
      requiredToDelete: "Provide expectedName or expectedId for live deletes."
    };
  }
  await graph(itemBase(args), { method: "DELETE" });
  return { dryRun: false, confirmed: true, deleted: item };
}

async function restoreDeleted(args = {}) {
  if (args.expectedId && args.expectedId !== args.itemId) {
    throw new Error(`Restore expected item ID ${args.expectedId}, but got ${args.itemId}. Refusing to continue.`);
  }
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldRestore: {
      itemId: args.itemId,
      destinationParentPath: args.destinationParentPath,
      destinationParentItemId: args.destinationParentItemId,
      destinationParentPreset: args.destinationParentPreset,
      destinationParentRelativePath: args.destinationParentRelativePath,
      newName: args.newName
    },
    permissionNote: "OneDrive Personal restore may require Microsoft Graph delegated Files.ReadWrite.All in addition to the plugin's standard Files.ReadWrite scope."
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...preview,
      requiredToRestore: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  const body = {};
  if (args.destinationParentPath || args.destinationParentItemId || args.destinationParentPreset) {
    const parent = await resolveDestinationParent(args);
    body.parentReference = { id: parent.id };
  }
  if (args.newName) body.name = args.newName;
  const result = await graph(`/me/drive/items/${encodeURIComponent(args.itemId)}/restore`, {
    method: "POST",
    body: JSON.stringify(body)
  });
  return { dryRun: false, confirmed: true, restored: simplifyItem(result) };
}

function textResult(value, isError = false) {
  const text = typeof value === "string" ? value : JSON.stringify(value, null, 2);
  return { content: [{ type: "text", text }], isError };
}

function sendResult(id, result) {
  process.stdout.write(JSON.stringify({ jsonrpc: "2.0", id, result }) + "\n");
}

function sendError(id, code, message) {
  process.stdout.write(JSON.stringify({ jsonrpc: "2.0", id, error: { code, message } }) + "\n");
}

async function callTool(name, args = {}) {
  switch (name) {
    case "onedrive_config": {
      const status = publicConfig();
      if (args.checkToken) {
        try {
          await getAccessToken();
          status.accessTokenAvailable = true;
        } catch (error) {
          status.accessTokenAvailable = false;
          status.tokenCheckError = error.message;
        }
      }
      return textResult(status);
    }
    case "onedrive_auth_device_start":
      return textResult(await startDeviceLogin(args));
    case "onedrive_auth_device_poll":
      return textResult(await pollDeviceLogin(args));
    case "onedrive_logout": {
      tokenCache = null;
      pendingDevice = null;
      return textResult({ memoryCleared: true, keychainTokenDeleted: args.deleteKeychainToken ? deleteKeychainToken() : false });
    }
    case "onedrive_me":
      return textResult(await graph("/me"));
    case "onedrive_drive":
      return textResult(await graph("/me/drive"));
    case "onedrive_presets":
      return textResult({ pathPresets: pathPresets(), configPath });
    case "onedrive_list":
      return textResult(await list(args));
    case "onedrive_list_all":
      return textResult(await listAll(args));
    case "onedrive_search":
      return textResult(await search(args));
    case "onedrive_search_all":
      return textResult(await searchAll(args));
    case "onedrive_delta":
      return textResult(await delta(args));
    case "onedrive_get_info":
      return textResult(await getInfo(args));
    case "onedrive_read_text":
      return textResult(await readText(args));
    case "onedrive_download":
      return textResult(await download(args));
    case "onedrive_download_excel":
      return textResult(await downloadOffice(args, "excel"));
    case "onedrive_download_word":
      return textResult(await downloadOffice(args, "word"));
    case "onedrive_download_powerpoint":
      return textResult(await downloadOffice(args, "powerpoint"));
    case "onedrive_upload":
      return textResult(await upload(args));
    case "onedrive_write_text":
      return textResult(await writeText(args));
    case "onedrive_create_folder":
      return textResult(await createFolder(args));
    case "onedrive_rename":
      return textResult(await rename(args));
    case "onedrive_move":
      return textResult(await moveItem(args));
    case "onedrive_copy":
      return textResult(await copyItem(args));
    case "onedrive_create_sharing_link":
      return textResult(await createSharingLink(args));
    case "onedrive_permissions":
      return textResult(await permissions(args));
    case "onedrive_restore_deleted":
      return textResult(await restoreDeleted(args));
    case "onedrive_delete":
      return textResult(await deleteItem(args));
    default:
      return textResult(`Unknown tool: ${name}`, true);
  }
}

async function handleRequest(message) {
  const { id, method, params = {} } = message;
  try {
    if (method === "initialize") {
      sendResult(id, {
        protocolVersion: "2024-11-05",
        capabilities: { tools: {} },
        serverInfo: { name: "onedrive", version: "0.1.0" }
      });
      return;
    }
    if (method === "tools/list") {
      sendResult(id, { tools });
      return;
    }
    if (method === "tools/call") {
      sendResult(id, await callTool(params.name, params.arguments || {}));
      return;
    }
    if (method?.startsWith("notifications/")) return;
    if (id !== undefined) sendError(id, -32601, `Method not found: ${method}`);
  } catch (error) {
    if (id !== undefined) sendResult(id, textResult(error.message, true));
    else console.error(error);
  }
}

let buffer = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
  buffer += chunk;
  for (;;) {
    const newline = buffer.indexOf("\n");
    if (newline === -1) break;
    const line = buffer.slice(0, newline).trim();
    buffer = buffer.slice(newline + 1);
    if (!line) continue;
    try {
      void handleRequest(JSON.parse(line));
    } catch (error) {
      sendError(null, -32700, `Parse error: ${error.message}`);
    }
  }
});
