#!/usr/bin/env node

import { execFileSync, spawn } from "node:child_process";
import { createReadStream, createWriteStream, readFileSync } from "node:fs";
import { appendFile, chmod, copyFile, lstat, mkdir, open, readFile, readdir, realpath, rename as renameFile, rm, stat, writeFile } from "node:fs/promises";
import { basename, dirname, extname, isAbsolute, join, parse, relative, resolve } from "node:path";
import { homedir, tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { createHash, randomUUID } from "node:crypto";
import { AsyncLocalStorage } from "node:async_hooks";
import { addSemanticAnchors, resolveSemanticOperations } from "./semantic-anchors.mjs";
import { applyTextPatch, boundedLineDiff, decodeTextBuffer } from "./text-patch.mjs";
import { createAuthVault } from "./auth-vault.mjs";
import { oauthChallenge, oauthSettings, toolSecuritySchemes } from "./oauth.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const pluginRoot = resolve(__dirname, "..");
const pluginManifest = JSON.parse(readFileSync(join(pluginRoot, ".codex-plugin", "plugin.json"), "utf8"));
const chatgptIconDataUri = `data:image/png;base64,${readFileSync(join(pluginRoot, "assets", "chatgpt-icon.png")).toString("base64")}`;
const defaultStorageRoot = join(homedir(), ".codex", "onedrive-plugin");
const configPath = join(defaultStorageRoot, "config.json");
let localConfigReadError = null;
const localConfig = readLocalConfig();
const storageRoot = resolve(process.env.ONEDRIVE_STORAGE_ROOT || localConfig.storageRoot || defaultStorageRoot);
const downloadRoot = join(storageRoot, "downloads");
const cacheRoot = resolve(process.env.ONEDRIVE_CACHE_ROOT || localConfig.cacheRoot || join(storageRoot, "cache"));
const cachePath = join(cacheRoot, "metadata-cache.json");
const contentIndexPath = join(cacheRoot, "content-index.json");
const metadataCacheLockPath = join(cacheRoot, "metadata-cache.lock");
const contentIndexLockPath = join(cacheRoot, "content-index.lock");
const updateRoot = join(storageRoot, "updates");
const backupRoot = join(storageRoot, "backups");
const auditRoot = join(storageRoot, "audit");
const auditPath = join(auditRoot, "mutations.jsonl");
const auditLockPath = join(auditRoot, "mutations.lock");
const officeEditingRoot = join(storageRoot, "office-editing");
const chatgptUploadRoot = join(storageRoot, "chatgpt-uploads");
const workspaceStateRoot = join(storageRoot, "workspaces");
const managedWorkspaceRootName = "Codex Editing Drafts";
const watchStateRoot = join(storageRoot, "watches");
const officeHelperPath = join(pluginRoot, "scripts", "office-openxml.py");
const commonTextHelperPath = join(pluginRoot, "scripts", "common-text.py");
const officePythonPath = process.env.ONEDRIVE_OFFICE_PYTHON || localConfig.officeEditing?.pythonPath || "/usr/bin/python3";
const maxOfficePackageBytes = 250 * 1024 * 1024;
const maxCommonExtractionBytes = 100 * 1024 * 1024;
const localOneDriveSyncRoots = [
  { path: join(homedir(), "Library", "CloudStorage", "OneDrive"), prefix: false },
  { path: join(homedir(), "Library", "CloudStorage", "OneDrive-"), prefix: true },
  { path: join(homedir(), "OneDrive"), prefix: false },
  { path: join(homedir(), "OneDrive - "), prefix: true },
  ...readAdditionalLocalSyncRoots()
];
const textFileLimit = 5 * 1024 * 1024;
const maxTextFileReadLimit = 10 * 1024 * 1024;
const defaultMaxIndexedFileSize = 512 * 1024;
const simpleUploadLimit = 250 * 1024 * 1024;
const uploadChunkUnit = 320 * 1024;
const defaultUploadChunkSize = 10 * 1024 * 1024;
const maxUploadChunkSize = 60 * 1024 * 1024;
const chatgptToolResponseByteLimit = 1024 * 1024;
const chatgptFetchTextByteLimit = 192 * 1024;
const chatgptInitialFetchTextByteLimit = 32 * 1024;
const chatgptFetchChunkByteLimit = 64 * 1024;
const chatgptFetchSnapshotTtlMs = 10 * 60 * 1000;
const chatgptFetchSnapshotMaxEntries = 32;
const chatgptStaleCacheMaxAgeSeconds = 24 * 60 * 60;
const chatgptRevalidationCooldownMs = 30 * 1000;
const defaultSelect = "id,name,size,folder,file,webUrl,createdDateTime,lastModifiedDateTime,parentReference,eTag,cTag,deleted";
const textExtensions = new Set([
  ".bat", ".c", ".cfg", ".conf", ".cpp", ".cs", ".css", ".csv", ".env", ".go", ".h", ".hpp", ".htm",
  ".html", ".ics", ".ini", ".ipynb", ".java", ".js", ".json", ".jsonl", ".jsx", ".log", ".md", ".mjs",
  ".ndjson", ".php", ".properties", ".py", ".rb", ".rs", ".rst", ".sh", ".sql", ".svelte", ".svg", ".tex",
  ".toml", ".ts", ".tsv", ".tsx", ".txt", ".vcf", ".vue", ".xml", ".yaml", ".yml"
]);
const textMimePrefixes = ["text/"];
const textMimeTypes = new Set([
  "application/csv",
  "application/ecmascript",
  "application/javascript",
  "application/json",
  "application/sql",
  "application/x-ndjson",
  "application/x-javascript",
  "application/xml",
  "text/calendar",
  "text/tab-separated-values",
  "text/vcard",
  "image/svg+xml"
]);
const canonicalTextMimeTypes = new Map([
  [".csv", "text/csv"],
  [".css", "text/css"],
  [".htm", "text/html"],
  [".html", "text/html"],
  [".js", "text/javascript"],
  [".json", "application/json"],
  [".jsonl", "application/x-ndjson"],
  [".md", "text/markdown"],
  [".mjs", "text/javascript"],
  [".svg", "image/svg+xml"],
  [".ts", "text/plain"],
  [".tsx", "text/plain"],
  [".txt", "text/plain"],
  [".tsv", "text/tab-separated-values"],
  [".xml", "application/xml"],
  [".yaml", "application/yaml"],
  [".yml", "application/yaml"]
]);
const searchTombstoneTtlMs = 24 * 60 * 60 * 1000;
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
let tokenRefreshPromise = null;
let authGeneration = 0;
let deviceLoginGeneration = 0;
let storageScopeGeneration = 0;
let metadataCacheMemory = null;
let contentIndexMemory = null;
let metadataCacheMemoryGeneration = null;
let contentIndexMemoryGeneration = null;
let metadataCacheLoadPromise = null;
let contentIndexLoadPromise = null;
let metadataCacheLoadGeneration = null;
let contentIndexLoadGeneration = null;
let metadataCacheFileVersion = null;
let contentIndexFileVersion = null;
let metadataMutationQueue = Promise.resolve();
let contentIndexMutationQueue = Promise.resolve();
let activeStorageScopePromise = null;
let activeStorageScopeKey = null;
let activeStorageScopeGeneration = null;
const testAuthContextId = process.env.ONEDRIVE_TEST_AUTH_CONTEXT_ID || randomUUID();

const previewTokens = new Map();
const watchTimers = new Map();
const excelSessionPool = new Map();
const chatgptFetchSnapshots = new Map();
const chatgptRevalidations = new Map();
const chatgptRevalidationLastStartedAt = new Map();
let watchesLoaded = false;
const previewTokenTtlMs = 15 * 60 * 1000;
const previewScopedTools = new Set([
  "onedrive_preview_actions",
  "onedrive_upload", "onedrive_upload_file", "onedrive_write_text", "onedrive_batch_delete", "onedrive_delete", "onedrive_permanent_delete",
  "onedrive_rename", "onedrive_move", "onedrive_copy",
  "onedrive_create_sharing_link", "onedrive_invite_permission", "onedrive_revoke_permission",
  "onedrive_batch_revoke_permissions", "onedrive_restore_deleted", "onedrive_word_batch_update",
  "onedrive_excel_batch_update", "onedrive_powerpoint_batch_update", "onedrive_office_batch_transform",
  "onedrive_office_restore_backup", "onedrive_patch_text", "onedrive_restore_version",
  "onedrive_workspace_create", "onedrive_workspace_promote", "onedrive_workspace_abandon"
]);
const partialBatchMutationWarning = "Batch live mutations are preflighted but not atomic; if a later item fails, earlier remote changes may already have taken effect.";
const toolCallContext = new AsyncLocalStorage();

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
const previewTokenSchema = {
  type: "string",
  description: "Token returned by the immediately preceding dry-run preview for this exact high-risk operation."
};
const jsonScalarSchema = { anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }, { type: "null" }] };
const strictCommand = (required, properties, anyOf) => ({ type: "object", required, properties, ...(anyOf ? { anyOf } : {}), additionalProperties: false });
const structuredTextPatchOperationSchema = {
  anyOf: [
    strictCommand(["op", "path", "value"], { op: { enum: ["add", "replace", "test"] }, path: { type: "string" }, value: {} }),
    strictCommand(["op", "path"], { op: { const: "remove" }, path: { type: "string" } }),
    strictCommand(["op", "path", "from"], { op: { enum: ["copy", "move"] }, path: { type: "string" }, from: { type: "string" } }),
    strictCommand(["op", "key", "values"], { op: { const: "insert" }, key: jsonScalarSchema, values: { type: "array", minItems: 1, maxItems: 1000, items: strictCommand(["column", "value"], { column: { type: "string", minLength: 1 }, value: {} }) } }),
    strictCommand(["op", "key", "values"], { op: { const: "update" }, key: jsonScalarSchema, values: { type: "array", minItems: 1, maxItems: 1000, items: strictCommand(["column", "value"], { column: { type: "string", minLength: 1 }, value: {} }) } }),
    strictCommand(["op", "key"], { op: { const: "delete" }, key: jsonScalarSchema })
  ]
};
const officeTargetProperties = {
  path: { type: "string", description: "Office file path relative to OneDrive root." },
  itemId: { type: "string", description: "Office file drive item ID." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const officeStructuredSearchProperties = {
  searchText: { type: "string", minLength: 1, description: "Optional in-package text/value query; returns bounded structured matches alongside the normal read result." },
  matchCase: { type: "boolean", default: false },
  maxMatches: { type: "integer", minimum: 1, maximum: 5000, default: 200 }
};
const officeBatchCommonProperties = {
  ...officeTargetProperties,
  operations: {
    type: "array",
    minItems: 1,
    maxItems: 100,
    items: { type: "object", additionalProperties: true },
    description: "Typed Office edit operations executed in order. Use the corresponding read tool first to ground indexes, sheet names, slide indexes, and shape IDs."
  },
  dryRun: { type: "boolean", default: true },
  confirmed: { type: "boolean", default: false },
  expectedName: { type: "string", description: "Safety check for the resolved remote file name." },
  expectedId: { type: "string", description: "Safety check for the resolved remote drive item ID." },
  expectedETag: { type: "string", description: "Optional explicit revision guard for live Office edits. When provided, it must exactly match the current remote eTag." },
  previewToken: previewTokenSchema,
  createBackup: { type: "boolean", default: true },
  verify: { type: "boolean", default: true },
  allowMacros: { type: "boolean", default: false, description: "Allow editing a macro-enabled package while preserving VBA parts. Never executes macros." },
  allowSignedPackage: { type: "boolean", default: false, description: "Reserved compatibility flag. Digitally signed packages are always refused because native edits invalidate signatures." }
};
const semanticAnchorSchema = {
  type: "object",
  required: ["kind"],
  properties: {
    kind: { type: "string", enum: ["paragraph", "table", "contentControl", "worksheet", "range", "excelTable", "slide", "shape"] },
    part: { type: "string" }, headingPath: { type: "array", maxItems: 16, items: { type: "string" } },
    textHash: { type: "string" }, beforeHash: { type: ["string", "null"] }, afterHash: { type: ["string", "null"] }, fingerprint: { type: "string" },
    headers: { type: "array", maxItems: 256, items: { type: "string" } }, id: { type: ["string", "integer", "null"] }, tag: { type: ["string", "null"] }, title: { type: ["string", "null"] },
    name: { type: "string" }, displayName: { type: ["string", "null"] }, sheet: { type: "string" }, address: { type: "string" }, valueHash: { type: "string" }, formulaHash: { type: "string" },
    slideId: { type: ["string", "integer"] }, slideTitle: { type: ["string", "null"] }, shapeId: { type: ["string", "integer", "null"] }, shapeName: { type: ["string", "null"] }, altText: { type: ["string", "null"] }
  },
  additionalProperties: false
};
const semanticOperationProperties = {
  anchor: semanticAnchorSchema,
  rebasePolicy: { type: "string", enum: ["unique", "fail"], default: "unique" }
};
const semanticSelectorKeys = new Set(["paragraphIndex", "afterIndex", "afterParagraphIndex", "tableIndex", "contentControlIndex", "sheet", "address", "table", "slideIndex", "shapeId"]);
const operationObject = (required, properties) => {
  const hasSemanticSelector = required.some((key) => semanticSelectorKeys.has(key));
  const anchorRequired = [...new Set([...required.filter((key) => !semanticSelectorKeys.has(key)), "anchor"])];
  return {
    type: "object",
    ...(hasSemanticSelector ? { anyOf: [{ required }, { required: anchorRequired }] } : { required }),
    properties: { ...semanticOperationProperties, ...properties },
    additionalProperties: false
  };
};
const replaceTextOperation = operationObject(["type", "find", "replace"], {
  type: { const: "replaceText" }, find: { type: "string", minLength: 1 }, replace: { type: "string" }, matchCase: { type: "boolean" }, all: { type: "boolean" }
});
const wordOperationsSchema = {
  type: "array", minItems: 1, maxItems: 100, items: { anyOf: [
    { ...replaceTextOperation, properties: { ...replaceTextOperation.properties, scope: { type: "string", enum: ["document", "headers", "footers", "all"] } } },
    operationObject(["type", "paragraphIndex", "text"], { type: { const: "setParagraphText" }, paragraphIndex: { type: "integer", minimum: 0 }, text: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "paragraphIndex", "style"], { type: { const: "setParagraphStyle" }, paragraphIndex: { type: "integer", minimum: 0 }, style: { type: "string", minLength: 1 }, part: { type: "string" } }),
    operationObject(["type", "text"], { type: { const: "insertParagraph" }, afterIndex: { type: "integer", minimum: 0 }, text: { type: "string" }, style: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "tableIndex", "rowIndex", "columnIndex", "text"], { type: { const: "setTableCell" }, tableIndex: { type: "integer", minimum: 0 }, rowIndex: { type: "integer", minimum: 0 }, columnIndex: { type: "integer", minimum: 0 }, text: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "text"], { type: { const: "setContentControlText" }, contentControlIndex: { type: "integer", minimum: 0 }, id: { type: "string" }, tag: { type: "string" }, text: { type: "string" }, part: { type: "string" } }),
    operationObject(["type", "paragraphIndex", "text", "url"], { type: { const: "addHyperlink" }, paragraphIndex: { type: "integer", minimum: 0 }, text: { type: "string", minLength: 1 }, url: { type: "string", minLength: 1 }, part: { type: "string" } }),
    operationObject(["type", "paragraphIndex", "text"], { type: { const: "addComment" }, paragraphIndex: { type: "integer", minimum: 0 }, text: { type: "string", minLength: 1 }, author: { type: "string", maxLength: 255 }, initials: { type: "string", maxLength: 16 } }),
    operationObject(["type", "rows"], { type: { const: "insertTable" }, afterParagraphIndex: { type: "integer", minimum: 0 }, rows: { type: "array", minItems: 1, maxItems: 100, items: { type: "array", minItems: 1, maxItems: 50, items: { type: "string" } } }, style: { type: "string", minLength: 1 } })
    ,operationObject(["type", "paragraphIndex", "base64", "contentType"], { type: { const: "insertImage" }, paragraphIndex: { type: "integer", minimum: 0 }, base64: { type: "string", minLength: 1, maxLength: 36700160 }, contentType: { type: "string", enum: ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff"] }, width: { type: "integer", minimum: 1 }, height: { type: "integer", minimum: 1 }, altText: { type: "string" }, part: { type: "string" } })
    ,operationObject(["type", "imageIndex", "base64", "contentType"], { type: { const: "replaceImage" }, imageIndex: { type: "integer", minimum: 0 }, base64: { type: "string", minLength: 1, maxLength: 36700160 }, contentType: { type: "string", enum: ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff"] }, part: { type: "string" } })
    ,operationObject(["type", "paragraphIndex"], { type: { const: "createContentControl" }, paragraphIndex: { type: "integer", minimum: 0 }, id: { type: "string" }, tag: { type: "string" }, title: { type: "string" }, text: { type: "string" }, part: { type: "string" } })
    ,operationObject(["type"], { type: { const: "deleteContentControl" }, contentControlIndex: { type: "integer", minimum: 0 }, id: { type: "string" }, tag: { type: "string" }, preserveContent: { type: "boolean", default: true }, part: { type: "string" } })
    ,operationObject(["type", "paragraphIndex", "name"], { type: { const: "createBookmark" }, paragraphIndex: { type: "integer", minimum: 0 }, name: { type: "string", pattern: "^[A-Za-z_][A-Za-z0-9_]{0,39}$" }, part: { type: "string" } })
    ,operationObject(["type", "name"], { type: { const: "deleteBookmark" }, name: { type: "string", minLength: 1 }, part: { type: "string" } })
    ,operationObject(["type", "tableIndex", "rowIndex", "values"], { type: { const: "insertTableRow" }, tableIndex: { type: "integer", minimum: 0 }, rowIndex: { type: "integer", minimum: 0 }, values: { type: "array", minItems: 1, maxItems: 50, items: { type: "string" } }, part: { type: "string" } })
    ,operationObject(["type", "tableIndex", "rowIndex"], { type: { const: "deleteTableRow" }, tableIndex: { type: "integer", minimum: 0 }, rowIndex: { type: "integer", minimum: 0 }, part: { type: "string" } })
    ,operationObject(["type", "tableIndex", "columnIndex"], { type: { const: "insertTableColumn" }, tableIndex: { type: "integer", minimum: 0 }, columnIndex: { type: "integer", minimum: 0 }, values: { type: "array", maxItems: 100, items: { type: "string" } }, part: { type: "string" } })
    ,operationObject(["type", "tableIndex", "columnIndex"], { type: { const: "deleteTableColumn" }, tableIndex: { type: "integer", minimum: 0 }, columnIndex: { type: "integer", minimum: 0 }, part: { type: "string" } })
    ,operationObject(["type", "part", "text"], { type: { const: "setHeaderFooterText" }, part: { type: "string", pattern: "^word/(header|footer)[0-9]+\\.xml$" }, text: { type: "string" } })
    ,operationObject(["type", "sectionIndex"], { type: { const: "setSectionProperties" }, sectionIndex: { type: "integer", minimum: 0 }, orientation: { type: "string", enum: ["portrait", "landscape"] }, pageWidth: { type: "integer", minimum: 1 }, pageHeight: { type: "integer", minimum: 1 }, marginTop: { type: "integer", minimum: 0 }, marginRight: { type: "integer", minimum: 0 }, marginBottom: { type: "integer", minimum: 0 }, marginLeft: { type: "integer", minimum: 0 } })
  ] }
};
const excelBaseOperation = { sheet: { type: "string", minLength: 1 }, address: { type: "string", pattern: "^[A-Za-z]{1,3}[1-9][0-9]*(?::[A-Za-z]{1,3}[1-9][0-9]*)?$" } };
const excelOperationsSchema = {
  type: "array", minItems: 1, maxItems: 100, items: { anyOf: [
    operationObject(["type", "sheet", "address"], { type: { const: "setCell" }, ...excelBaseOperation, value: {} }),
    operationObject(["type", "sheet", "address", "formula"], { type: { const: "setFormula" }, ...excelBaseOperation, formula: { type: "string" }, value: {} }),
    operationObject(["type", "sheet", "address"], { type: { const: "setRange" }, ...excelBaseOperation, values: { type: "array", items: { type: "array" } }, formulas: { type: "array", items: { type: "array" } } }),
    operationObject(["type", "sheet", "address"], { type: { const: "clearRange" }, ...excelBaseOperation, contents: { type: "boolean" }, format: { type: "boolean" } }),
    operationObject(["type", "sheet", "address", "styleIndex"], { type: { const: "setStyle" }, ...excelBaseOperation, styleIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "sheet", "address", "formatCode"], { type: { const: "setNumberFormat" }, ...excelBaseOperation, formatCode: { type: "string", minLength: 1, maxLength: 255 } }),
    operationObject(["type", "sheet", "address", "formula", "fillColor"], { type: { const: "addConditionalFormat" }, ...excelBaseOperation, ruleType: { type: "string", enum: ["expression", "cellIs"] }, operator: { type: "string", enum: ["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"] }, formula: { type: "string", minLength: 1 }, fillColor: { type: "string", pattern: "^[0-9A-Fa-f]{6}$" } }),
    operationObject(["type", "sheet", "address", "validationType"], { type: { const: "setDataValidation" }, ...excelBaseOperation, validationType: { type: "string", enum: ["whole", "decimal", "list", "date", "time", "textLength", "custom"] }, operator: { type: "string" }, formula1: { type: "string" }, formula2: { type: "string" }, allowBlank: { type: "boolean" } }),
    operationObject(["type", "sheet"], { type: { const: "freezePanes" }, sheet: { type: "string", minLength: 1 }, rows: { type: "integer", minimum: 0, maximum: 1048575 }, columns: { type: "integer", minimum: 0, maximum: 16383 } }),
    operationObject(["type", "sheet", "address", "width"], { type: { const: "setColumnWidth" }, ...excelBaseOperation, width: { type: "number", exclusiveMinimum: 0, maximum: 255 } }),
    operationObject(["type", "table", "values"], { type: { const: "addTableRow" }, table: { type: "string", minLength: 1 }, index: { type: ["integer", "null"], minimum: 0 }, values: { type: "array", minItems: 1, maxItems: 1000, items: { type: "array", minItems: 1, maxItems: 16384 } } }),
    operationObject(["type", "table", "index"], { type: { const: "deleteTableRow" }, table: { type: "string", minLength: 1 }, index: { type: "integer", minimum: 0 } }),
    operationObject(["type", "table", "enabled"], { type: { const: "setTableTotals" }, table: { type: "string", minLength: 1 }, enabled: { type: "boolean" }, columns: { type: "array", maxItems: 16384, items: operationObject(["column"], { column: { anyOf: [{ type: "string", minLength: 1 }, { type: "integer", minimum: 0 }] }, function: { type: "string", enum: ["average", "count", "countNums", "custom", "max", "min", "none", "stdDev", "sum", "var"] }, label: { type: "string" }, formula: { type: "string", minLength: 1 } }) } }),
    operationObject(["type", "sheet", "chartType", "sourceData"], { type: { const: "createChart" }, sheet: { type: "string", minLength: 1 }, chartType: { type: "string", enum: ["BarClustered", "ColumnClustered", "Line", "Pie"] }, sourceData: { type: "string", minLength: 1 }, seriesBy: { type: "string", enum: ["Auto", "Columns", "Rows"] }, name: { type: "string", minLength: 1 }, titleText: { type: "string" }, height: { type: "number", exclusiveMinimum: 0 }, width: { type: "number", exclusiveMinimum: 0 }, left: { type: "number", minimum: 0 }, top: { type: "number", minimum: 0 } }),
    operationObject(["type", "sheet", "chart"], { type: { const: "updateChart" }, sheet: { type: "string", minLength: 1 }, chart: { type: "string", minLength: 1 }, chartType: { type: "string", enum: ["BarClustered", "ColumnClustered", "Line", "Pie"] }, name: { type: "string", minLength: 1 }, titleText: { type: "string" }, height: { type: "number", exclusiveMinimum: 0 }, width: { type: "number", exclusiveMinimum: 0 }, left: { type: "number", minimum: 0 }, top: { type: "number", minimum: 0 }, sourceData: { type: "string", minLength: 1 }, seriesBy: { type: "string", enum: ["Auto", "Columns", "Rows"] } }),
    operationObject(["type", "sheet", "newName"], { type: { const: "renameSheet" }, sheet: { type: "string", minLength: 1 }, newName: { type: "string", minLength: 1, maxLength: 31 } }),
    operationObject(["type", "name", "formula"], { type: { const: "setDefinedName" }, name: { type: "string", minLength: 1 }, formula: { type: "string" } }),
    operationObject(["type"], { type: { const: "recalculate" } }),
    operationObject(["type", "name"], { type: { const: "addWorksheet" }, name: { type: "string", minLength: 1, maxLength: 31 } }),
    operationObject(["type", "sheet"], { type: { const: "deleteWorksheet" }, sheet: { type: "string", minLength: 1 } }),
    operationObject(["type", "sheet", "address", "name"], { type: { const: "addTable" }, ...excelBaseOperation, name: { type: "string", minLength: 1 }, hasHeaders: { type: "boolean", default: true } }),
    operationObject(["type", "table"], { type: { const: "deleteTable" }, table: { type: "string", minLength: 1 }, preserveData: { type: "boolean", default: true } }),
    operationObject(["type", "sheet", "address"], { type: { const: "mergeRange" }, ...excelBaseOperation }),
    operationObject(["type", "sheet", "address"], { type: { const: "unmergeRange" }, ...excelBaseOperation }),
    operationObject(["type", "sheet", "address", "keys"], { type: { const: "sortRange" }, ...excelBaseOperation, keys: { type: "array", minItems: 1, maxItems: 64, items: { type: "object", required: ["column"], properties: { column: { type: "integer", minimum: 0 }, descending: { type: "boolean" } }, additionalProperties: false } }, hasHeaders: { type: "boolean", default: false } }),
    operationObject(["type", "sheet", "address"], { type: { const: "setAutoFilter" }, ...excelBaseOperation, column: { type: "integer", minimum: 0 }, criteria: { type: "string" }, clear: { type: "boolean" } }),
    operationObject(["type", "sheet", "address", "url"], { type: { const: "setHyperlink" }, ...excelBaseOperation, url: { type: "string", minLength: 1 }, display: { type: "string" } }),
    operationObject(["type", "sheet", "address", "text"], { type: { const: "addNote" }, ...excelBaseOperation, text: { type: "string" }, author: { type: "string" } }),
    operationObject(["type", "sheet", "address"], { type: { const: "deleteNote" }, ...excelBaseOperation }),
    operationObject(["type", "sheet", "base64", "contentType", "fromAddress"], { type: { const: "insertImage" }, sheet: { type: "string", minLength: 1 }, base64: { type: "string", minLength: 1, maxLength: 36700160 }, contentType: { type: "string", enum: ["image/png", "image/jpeg"] }, fromAddress: { type: "string" }, toAddress: { type: "string" }, altText: { type: "string" } }),
    operationObject(["type", "sheet", "chart"], { type: { const: "formatChart" }, sheet: { type: "string", minLength: 1 }, chart: { type: "string", minLength: 1 }, titleText: { type: "string" }, legendPosition: { type: "string", enum: ["top", "bottom", "left", "right", "none"] }, style: { type: "integer", minimum: 1, maximum: 48 } }),
    operationObject(["type", "sheet", "enabled"], { type: { const: "setSheetProtection" }, sheet: { type: "string", minLength: 1 }, enabled: { type: "boolean" }, allowSelectLockedCells: { type: "boolean" }, allowSelectUnlockedCells: { type: "boolean" }, allowFormatCells: { type: "boolean" } }),
    operationObject(["type"], { type: { const: "refreshPivot" }, cacheId: { type: "integer", minimum: 1 } })
  ] }
};
const powerpointSelector = { slideIndex: { type: "integer", minimum: 0 }, shapeId: { type: ["string", "integer"] } };
const powerpointOperationsSchema = {
  type: "array", minItems: 1, maxItems: 100, items: { anyOf: [
    { ...replaceTextOperation, properties: { ...replaceTextOperation.properties, ...powerpointSelector } },
    operationObject(["type", "slideIndex", "shapeId", "text"], { type: { const: "setShapeText" }, ...powerpointSelector, text: { type: "string" } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "setShapeGeometry" }, ...powerpointSelector, x: { type: "integer" }, y: { type: "integer" }, width: { type: "integer", minimum: 0 }, height: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex", "shapeId", "rowIndex", "columnIndex", "text"], { type: { const: "setTableCell" }, ...powerpointSelector, rowIndex: { type: "integer", minimum: 0 }, columnIndex: { type: "integer", minimum: 0 }, text: { type: "string" } }),
    operationObject(["type", "slideIndex", "text", "x", "y", "width", "height"], { type: { const: "addTextBox" }, slideIndex: { type: "integer", minimum: 0 }, shapeId: { type: "integer", minimum: 1 }, name: { type: "string", minLength: 1 }, text: { type: "string" }, x: { type: "integer" }, y: { type: "integer" }, width: { type: "integer", minimum: 1 }, height: { type: "integer", minimum: 1 } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "deleteShape" }, ...powerpointSelector }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "setTextStyle" }, ...powerpointSelector, fontFamily: { type: "string", minLength: 1 }, fontSize: { type: "number", minimum: 1, maximum: 400 }, bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, color: { type: "string", pattern: "^[0-9A-Fa-f]{6}$" } }),
    operationObject(["type", "slideIndex", "shapeId", "base64", "contentType"], { type: { const: "replaceImage" }, ...powerpointSelector, base64: { type: "string", minLength: 1, maxLength: 36700160 }, contentType: { type: "string", enum: ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff"] } }),
    operationObject(["type", "slideIndex", "text"], { type: { const: "setNotes" }, slideIndex: { type: "integer", minimum: 0 }, text: { type: "string" } }),
    operationObject(["type", "slideIndex"], { type: { const: "duplicateSlide" }, slideIndex: { type: "integer", minimum: 0 }, toIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex"], { type: { const: "deleteSlide" }, slideIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex", "toIndex"], { type: { const: "moveSlide" }, slideIndex: { type: "integer", minimum: 0 }, toIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type"], { type: { const: "addSlide" }, afterIndex: { type: "integer", minimum: 0 }, layoutName: { type: "string" } }),
    operationObject(["type", "slideIndex", "base64", "contentType", "x", "y", "width", "height"], { type: { const: "addImage" }, slideIndex: { type: "integer", minimum: 0 }, base64: { type: "string", minLength: 1, maxLength: 36700160 }, contentType: { type: "string", enum: ["image/png", "image/jpeg", "image/gif"] }, x: { type: "integer" }, y: { type: "integer" }, width: { type: "integer", minimum: 1 }, height: { type: "integer", minimum: 1 }, name: { type: "string" }, altText: { type: "string" } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "cropImage" }, ...powerpointSelector, left: { type: "number", minimum: 0, maximum: 1 }, top: { type: "number", minimum: 0, maximum: 1 }, right: { type: "number", minimum: 0, maximum: 1 }, bottom: { type: "number", minimum: 0, maximum: 1 } }),
    operationObject(["type", "slideIndex", "rows", "x", "y", "width", "height"], { type: { const: "addTable" }, slideIndex: { type: "integer", minimum: 0 }, rows: { type: "array", minItems: 1, maxItems: 100, items: { type: "array", minItems: 1, maxItems: 50, items: { type: "string" } } }, x: { type: "integer" }, y: { type: "integer" }, width: { type: "integer", minimum: 1 }, height: { type: "integer", minimum: 1 } }),
    operationObject(["type", "slideIndex", "shapeId", "rowIndex"], { type: { const: "insertTableRow" }, ...powerpointSelector, rowIndex: { type: "integer", minimum: 0 }, values: { type: "array", maxItems: 50, items: { type: "string" } } }),
    operationObject(["type", "slideIndex", "shapeId", "rowIndex"], { type: { const: "deleteTableRow" }, ...powerpointSelector, rowIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex", "shapeId", "columnIndex"], { type: { const: "insertTableColumn" }, ...powerpointSelector, columnIndex: { type: "integer", minimum: 0 }, values: { type: "array", maxItems: 100, items: { type: "string" } } }),
    operationObject(["type", "slideIndex", "shapeId", "columnIndex"], { type: { const: "deleteTableColumn" }, ...powerpointSelector, columnIndex: { type: "integer", minimum: 0 } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "setShapeAltText" }, ...powerpointSelector, title: { type: "string" }, description: { type: "string" } }),
    operationObject(["type", "slideIndex", "shapeId", "position"], { type: { const: "setZOrder" }, ...powerpointSelector, position: { type: "string", enum: ["front", "back", "forward", "backward"] } }),
    operationObject(["type", "slideIndex", "shapeIds"], { type: { const: "groupShapes" }, slideIndex: { type: "integer", minimum: 0 }, shapeIds: { type: "array", minItems: 2, maxItems: 100, items: { type: ["string", "integer"] } }, name: { type: "string" } }),
    operationObject(["type", "slideIndex", "shapeId"], { type: { const: "ungroupShape" }, ...powerpointSelector }),
    operationObject(["type", "slideIndex", "layoutName"], { type: { const: "applySlideLayout" }, slideIndex: { type: "integer", minimum: 0 }, layoutName: { type: "string", minLength: 1 } })
  ] }
};
const pathTargetProperties = {
  path: { type: "string", description: "Item path relative to OneDrive root." },
  itemId: { type: "string", description: "Drive item ID." },
  preset: presetSchema,
  relativePath: relativePathSchema
};
const driveRecipientSchema = {
  type: "object",
  properties: {
    email: { type: "string", minLength: 1 },
    alias: { type: "string", minLength: 1 },
    objectId: { type: "string", minLength: 1 }
  },
  additionalProperties: false
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
  { required: ["remotePreset", "remoteRelativePath"] }
];

const tools = [
  {
    name: "onedrive_config",
    description: "Show OneDrive plugin configuration status without exposing secrets.",
    inputSchema: {
      type: "object",
      properties: {
        checkToken: {
          type: "boolean",
          default: false,
          description: "When true, try to get an access token from the configured secure authentication store or memory."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_auth_device_start",
    description: "Verify existing stored authentication first, and start Microsoft device-code login only when reauthentication is required or explicitly forced.",
    inputSchema: {
      type: "object",
      properties: {
        tenant: { type: "string", description: "Optional tenant override. Defaults to configured tenant or common." },
        scopes: { type: "string", description: "Optional space-separated scope override." },
        forceReauth: {
          type: "boolean",
          default: false,
          description: "Generate a new device code even when existing stored authentication is healthy. Use only for explicit account switching or consent repair."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_auth_device_poll",
    description: "Poll Microsoft token endpoint after the user completes device-code login, then save tokens in the configured secure authentication store.",
    inputSchema: {
      type: "object",
      properties: {
        deviceCode: { type: "string", description: "Optional device_code from onedrive_auth_device_start. Defaults to the latest pending code." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_logout",
    description: "Forget cached OneDrive tokens from memory and optionally delete the securely stored token.",
    inputSchema: {
      type: "object",
      properties: {
        deleteKeychainToken: { type: "boolean", default: false },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation to delete the securely stored refresh token."
        }
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
    name: "onedrive_doctor",
    description: "Run a bundled OneDrive plugin health check for config, auth, profile, drive, presets, and optional root listing.",
    inputSchema: {
      type: "object",
      properties: {
        checkRootList: {
          type: "boolean",
          default: true,
          description: "When true, list a few root items to verify Files.Read permissions."
        },
        rootListLimit: {
          type: "integer",
          minimum: 1,
          maximum: 20,
          default: 5
        },
        checkPresets: {
          type: "boolean",
          default: true,
          description: "When true, resolve every configured path preset and warn when a target does not exist."
        }
      },
      additionalProperties: false
    }
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
        limit: { type: "integer", minimum: 1, default: 100 },
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
    name: "onedrive_scan",
    description: "Recursively scan OneDrive folders from a starting folder, following pagination and subfolders up to safe caps.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        select: { type: "string", description: "Optional Graph $select fields." },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxResults: {
          type: "integer",
          minimum: 1,
          maximum: 5000,
          default: 500,
          description: "Maximum matching items returned in the response. The scan can inspect more items than it returns."
        },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        maxFolders: {
          type: "integer",
          minimum: 1,
          maximum: 10000,
          default: 1000,
          description: "Maximum folders to visit during recursive traversal."
        },
        scanConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          description: "Bounded number of folders scanned concurrently. Defaults to the configured concurrency limit, capped at 4."
        },
        nameContains: { type: "string", description: "Optional case-insensitive filename/folder-name substring filter." },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional file extension filter such as ['.pdf', 'pptx']. Folders are unaffected unless includeFolders is false."
        },
        includeFiles: { type: "boolean", default: true },
        includeFolders: { type: "boolean", default: true },
        stopAfterResults: {
          type: "boolean",
          default: false,
          description: "When true, stop traversal once maxResults matching items have been returned."
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_search",
    title: "Raw Graph search (advanced)",
    description: "Run one direct Microsoft Graph search query against the signed-in user's OneDrive. Use only when the user explicitly requests raw Graph search behavior; for normal file lookup by exact name, partial name, or fuzzy wording, use onedrive_find instead because it adds adaptive queries, ranking, cache evidence, and bounded scan fallback.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        limit: { type: "integer", minimum: 1, default: 50 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_search_all",
    title: "Paginated raw Graph search (advanced)",
    description: "Run one direct Microsoft Graph search query and follow its pagination up to a safe item cap. Use only for explicitly requested paginated raw Graph results; for broad or exhaustive user-facing file lookup, use onedrive_find_all instead.",
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
    name: "onedrive_find",
    title: "Find OneDrive files",
    description: "Default tool for normal OneDrive file lookup by exact name, partial name, or fuzzy user wording. This cache-assisted remote-first finder runs the canonical Graph query first, expands terms adaptively with bounded concurrency, ranks matches in memory, and optionally falls back to bounded recursive scans.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional hard extension filter such as ['.pdf', 'pptx']."
        },
        folderHints: {
          type: "array",
          items: { type: "string" },
          description: "Optional root-relative folders to try first during fallback scans."
        },
        includeFolders: { type: "boolean", default: false },
        maxResults: { type: "integer", minimum: 1, maximum: 50, default: 10 },
        maxSearchTerms: { type: "integer", minimum: 1, maximum: 12, default: 8 },
        searchConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Bounded concurrency for adaptive Graph search-term expansion after the canonical query."
        },
        searchPageSize: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        searchMaxItemsPerTerm: { type: "integer", minimum: 1, maximum: 500, default: 100 },
        scanFallback: { type: "boolean", default: true },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 10000, default: 1500 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 2000, default: 250 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 20 },
        scanConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Low-concurrency fallback scan fan-out across top folder hints."
        },
        minConfidenceForSearchOnly: {
          type: "integer",
          minimum: 0,
          maximum: 100,
          default: 78,
          description: "If the best search-only match scores below this, use scan fallback when enabled."
        },
        useCache: {
          type: "boolean",
          default: true,
          description: "Use the local metadata cache as a fast ranking source before live Graph search."
        },
        useContentIndex: {
          type: "boolean",
          default: true,
          description: "Use the local content index if it has already been built. This never fetches file contents during find."
        },
        contentMaxResults: {
          type: "integer",
          minimum: 0,
          maximum: 50,
          default: 10,
          description: "Maximum indexed-content matches to merge into ranked results."
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_find_all",
    title: "Find files across OneDrive",
    description: "Broader cache-assisted remote-first file locator. Uses adaptive bounded-concurrency Graph term expansion, common-folder hints, and bounded scan fallback when needed.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional hard extension filter such as ['.pdf', 'pptx']."
        },
        folderHints: {
          type: "array",
          items: { type: "string" },
          description: "Optional root-relative folders to scan before the default common folders."
        },
        includeFolders: { type: "boolean", default: false },
        maxResults: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxSearchTerms: { type: "integer", minimum: 1, maximum: 12, default: 8 },
        searchConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Bounded concurrency for adaptive Graph search-term expansion after the canonical query."
        },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        scanConcurrency: {
          type: "integer",
          minimum: 1,
          maximum: 4,
          default: 2,
          description: "Low-concurrency fallback scan fan-out across top folder hints."
        },
        searchPageSize: { type: "integer", minimum: 1, maximum: 200, default: 100 },
        searchMaxItemsPerTerm: { type: "integer", minimum: 1, maximum: 1000, default: 250 },
        useCache: {
          type: "boolean",
          default: true,
          description: "Use the local metadata cache as a fast ranking source before live Graph search."
        },
        useContentIndex: {
          type: "boolean",
          default: true,
          description: "Use the local content index if it has already been built. This never fetches file contents during find_all."
        },
        contentMaxResults: { type: "integer", minimum: 0, maximum: 100, default: 25 },
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
        maxPages: { type: "integer", minimum: 1, maximum: 100, description: "Optional maximum number of Microsoft Graph delta pages to fetch in this call. Returns the advanced nextLink or terminal deltaLink when reached." },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_sync_status",
    description: "Show local metadata-cache status, delta cursor availability, and plugin storage locations.",
    inputSchema: {
      type: "object",
      properties: {
        includeSamples: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_cache_refresh",
    description: "Refresh the local metadata cache from a bounded recursive scan, or from delta when a previous cursor exists.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        mode: { type: "string", enum: ["auto", "scan", "delta"], default: "auto" },
        pageSize: { type: "integer", minimum: 1, maximum: 200, default: 200 },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        includeFolders: { type: "boolean", default: true },
        replaceCache: { type: "boolean", default: false, description: "When true, clear existing metadata before the scan. Defaults to merge so bounded scans do not shrink a fuller cache." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_cache_clear",
    description: "Clear the local OneDrive metadata cache.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_content_index_refresh",
    description: "Build or refresh the optional local text content index from cached metadata or one selected file. This is the explicit expensive content-reading step.",
    inputSchema: {
      type: "object",
      properties: {
        ...pathTargetProperties,
        refreshMetadata: {
          type: "boolean",
          default: false,
          description: "When true, rebuild metadata with onedrive_cache_refresh before indexing cached files."
        },
        metadataMode: { type: "string", enum: ["auto", "delta", "scan"], default: "auto", description: "Metadata refresh strategy. Auto prefers the stored delta cursor and falls back to a bounded scan." },
        maxFiles: { type: "integer", minimum: 1, maximum: 1000, default: 100 },
        maxBytesPerFile: { type: "integer", minimum: 1024, maximum: 5242880, default: defaultMaxIndexedFileSize },
        concurrencyLimit: { type: "integer", minimum: 1, maximum: 8, default: 2 },
        extensions: {
          type: "array",
          items: { type: "string" },
          description: "Optional extension allow-list. Defaults to configured supported indexed file types."
        },
        includeOfficeExport: {
          type: "boolean",
          default: false,
          description: "When true, try Microsoft Graph text export for Office-like files. Unsupported file types fail per item, not the whole refresh."
        },
        includeOfficeStructured: { type: "boolean", default: true, description: "Index structured Word, Excel, and PowerPoint Open XML content with semantic anchors." },
        maxOfficeSegments: { type: "integer", minimum: 1, maximum: 100000, default: 50000 },
        maxOfficePackageBytes: { type: "integer", minimum: 1024, maximum: 262144000, default: 52428800 },
        force: {
          type: "boolean",
          default: false,
          description: "Re-extract unchanged files instead of reusing index entries with matching ETag/cTag/mtime/size."
        },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_content_search",
    description: "Search the optional local text content index and return lightweight metadata plus snippets. Does not call Microsoft Graph.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        maxResults: { type: "integer", minimum: 1, maximum: 100, default: 20 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_index_refresh",
    description: "Incrementally refresh the structured local Office index using item identity and eTag/cTag freshness, preferring OneDrive delta metadata updates.",
    inputSchema: {
      type: "object",
      properties: {
        ...pathTargetProperties,
        refreshMetadata: { type: "boolean", default: true },
        metadataMode: { type: "string", enum: ["auto", "delta", "scan"], default: "auto" },
        maxFiles: { type: "integer", minimum: 1, maximum: 1000, default: 100 },
        concurrencyLimit: { type: "integer", minimum: 1, maximum: 8, default: 2 },
        force: { type: "boolean", default: false },
        maxOfficeSegments: { type: "integer", minimum: 1, maximum: 100000, default: 50000 },
        maxOfficePackageBytes: { type: "integer", minimum: 1024, maximum: 262144000, default: 52428800 },
        maxOfficeParagraphs: { type: "integer", minimum: 1, maximum: 10000, default: 10000 },
        maxOfficeCells: { type: "integer", minimum: 1, maximum: 50000, default: 50000 },
        maxOfficeSlides: { type: "integer", minimum: 1, maximum: 5000, default: 5000 },
        scanMaxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        scanMaxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        scanMaxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_search",
    description: "Search the structured local Office index and return semantic paragraph, cell, formula, table, shape, slide-note, content-control, and comment anchors without Graph calls.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 },
        maxResults: { type: "integer", minimum: 1, maximum: 100, default: 20 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_content_index_clear",
    description: "Clear the optional local OneDrive content index.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_office_capabilities",
    description: "Report available native Office editing backends, runtime health, account limitations, and supported Open XML formats.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_office_validate",
    description: "Download and safely validate a Word, Excel, or PowerPoint Open XML package without changing the remote file.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        expectedKind: { type: "string", enum: ["word", "excel", "powerpoint"] },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_word_get_document",
    description: "Read structured paragraphs, tables, content controls, headers, footers, comments, and package safety metadata from a .docx file.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        ...officeStructuredSearchProperties,
        maxParagraphs: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        includeHeadersFooters: { type: "boolean", default: true },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_excel_get_workbook",
    description: "Read structured worksheets, bounded cell/range selections, searchable values/formulas, table metadata, styles, and defined names from an .xlsx file using the local Open XML backend.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        ...officeStructuredSearchProperties,
        includeCells: { type: "boolean", default: true },
        includeTables: { type: "boolean", default: true },
        includeCharts: { type: "boolean", default: true },
        includePivots: { type: "boolean", default: true },
        includeFormulaDependencies: { type: "boolean", default: false },
        sheetNames: { type: "array", minItems: 1, maxItems: 100, uniqueItems: true, items: { type: "string", minLength: 1 }, description: "Optional worksheet-name selector." },
        address: { type: "string", pattern: "^[A-Za-z]{1,3}[1-9][0-9]*(?::[A-Za-z]{1,3}[1-9][0-9]*)?$", description: "Optional bounded A1 cell or range selector applied to selected worksheets." },
        maxCells: { type: "integer", minimum: 1, maximum: 50000, default: 5000 },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_powerpoint_get_presentation",
    description: "Read structured slides, shapes, text, placeholders, geometry, tables, images, notes, and package safety metadata from a .pptx file.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...officeTargetProperties,
        ...officeStructuredSearchProperties,
        maxSlides: { type: "integer", minimum: 1, maximum: 5000, default: 500 },
        strictRelationships: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_word_batch_update",
    description: "Preview or apply typed Word edits for text, paragraphs, styles, tables, content controls, external hyperlinks, and anchored comments. Documents containing tracked changes are refused; live commits require expected identity and the dry-run preview token.",
    inputSchema: {
      type: "object",
      required: ["operations"],
      anyOf: itemTargetAnyOf,
      properties: { ...officeBatchCommonProperties, operations: wordOperationsSchema, trackedChanges: { type: "string", enum: ["refuse"], default: "refuse", description: "Tracked-change markup is currently refused because native edits cannot safely preserve review semantics." } },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_excel_batch_update",
    description: "Preview or apply typed Excel cell/range, formatting, validation, pane, table-row, chart, sheet-name, defined-name, and recalculation edits. Auto uses Graph sessions only when every operation is supported for a business .xlsx file and Open XML otherwise.",
    inputSchema: {
      type: "object",
      required: ["operations"],
      anyOf: itemTargetAnyOf,
      properties: { ...officeBatchCommonProperties, operations: excelOperationsSchema, backend: { type: "string", enum: ["auto", "openxml", "graph"], default: "auto" } },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_powerpoint_batch_update",
    description: "Preview or apply typed PowerPoint text, text-box, text-style, shape deletion, geometry, image replacement, table, notes, duplicate, delete, and move-slide edits; live commits require a preview token.",
    inputSchema: {
      type: "object",
      required: ["operations"],
      anyOf: itemTargetAnyOf,
      properties: { ...officeBatchCommonProperties, operations: powerpointOperationsSchema },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_batch_transform",
    description: "Preflight and apply one atomic-plan Office transformation across up to 25 files. Every file is fully previewed before the first mutation; live partial state includes recovery backups and remaining items.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 25,
          items: {
            type: "object",
            required: ["kind", "operations"],
            anyOf: itemTargetAnyOf,
            properties: {
              ...officeTargetProperties,
              kind: { type: "string", enum: ["word", "excel", "powerpoint"] },
              operations: { anyOf: [wordOperationsSchema, excelOperationsSchema, powerpointOperationsSchema] },
              expectedName: { type: "string" },
              expectedId: { type: "string" },
              backend: { type: "string", enum: ["auto", "openxml", "graph"], default: "auto" },
              allowMacros: { type: "boolean", default: false }
            },
            additionalProperties: false
          }
        },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        previewToken: previewTokenSchema,
        createBackup: { type: "boolean", default: true },
        verify: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_backups",
    description: "List plugin-managed Office backups with stable backup IDs and original remote item metadata.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", minLength: 1 },
        kind: { type: "string", enum: ["word", "excel", "powerpoint"] },
        limit: { type: "integer", minimum: 1, maximum: 500, default: 100 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_compare_backup",
    description: "Compare one plugin-managed Office backup with the current content of its original stable remote item ID.",
    inputSchema: {
      type: "object",
      required: ["backupId"],
      properties: { backupId: { type: "string", pattern: "^[0-9a-fA-F-]{36}$" }, maxChanges: { type: "integer", minimum: 1, maximum: 500, default: 100 } },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_office_restore_backup",
    description: "Preview or restore an Office backup to its original stable item ID. Live restore requires expected ID/eTag and the dry-run preview token.",
    inputSchema: {
      type: "object",
      required: ["backupId"],
      properties: {
        backupId: { type: "string", pattern: "^[0-9a-fA-F-]{36}$" },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        expectedId: { type: "string" },
        expectedETag: { type: "string" },
        previewToken: previewTokenSchema,
        verify: { type: "boolean", default: true }
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
    name: "onedrive_preview",
    description: "Return a bounded safe preview for text files and Graph-supported document text exports.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        maxBytes: { type: "integer", minimum: 1, maximum: 1048576, default: 65536 },
        preferExportText: {
          type: "boolean",
          default: true,
          description: "For likely Office files, ask Graph for a text export before falling back to raw text checks."
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
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
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
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
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
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
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
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_export_pdf",
    description: "Ask Microsoft Graph to export a supported OneDrive document to PDF and save it locally.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/export/<name>.pdf." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_export_text",
    description: "Ask Microsoft Graph to export a supported OneDrive document to plain text and save it locally when Graph supports that conversion.",
    inputSchema: {
      type: "object",
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        localPath: { type: "string", description: "Optional destination path. Defaults to ~/.codex/onedrive-plugin/downloads/export/<name>.txt." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder. Prefer remote plugin operations instead."
        }
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
        dryRun: { type: "boolean", default: true, description: "When replacing an existing file, return a guarded preview unless explicitly set to false." },
        confirmed: { type: "boolean", default: false },
        expectedName: { type: "string", description: "Required identity check when replacing an existing file." },
        expectedId: { type: "string", description: "Required identity check when replacing an existing file." },
        previewToken: previewTokenSchema,
        uploadMode: { type: "string", enum: ["auto", "simple", "session"], default: "auto" },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to read from a locally synced OneDrive folder. Prefer remote plugin operations instead."
        },
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
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" },
        dryRun: { type: "boolean", default: true, description: "When replacing an existing file, return a guarded preview unless explicitly set to false." },
        confirmed: { type: "boolean", default: false },
        expectedName: { type: "string", description: "Required identity check when replacing an existing file." },
        expectedId: { type: "string", description: "Required identity check when replacing an existing file." },
        previewToken: previewTokenSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_patch_text",
    description: "Preview or conditionally apply a unified, JSON, safe-YAML, or CSV patch while preserving supported encoding and newline style.",
    inputSchema: {
      type: "object", anyOf: itemTargetAnyOf, required: ["patch"],
      properties: {
        ...pathTargetProperties,
        patch: { type: "object", required: ["mode"], properties: {
          mode: { type: "string", enum: ["unified", "json", "yaml", "csv"] }, diff: { type: "string" },
          operations: { type: "array", minItems: 1, maxItems: 1000, items: structuredTextPatchOperationSchema },
          indent: { type: "integer", minimum: 0, maximum: 8 }, keyColumn: { type: "string" }, delimiter: { type: "string", minLength: 1, maxLength: 1 }
        }, additionalProperties: false },
        dryRun: { type: "boolean", default: true }, confirmed: { type: "boolean", default: false }, expectedName: { type: "string" }, expectedId: { type: "string" }, expectedETag: { type: "string" }, previewToken: previewTokenSchema
      }, additionalProperties: false
    }
  },
  {
    name: "onedrive_versions",
    description: "List bounded OneDrive version-history metadata for one file.",
    inputSchema: { type: "object", anyOf: itemTargetAnyOf, properties: { ...pathTargetProperties, maxItems: { type: "integer", minimum: 1, maximum: 200, default: 50 } }, additionalProperties: false }
  },
  {
    name: "onedrive_compare_version",
    description: "Compare a historical OneDrive version with the current file or another historical version using semantic Office, text, or binary evidence.",
    inputSchema: { type: "object", anyOf: itemTargetAnyOf, required: ["versionId"], properties: { ...pathTargetProperties, versionId: { type: "string" }, compareToVersionId: { type: "string" }, maxChanges: { type: "integer", minimum: 1, maximum: 1000, default: 200 } }, additionalProperties: false }
  },
  {
    name: "onedrive_restore_version",
    description: "Preview or natively restore one historical OneDrive version as the current version with identity, eTag, and preview-token guards.",
    inputSchema: { type: "object", anyOf: itemTargetAnyOf, required: ["versionId"], properties: { ...pathTargetProperties, versionId: { type: "string" }, dryRun: { type: "boolean", default: true }, confirmed: { type: "boolean", default: false }, expectedName: { type: "string" }, expectedId: { type: "string" }, expectedETag: { type: "string" }, previewToken: previewTokenSchema }, additionalProperties: false }
  },
  {
    name: "onedrive_workspace_list",
    description: "List auth-context and drive-scoped managed edit workspaces.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "onedrive_workspace_create",
    description: "Preview or create an owner-only remote draft workspace for one file.",
    inputSchema: { type: "object", anyOf: itemTargetAnyOf, properties: { ...pathTargetProperties, dryRun: { type: "boolean", default: true }, confirmed: { type: "boolean", default: false }, expectedName: { type: "string" }, expectedId: { type: "string" }, expectedETag: { type: "string" }, previewToken: previewTokenSchema }, additionalProperties: false }
  },
  {
    name: "onedrive_workspace_status",
    description: "Return source/draft drift, semantic comparison, and promotion readiness for a managed edit workspace.",
    inputSchema: { type: "object", required: ["workspaceId"], properties: { workspaceId: { type: "string" }, maxChanges: { type: "integer", minimum: 1, maximum: 1000, default: 200 } }, additionalProperties: false }
  },
  {
    name: "onedrive_workspace_promote",
    description: "Preview or promote a verified draft to its original stable item ID, then remove the successful draft.",
    inputSchema: { type: "object", required: ["workspaceId"], properties: { workspaceId: { type: "string" }, dryRun: { type: "boolean", default: true }, confirmed: { type: "boolean", default: false }, expectedId: { type: "string" }, expectedETag: { type: "string" }, previewToken: previewTokenSchema }, additionalProperties: false }
  },
  {
    name: "onedrive_workspace_abandon",
    description: "Preview or abandon one managed remote edit workspace and delete only its draft artifacts.",
    inputSchema: { type: "object", required: ["workspaceId"], properties: { workspaceId: { type: "string" }, dryRun: { type: "boolean", default: true }, confirmed: { type: "boolean", default: false }, expectedId: { type: "string" }, expectedETag: { type: "string" }, previewToken: previewTokenSchema }, additionalProperties: false }
  },
  {
    name: "onedrive_watch_start",
    description: "Start or resume a scoped local delta watch with bounded polling, expiry, and stale-preview invalidation.",
    inputSchema: { type: "object", properties: { ...folderTargetProperties, intervalSeconds: { type: "integer", minimum: 15, maximum: 300, default: 30 }, expiresInSeconds: { type: "integer", minimum: 60, maximum: 28800, default: 3600 } }, additionalProperties: false }
  },
  {
    name: "onedrive_watch_status",
    description: "Report active watches and bounded recent events.",
    inputSchema: { type: "object", properties: { watchId: { type: "string" }, maxEvents: { type: "integer", minimum: 0, maximum: 500, default: 100 }, consume: { type: "boolean", default: false } }, additionalProperties: false }
  },
  {
    name: "onedrive_watch_stop",
    description: "Stop one local delta watch and retain its final bounded status.",
    inputSchema: { type: "object", required: ["watchId"], properties: { watchId: { type: "string" } }, additionalProperties: false }
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
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this renames a file or folder."
        },
        expectedName: { type: "string", description: "Required for a live rename unless expectedId is provided; item name must match. Optional for dry-run." },
        expectedId: { type: "string", description: "Required for a live rename unless expectedName is provided; item ID must match. Optional for dry-run." },
        previewToken: previewTokenSchema
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
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this moves a file or folder."
        },
        expectedName: { type: "string", description: "Required for a live move unless expectedId is provided; source item name must match. Optional for dry-run." },
        expectedId: { type: "string", description: "Required for a live move unless expectedName is provided; source item ID must match. Optional for dry-run." },
        previewToken: previewTokenSchema
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
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this copies a file or folder."
        },
        waitForCompletion: { type: "boolean", default: false },
        timeoutSeconds: { type: "integer", minimum: 1, maximum: 300, default: 60 },
        expectedName: { type: "string", description: "Required for a live copy unless expectedId is provided; source item name must match. Optional for dry-run." },
        expectedId: { type: "string", description: "Required for a live copy unless expectedName is provided; source item ID must match. Optional for dry-run." },
        previewToken: previewTokenSchema
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
        password: { type: "string", minLength: 1, description: "Optional sharing-link password. Never returned in dry-runs or audit logs." },
        expirationDateTime: { type: "string", minLength: 1, description: "Optional ISO 8601 link expiration timestamp." },
        retainInheritedPermissions: { type: "boolean" },
        includePermissionDiff: {
          type: "boolean",
          default: true,
          description: "When true, include before-permissions in dry-run and before/after permission diff for live creation."
        },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this can expose access to a file or folder."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_invite_permission",
    description: "Invite users or groups to a OneDrive item. Defaults to dry-run and silent direct grants unless sendInvitation=true is set.",
    inputSchema: {
      type: "object",
      required: ["recipients"],
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        recipients: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: driveRecipientSchema,
          description: "Recipients to grant access to. Each item must include exactly one of email, alias, or objectId."
        },
        role: { type: "string", enum: ["read", "write"], default: "read" },
        sendInvitation: { type: "boolean", default: false },
        requireSignIn: { type: "boolean", default: true },
        message: { type: "string", description: "Optional invitation message. Only useful when sendInvitation=true; never written to audit logs." },
        password: { type: "string", minLength: 1, description: "Optional invite password when supported by Microsoft Graph. Never returned in dry-runs or audit logs." },
        expirationDateTime: { type: "string", minLength: 1, description: "Optional ISO 8601 expiration timestamp." },
        retainInheritedPermissions: { type: "boolean" },
        includePermissionDiff: {
          type: "boolean",
          default: true,
          description: "When true, include before-permissions in dry-run and before/after permission diff for live invitation."
        },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this grants access to a file or folder."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_revoke_permission",
    description: "Preview or revoke a OneDrive sharing permission. Defaults to dry-run and requires confirmation plus expected item identity for live revocation.",
    inputSchema: {
      type: "object",
      required: ["permissionId"],
      anyOf: itemTargetAnyOf,
      properties: {
        ...pathTargetProperties,
        permissionId: { type: "string", minLength: 1 },
        includePermissions: {
          type: "boolean",
          default: true,
          description: "When true, include before permissions in dry-run and before/after permission diff for live revocation."
        },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation because this removes sharing access."
        },
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Optional safety check: item name must match." },
        expectedId: { type: "string", description: "Optional safety check: item ID must match." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_revoke_permissions",
    description: "Preview or revoke multiple OneDrive sharing permissions. Live revoke preflights every item before any DELETE.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            required: ["permissionId"],
            anyOf: itemTargetAnyOf,
            properties: {
              ...pathTargetProperties,
              permissionId: { type: "string", minLength: 1 },
              expectedName: { type: "string" },
              expectedId: { type: "string" }
            },
            additionalProperties: false
          }
        },
        includePermissions: { type: "boolean", default: true },
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation for live batch permission revocation."
        },
        previewToken: previewTokenSchema
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
        previewToken: previewTokenSchema,
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
    name: "onedrive_batch_get_info",
    description: "Get metadata for up to 20 OneDrive items using Microsoft Graph batching.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 20,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: pathTargetProperties,
            additionalProperties: false
          }
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_permissions",
    description: "Audit sharing permissions for up to 20 OneDrive files or folders using Microsoft Graph batching.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 20,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: pathTargetProperties,
            additionalProperties: false
          }
        },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_download",
    description: "Download multiple OneDrive files serially with one result per item.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: {
              ...pathTargetProperties,
              localPath: { type: "string" },
              overwrite: { type: "boolean" }
            },
            additionalProperties: false
          }
        },
        destinationFolder: { type: "string", description: "Optional local folder for downloaded files." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_delete",
    description: "Preview or delete multiple OneDrive items. Live deletes require dryRun=false and confirmed=true.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: {
              ...pathTargetProperties,
              expectedName: { type: "string" },
              expectedId: { type: "string" }
            },
            additionalProperties: false
          }
        },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        previewToken: previewTokenSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_batch_move",
    description: "Preview or move multiple OneDrive items to one destination folder.",
    inputSchema: {
      type: "object",
      required: ["items"],
      properties: {
        items: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            anyOf: itemTargetAnyOf,
            properties: {
              path: { type: "string", description: "Source item path relative to OneDrive root." },
              itemId: { type: "string", description: "Source drive item ID." },
              preset: presetSchema,
              relativePath: relativePathSchema,
              newName: { type: "string" },
              expectedName: { type: "string" },
              expectedId: { type: "string" }
            },
            additionalProperties: false
          }
        },
        destinationParentPath: { type: "string", description: "Destination folder path relative to root. Omit or use / for root." },
        destinationParentItemId: { type: "string", description: "Destination folder drive item ID." },
        ...destinationPresetProperties,
        dryRun: { type: "boolean", default: true },
        confirmed: {
          type: "boolean",
          default: false,
          description: "Must be true after explicit user confirmation for live batch moves."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_update_file",
    description: "Safe file update workflow: checkout downloads a remote file with a manifest; commit uploads an edited file back after conflict checks.",
    inputSchema: {
      type: "object",
      required: ["mode", "remotePath"],
      properties: {
        mode: { type: "string", enum: ["checkout", "commit"] },
        remotePath: { type: "string", description: "Remote OneDrive path including filename." },
        itemId: { type: "string", description: "Optional item ID for checkout." },
        localPath: { type: "string", description: "Checkout destination or commit source. Checkout defaults to ~/.codex/onedrive-plugin/updates." },
        manifestPath: { type: "string", description: "Optional checkout manifest path. Defaults to localPath + .onedrive-update.json." },
        conflictCheck: { type: "boolean", default: true },
        force: { type: "boolean", default: false },
        createBackup: { type: "boolean", default: true },
        verify: { type: "boolean", default: true },
        overwriteLocal: { type: "boolean", default: false },
        overwriteManifest: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: { type: "boolean", default: false }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_recent",
    description: "List recently used OneDrive files from Microsoft Graph.",
    inputSchema: {
      type: "object",
      properties: {
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_large_files",
    description: "Find large files by scanning OneDrive and sorting by size.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        minBytes: { type: "integer", minimum: 0, default: 104857600 },
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_duplicates",
    description: "Find likely duplicate files by scanning OneDrive and grouping by hash when available, otherwise by name and size.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 50000, default: 10000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 10000, default: 2000 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 25 },
        format: outputFormatSchema
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_shared_by_me",
    description: "Scan files and folders and return items with explicit sharing permissions.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 1000, default: 250 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 15 },
        includeFolders: { type: "boolean", default: true },
        includeOwnerPermissions: { type: "boolean", default: false, description: "When true, include owner/self grants in addition to external sharing permissions." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_public_links",
    description: "Scan files and folders and return anonymous sharing links.",
    inputSchema: {
      type: "object",
      properties: {
        ...folderTargetProperties,
        limit: { type: "integer", minimum: 1, maximum: 200, default: 50 },
        maxItems: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
        maxFolders: { type: "integer", minimum: 1, maximum: 1000, default: 250 },
        maxDepth: { type: "integer", minimum: 0, maximum: 50, default: 15 },
        includeFolders: { type: "boolean", default: true }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_audit_recent",
    description: "Read recent local OneDrive live-mutation audit log entries.",
    inputSchema: {
      type: "object",
      properties: {
        limit: { type: "integer", minimum: 1, maximum: 500, default: 50 },
        tool: { type: "string", description: "Optional exact tool name filter, such as onedrive_upload." },
        status: { type: "string", description: "Optional mutation status filter, such as success or failed." },
        pathContains: { type: "string", description: "Optional case-insensitive filter matched against target, before, and after paths/names." },
        since: { type: "string", description: "Optional ISO timestamp; only entries at or after this time are returned." },
        until: { type: "string", description: "Optional ISO timestamp; only entries at or before this time are returned." },
        newestFirst: { type: "boolean", default: true, description: "Return newest entries first. Set false to preserve log order." }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_audit_export",
    description: "Export the local OneDrive live-mutation audit log to a JSONL file.",
    inputSchema: {
      type: "object",
      properties: {
        localPath: { type: "string", description: "Optional export path. Defaults to ~/.codex/onedrive-plugin/audit/export-<timestamp>.jsonl." },
        overwrite: { type: "boolean", default: false },
        allowLocalOneDriveSyncPath: {
          type: "boolean",
          default: false,
          description: "Explicit override to write into a locally synced OneDrive folder."
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_audit_clear",
    description: "Clear the local OneDrive live-mutation audit log. Requires confirmed=true.",
    inputSchema: {
      type: "object",
      properties: {
        confirmed: { type: "boolean", default: false }
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
        previewToken: previewTokenSchema,
        expectedName: { type: "string", description: "Required for a live delete unless expectedId is provided; item name must match. Optional for dry-run." },
        expectedId: { type: "string", description: "Required for a live delete unless expectedName is provided; item ID must match. Optional for dry-run." }
      },
      additionalProperties: false
    }
  }
];

// ChatGPT and company-knowledge surfaces give special treatment to the
// standard search/fetch contract and can pass uploaded-file references through
// file parameters. Keep these ChatGPT-specific tools out of the immutable
// 84-tool local/Codex contract, but advertise and execute them in the focused
// profile so ordinary retrieval avoids the exhaustive finder path.
const chatgptCompatibilityTools = [
  {
    name: "search",
    title: "Search OneDrive",
    description: "Find OneDrive files or folders by name, topic, aliases, or indexed content, including records saved under an unknown contractor or work-order name. Every result includes the opaque ID required by fetch.",
    inputSchema: {
      type: "object",
      required: ["query"],
      properties: {
        query: { type: "string", minLength: 1 }
      },
      additionalProperties: false
    },
    outputSchema: {
      type: "object",
      required: ["results"],
      properties: {
        results: {
          type: "array",
          maxItems: 10,
          items: {
            type: "object",
            required: ["id", "title", "url"],
            properties: {
              id: { type: "string" },
              title: { type: "string" },
              url: { type: "string" }
            },
            additionalProperties: false
          }
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "fetch",
    title: "Fetch OneDrive item",
    description: "Use this when the user wants to read an item returned by search. Pass the returned ID unchanged; a continuation ID appears only when a large file needs another fetch.",
    inputSchema: {
      type: "object",
      required: ["id"],
      properties: {
        id: { type: "string", minLength: 1 }
      },
      additionalProperties: false
    },
    outputSchema: {
      type: "object",
      required: ["id", "title", "text", "url"],
      properties: {
        id: { type: "string" },
        title: { type: "string" },
        text: { type: "string" },
        url: { type: "string" },
        metadata: {
          type: "object",
          additionalProperties: { type: "string" }
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_open_files",
    title: "Open exact OneDrive files",
    description: "Read one or more specifically named OneDrive files in one bounded call.",
    inputSchema: {
      type: "object",
      required: ["names"],
      properties: {
        names: {
          type: "array",
          minItems: 1,
          maxItems: 5,
          uniqueItems: true,
          items: { type: "string", minLength: 1 },
          description: "Exact filenames, including extensions, to locate and read."
        }
      },
      additionalProperties: false
    },
    outputSchema: {
      type: "object",
      required: ["files", "durationMs"],
      properties: {
        files: {
          type: "array",
          maxItems: 5,
          items: {
            type: "object",
            required: ["name", "status", "durationMs"],
            properties: {
              name: { type: "string" },
              status: { type: "string", enum: ["found", "not_found", "ambiguous", "error"] },
              id: { type: "string" },
              title: { type: "string" },
              text: { type: "string" },
              url: { type: "string" },
              metadata: { type: "object", additionalProperties: { type: "string" } },
              error: { type: "string" },
              candidates: {
                type: "array",
                maxItems: 3,
                items: {
                  type: "object",
                  required: ["id", "title", "url"],
                  properties: { id: { type: "string" }, title: { type: "string" }, url: { type: "string" } },
                  additionalProperties: false
                }
              },
              durationMs: { type: "integer", minimum: 0 }
            },
            additionalProperties: false
          }
        },
        durationMs: { type: "integer", minimum: 0 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_preview_actions",
    title: "Preview OneDrive actions",
    description: "Preview one or more OneDrive item or sharing changes without modifying data or access.",
    inputSchema: {
      type: "object",
      required: ["actions"],
      properties: {
        actions: {
          type: "array",
          minItems: 1,
          maxItems: 10,
          items: {
            type: "object",
            required: ["operation", "itemId"],
            properties: {
              operation: { type: "string", enum: ["rename", "move", "copy", "createSharingLink", "revokePermission"] },
              itemId: { type: "string", minLength: 1 },
              newName: { type: "string", minLength: 1 },
              destinationParentPath: { type: "string" },
              destinationParentItemId: { type: "string", minLength: 1 },
              linkType: { type: "string", enum: ["view", "edit", "embed"], default: "view" },
              scope: { type: "string", enum: ["anonymous", "organization", "users"], default: "anonymous" },
              permissionId: { type: "string", minLength: 1 }
            },
            additionalProperties: false
          }
        }
      },
      additionalProperties: false
    },
    outputSchema: {
      type: "object",
      required: ["dryRun", "results", "durationMs"],
      properties: {
        dryRun: { type: "boolean" },
        results: {
          type: "array",
          maxItems: 10,
          items: {
            type: "object",
            required: ["index", "operation", "itemId", "isError", "dryRun", "previewTokenPresent", "durationMs"],
            properties: {
              index: { type: "integer", minimum: 0 },
              operation: { type: "string" },
              itemId: { type: "string" },
              expectedId: { type: "string" },
              isError: { type: "boolean" },
              dryRun: { type: "boolean" },
              previewTokenPresent: { type: "boolean" },
              previewToken: { type: "string" },
              previewTokenExpiresAt: { type: "string" },
              summary: { type: "string" },
              error: { type: "string" },
              accessSummary: {
                type: "object",
                required: ["permissionCount", "sharingLinkCount", "anonymousLinkCount", "roles"],
                properties: {
                  permissionCount: { type: "integer", minimum: 0 },
                  sharingLinkCount: { type: "integer", minimum: 0 },
                  anonymousLinkCount: { type: "integer", minimum: 0 },
                  roles: { type: "array", items: { type: "string" } }
                },
                additionalProperties: false
              },
              accessSummaryError: { type: "string" },
              durationMs: { type: "integer", minimum: 0 }
            },
            additionalProperties: false
          }
        },
        durationMs: { type: "integer", minimum: 0 }
      },
      additionalProperties: false
    }
  },
  {
    name: "onedrive_upload_file",
    title: "Upload File to OneDrive",
    description: "Preview, then upload a ChatGPT-provided file to OneDrive. Every upload requires confirmation and a scoped preview token; replacement also requires the existing item's expected identity.",
    inputSchema: {
      type: "object",
      required: ["sourceFile", "remotePath"],
      properties: {
        sourceFile: {
          type: "object",
          required: ["download_url", "file_id"],
          properties: {
            download_url: { type: "string", minLength: 1 },
            file_id: { type: "string", minLength: 1 },
            mime_type: { type: "string" },
            file_name: { type: "string" }
          },
          additionalProperties: false
        },
        remotePath: { type: "string", minLength: 1, description: "Destination path relative to the OneDrive root, including filename." },
        conflictBehavior: { type: "string", enum: ["fail", "replace", "rename"], default: "fail" },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        expectedName: { type: "string", description: "Required when replacing an existing file unless expectedId is provided." },
        expectedId: { type: "string", description: "Required when replacing an existing file unless expectedName is provided." },
        previewToken: previewTokenSchema
      },
      additionalProperties: false
    },
    _meta: {
      "openai/fileParams": ["sourceFile"]
    }
  },
  {
    name: "onedrive_permanent_delete",
    title: "Permanently Delete OneDrive Item",
    description: "Permanently delete a OneDrive item without using the recycle bin. Irreversible; defaults to preview and requires confirmation, expected identity, an acknowledgement, and a scoped preview token.",
    inputSchema: {
      type: "object",
      required: ["itemId"],
      properties: {
        itemId: { type: "string", minLength: 1 },
        dryRun: { type: "boolean", default: true },
        confirmed: { type: "boolean", default: false },
        acknowledgeIrreversible: { type: "boolean", default: false },
        expectedName: { type: "string" },
        expectedId: { type: "string" },
        previewToken: previewTokenSchema
      },
      additionalProperties: false
    }
  }
];

// ChatGPT developer-mode apps reject MCP tools that omit impact annotations.
// Keep these classifications centralized so every descriptor carries the three
// required hints and newly added tools fail closed during packaging.
//
// Auth is selected at process start. Local/Codex and legacy ChatGPT deployments
// default to noauth and use the encrypted device-code credential. Work-compatible
// HTTP deployments set ONEDRIVE_MCP_AUTH_MODE=oauth; each tool then advertises the
// Entra API scope and Graph calls use the request's on-behalf-of token.
const configuredSecuritySchemes = Object.freeze(toolSecuritySchemes());
const readOnlyToolNames = new Set([
  "search",
  "fetch",
  "onedrive_open_files",
  "onedrive_preview_actions",
  "onedrive_config",
  "onedrive_me",
  "onedrive_drive",
  "onedrive_doctor",
  "onedrive_presets",
  "onedrive_list",
  "onedrive_list_all",
  "onedrive_scan",
  "onedrive_search",
  "onedrive_search_all",
  "onedrive_find",
  "onedrive_find_all",
  "onedrive_delta",
  "onedrive_sync_status",
  "onedrive_content_search",
  "onedrive_office_search",
  "onedrive_office_capabilities",
  "onedrive_office_validate",
  "onedrive_word_get_document",
  "onedrive_excel_get_workbook",
  "onedrive_powerpoint_get_presentation",
  "onedrive_office_backups",
  "onedrive_office_compare_backup",
  "onedrive_get_info",
  "onedrive_read_text",
  "onedrive_preview",
  "onedrive_versions",
  "onedrive_compare_version",
  "onedrive_workspace_list",
  "onedrive_workspace_status",
  "onedrive_watch_status",
  "onedrive_permissions",
  "onedrive_batch_get_info",
  "onedrive_batch_permissions",
  "onedrive_recent",
  "onedrive_large_files",
  "onedrive_duplicates",
  "onedrive_shared_by_me",
  "onedrive_public_links",
  "onedrive_audit_recent"
]);

const destructiveToolNames = new Set([
  "onedrive_logout",
  "onedrive_cache_clear",
  "onedrive_content_index_clear",
  "onedrive_word_batch_update",
  "onedrive_excel_batch_update",
  "onedrive_powerpoint_batch_update",
  "onedrive_office_batch_transform",
  "onedrive_office_restore_backup",
  "onedrive_download",
  "onedrive_download_excel",
  "onedrive_download_word",
  "onedrive_download_powerpoint",
  "onedrive_export_pdf",
  "onedrive_export_text",
  "onedrive_write_text",
  "onedrive_patch_text",
  "onedrive_restore_version",
  "onedrive_workspace_promote",
  "onedrive_workspace_abandon",
  "onedrive_revoke_permission",
  "onedrive_batch_revoke_permissions",
  "onedrive_batch_download",
  "onedrive_batch_delete",
  "onedrive_update_file",
  "onedrive_audit_export",
  "onedrive_audit_clear",
  "onedrive_delete",
  "onedrive_permanent_delete",
  "onedrive_upload_file"
]);

const openWorldToolNames = new Set([
  "onedrive_download",
  "onedrive_download_excel",
  "onedrive_download_word",
  "onedrive_download_powerpoint",
  "onedrive_export_pdf",
  "onedrive_export_text",
  "onedrive_invite_permission",
  "onedrive_create_sharing_link",
  "onedrive_batch_download",
  "onedrive_update_file",
  "onedrive_audit_export",
  "onedrive_upload_file"
]);

function toolTitle(name) {
  return name
    .replace(/^onedrive_/, "OneDrive ")
    .split("_")
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

const executableTools = [...tools, ...chatgptCompatibilityTools];

for (const tool of executableTools) {
  tool.title = tool.title || toolTitle(tool.name);
  tool.securitySchemes = configuredSecuritySchemes;
  tool._meta = {
    ...(tool._meta || {}),
    securitySchemes: configuredSecuritySchemes
  };
  tool.annotations = {
    readOnlyHint: readOnlyToolNames.has(tool.name),
    openWorldHint: openWorldToolNames.has(tool.name),
    destructiveHint: destructiveToolNames.has(tool.name)
  };
}

// ChatGPT sends the advertised tool contract through the model's selection
// path on every turn. The full OneDrive surface is intentionally broad, but
// exposing every maintenance, batch, watch, and workspace helper makes simple
// requests pay for 300+ KB of schemas before Graph is called. Keep the full
// contract for Codex and local automation while offering a focused ChatGPT
// profile for tunnel deployments.
const chatgptToolNames = new Set([
  "search",
  "fetch",
  "onedrive_open_files",
  "onedrive_preview_actions",
  "onedrive_list",
  "onedrive_office_capabilities",
  "onedrive_office_batch_transform",
  "onedrive_upload_file",
  "onedrive_write_text",
  "onedrive_patch_text",
  "onedrive_create_folder",
  "onedrive_rename",
  "onedrive_move",
  "onedrive_copy",
  "onedrive_create_sharing_link",
  "onedrive_invite_permission",
  "onedrive_revoke_permission",
  "onedrive_permissions",
  "onedrive_delete",
  "onedrive_restore_deleted",
  "onedrive_permanent_delete"
]);

const compactOfficeOperationSchema = {
  type: "array",
  minItems: 1,
  maxItems: 100,
  items: {
    type: "object",
    required: ["type"],
    properties: {
      type: { type: "string", minLength: 1 }
    },
    additionalProperties: true
  },
  description: "Typed Office operations. Use fetch to read the document first; the server validates every operation against the full schema."
};

const chatgptToolMetadata = Object.freeze({
  search: {
    description: "Use this when the user wants OneDrive discovery from a natural-language or subtle description, topic, partial name, keywords, aliases, indexed content, or an unknown title. Pass the intent once; bounded concept inference handles common document and subject aliases. For exact filenames plus content use onedrive_open_files; otherwise pass the chosen opaque id unchanged to fetch.",
    invoking: "Searching OneDrive…",
    invoked: "OneDrive results ready"
  },
  fetch: {
    description: "Use this when the user wants to read an item returned by search. Pass the returned id unchanged; it extracts bounded content and returns a continuation ID only when more detail is needed.",
    invoking: "Reading OneDrive item…",
    invoked: "OneDrive item ready"
  },
  onedrive_open_files: {
    description: "Use this when the user provides one or more exact filenames and wants their contents. It locates and extracts up to five files in one read-only call; use search then fetch for discovery, partial names, or ambiguous results.",
    invoking: "Opening OneDrive files…",
    invoked: "OneDrive files ready"
  },
  onedrive_preview_actions: {
    description: "Use this when the user wants to preview one or more rename, move, copy, sharing-link, or permission-revocation actions. This read-only batch makes no changes, returns scoped preview tokens for later live calls, and returns sharing counts without identities.",
    invoking: "Previewing OneDrive actions…",
    invoked: "OneDrive previews ready"
  },
  onedrive_list: {
    description: "Use this when the user wants the direct children of a known OneDrive folder or path. Use search instead when the folder location is not already known.",
    invoking: "Listing OneDrive folder…",
    invoked: "Folder listing ready"
  },
  onedrive_office_capabilities: {
    description: "Use this when preparing to edit a Word, Excel, or PowerPoint file and the supported structured operations must be checked before the edit.",
    invoking: "Checking Office capabilities…",
    invoked: "Office capabilities ready"
  },
  onedrive_office_batch_transform: {
    description: "Use this when the user wants structured edits across one or more Word, Excel, or PowerPoint files. Read each file with fetch first, check Office capabilities, and preview before confirmation.",
    invoking: "Preparing Office changes…",
    invoked: "Office change result ready"
  },
  onedrive_upload_file: {
    description: "Use this when the user wants to upload a ChatGPT-provided file to OneDrive. Preview first; replacement requires the existing item's expected identity and explicit confirmation.",
    invoking: "Preparing OneDrive upload…",
    invoked: "OneDrive upload result ready"
  },
  onedrive_write_text: {
    description: "Use this when the user wants to create or fully replace a UTF-8 text or code file in OneDrive. Preview before replacing an existing file.",
    invoking: "Preparing text file…",
    invoked: "Text file result ready"
  },
  onedrive_patch_text: {
    description: "Use this when the user wants a targeted line or text change in an existing OneDrive text file while preserving the rest of the file. Fetch the file before patching and preview before confirmation.",
    invoking: "Preparing text patch…",
    invoked: "Text patch result ready"
  },
  onedrive_create_folder: {
    description: "Use this when the user wants to create a new folder in a known OneDrive location. Preview any conflict behavior before confirmation.",
    invoking: "Preparing OneDrive folder…",
    invoked: "Folder creation result ready"
  },
  onedrive_rename: {
    description: "Use this when the user has approved a live rename already previewed by onedrive_preview_actions. Inputs: itemId or path, newName, dryRun false, confirmed true, expectedId or expectedName, and the returned previewToken.",
    invoking: "Preparing OneDrive rename…",
    invoked: "Rename result ready"
  },
  onedrive_move: {
    description: "Use this when the user has approved a live move already previewed by onedrive_preview_actions. Inputs: source itemId or path, destination parent, dryRun false, confirmed true, expectedId or expectedName, and the returned previewToken.",
    invoking: "Preparing OneDrive move…",
    invoked: "Move result ready"
  },
  onedrive_copy: {
    description: "Use this when the user has approved a live copy already previewed by onedrive_preview_actions. Inputs: source itemId or path, destination parent, dryRun false, confirmed true, expectedId or expectedName, previewToken, and optional newName or waitForCompletion.",
    invoking: "Preparing OneDrive copy…",
    invoked: "Copy result ready"
  },
  onedrive_create_sharing_link: {
    description: "Use this when the user has approved a live sharing link already previewed by onedrive_preview_actions. Inputs: itemId or path, type, scope, dryRun false, confirmed true, expectedId or expectedName, and previewToken; use invite permission for named recipients.",
    invoking: "Preparing sharing link…",
    invoked: "Sharing link result ready"
  },
  onedrive_invite_permission: {
    description: "Use this when the user wants to grant OneDrive access to specific named recipients. Use create sharing link for a general link, and preview recipients and roles before sending.",
    invoking: "Preparing OneDrive invitation…",
    invoked: "Invitation result ready"
  },
  onedrive_revoke_permission: {
    description: "Use this when the user has approved a live permission removal already previewed by onedrive_preview_actions. Inputs: itemId or path, permissionId, dryRun false, confirmed true, expectedId or expectedName, and previewToken.",
    invoking: "Preparing permission removal…",
    invoked: "Permission removal result ready"
  },
  onedrive_permissions: {
    description: "Use this when the user explicitly wants the identities that can access a OneDrive file or folder. For permission counts or a sharing-link preview, use onedrive_preview_actions so identity details are not returned.",
    invoking: "Checking OneDrive permissions…",
    invoked: "Permissions ready"
  },
  onedrive_delete: {
    description: "Use this when the user wants to move one OneDrive file or folder to the recycle bin. Locate the exact item and preview before confirmation; do not use for permanent deletion.",
    invoking: "Preparing recycle-bin move…",
    invoked: "Recycle-bin result ready"
  },
  onedrive_restore_deleted: {
    description: "Use this when the user wants to restore a known OneDrive recycle-bin item by item ID. Preview the exact restore target and optional new name before confirmation.",
    invoking: "Preparing OneDrive restore…",
    invoked: "Restore result ready"
  },
  onedrive_permanent_delete: {
    description: "Use this when the user explicitly wants to irreversibly delete a known OneDrive item without the recycle bin. Require exact identity, irreversible acknowledgement, preview, and confirmation.",
    invoking: "Preparing permanent deletion…",
    invoked: "Permanent deletion result ready"
  }
});

function compactChatgptToolDescriptor(tool) {
  const compact = JSON.parse(JSON.stringify(tool));
  const metadata = chatgptToolMetadata[compact.name];
  if (metadata) {
    compact.description = metadata.description;
    compact._meta = {
      ...(compact._meta || {}),
      "openai/toolInvocation/invoking": metadata.invoking,
      "openai/toolInvocation/invoked": metadata.invoked
    };
  }
  if (compact.name === "onedrive_office_batch_transform") {
    const item = compact.inputSchema?.properties?.items?.items;
    if (item?.properties) item.properties.operations = compactOfficeOperationSchema;
  } else if ([
    "onedrive_word_batch_update",
    "onedrive_excel_batch_update",
    "onedrive_powerpoint_batch_update"
  ].includes(compact.name) && compact.inputSchema?.properties) {
    compact.inputSchema.properties.operations = compactOfficeOperationSchema;
  }
  return compact;
}

function selectedToolProfile(env = process.env) {
  const profile = String(env.ONEDRIVE_TOOL_PROFILE || "full").trim().toLowerCase();
  if (!["full", "chatgpt"].includes(profile)) {
    throw new Error("ONEDRIVE_TOOL_PROFILE must be full or chatgpt.");
  }
  return profile;
}

const toolProfile = selectedToolProfile();
const advertisedTools = toolProfile === "chatgpt"
  ? executableTools.filter((tool) => chatgptToolNames.has(tool.name)).map(compactChatgptToolDescriptor)
  : tools;
const advertisedContractHash = createHash("sha256").update(JSON.stringify(advertisedTools)).digest("hex").slice(0, 12);
const manifestServerVersion = pluginManifest.version || "0.1.0";
const advertisedServerVersion = toolProfile === "chatgpt"
  ? `${manifestServerVersion}${manifestServerVersion.includes("+") ? "." : "+"}chatgpt.${advertisedContractHash}`
  : manifestServerVersion;
const serverInstructions = toolProfile === "chatgpt"
  ? "Use onedrive_open_files once for exact filenames/content. For subtle descriptions, topics, aliases, or unknown titles, pass natural-language intent once to search, then fetch; search infers aliases. Pass ids unchanged. Use onedrive_preview_actions for read-only rename/move/copy/sharing/revoke previews with identity-free access counts. Live mutations require approval, preview token, and expected id. Use onedrive_list only for known folders. For Office edits, fetch, check capabilities, then transform."
  : "Use onedrive_find for normal OneDrive lookup and the matching structured read tool before an Office edit. Use onedrive_list only for direct folder listings. Keep results bounded. Locate an item before changing it. Mutations default to preview and require confirmation.";

const toolByName = new Map(executableTools.map((tool) => [tool.name, tool]));

function schemaTypeMatches(value, type) {
  if (Array.isArray(type)) return type.some((candidate) => candidate === "null" ? value === null : schemaTypeMatches(value, candidate));
  if (type === "array") return Array.isArray(value);
  if (type === "integer") return Number.isInteger(value);
  if (type === "number") return typeof value === "number" && Number.isFinite(value);
  if (type === "object") return value !== null && typeof value === "object" && !Array.isArray(value);
  return typeof value === type;
}

function validationDetail(path, message) {
  return { path, message };
}

function cloneDefault(value) {
  if (value === null || typeof value !== "object") return value;
  return JSON.parse(JSON.stringify(value));
}

function jsonSchemaValuesEqual(left, right) {
  if (left === right) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left)
      && Array.isArray(right)
      && left.length === right.length
      && left.every((value, index) => jsonSchemaValuesEqual(value, right[index]));
  }
  if (left && right && typeof left === "object" && typeof right === "object") {
    const leftKeys = Object.keys(left).sort();
    const rightKeys = Object.keys(right).sort();
    return leftKeys.length === rightKeys.length
      && leftKeys.every((key, index) => key === rightKeys[index] && jsonSchemaValuesEqual(left[key], right[key]));
  }
  return false;
}

function validateSchemaValue(value, schema = {}, path = "$") {
  const details = [];
  let normalized = value;

  if (schema.default !== undefined && normalized === undefined) {
    normalized = cloneDefault(schema.default);
  }

  if (normalized === undefined) return { ok: true, value: normalized, details };

  if (schema.anyOf?.length && !schema.type) {
    const branches = schema.anyOf.map((branch) => validateSchemaValue(normalized, branch, path));
    const matched = branches.filter((branch) => branch.ok);
    if (!matched.length) {
      details.push(validationDetail(path, "Value does not match any allowed schema."));
      const mostSpecific = branches.sort((left, right) => left.details.length - right.details.length)[0];
      if (mostSpecific) details.push(...mostSpecific.details);
      return { ok: false, value: normalized, details };
    }
    return { ok: true, value: matched[0].value, details };
  }

  if (schema.type && !schemaTypeMatches(normalized, schema.type)) {
    details.push(validationDetail(path, `Expected ${Array.isArray(schema.type) ? schema.type.join(" or ") : schema.type}.`));
    return { ok: false, value: normalized, details };
  }

  if (schema.const !== undefined && normalized !== schema.const) {
    details.push(validationDetail(path, `Expected constant value ${JSON.stringify(schema.const)}.`));
  }

  if (schema.enum && !schema.enum.includes(normalized)) {
    details.push(validationDetail(path, `Expected one of: ${schema.enum.join(", ")}.`));
  }

  if ((schema.type === "number" || schema.type === "integer") && typeof normalized === "number") {
    if (schema.minimum !== undefined && normalized < schema.minimum) {
      details.push(validationDetail(path, `Must be >= ${schema.minimum}.`));
    }
    if (schema.maximum !== undefined && normalized > schema.maximum) {
      details.push(validationDetail(path, `Must be <= ${schema.maximum}.`));
    }
    if (schema.exclusiveMinimum !== undefined && normalized <= schema.exclusiveMinimum) {
      details.push(validationDetail(path, `Must be > ${schema.exclusiveMinimum}.`));
    }
    if (schema.exclusiveMaximum !== undefined && normalized >= schema.exclusiveMaximum) {
      details.push(validationDetail(path, `Must be < ${schema.exclusiveMaximum}.`));
    }
  }

  if (schema.type === "string" && typeof normalized === "string") {
    if (schema.minLength !== undefined && normalized.length < schema.minLength) {
      details.push(validationDetail(path, `Must be at least ${schema.minLength} characters.`));
    }
    if (schema.maxLength !== undefined && normalized.length > schema.maxLength) {
      details.push(validationDetail(path, `Must be at most ${schema.maxLength} characters.`));
    }
    if (schema.pattern !== undefined && !(new RegExp(schema.pattern).test(normalized))) {
      details.push(validationDetail(path, `Must match pattern ${schema.pattern}.`));
    }
  }

  if (schema.type === "array" && Array.isArray(normalized)) {
    if (schema.minItems !== undefined && normalized.length < schema.minItems) {
      details.push(validationDetail(path, `Must contain at least ${schema.minItems} item(s).`));
    }
    if (schema.maxItems !== undefined && normalized.length > schema.maxItems) {
      details.push(validationDetail(path, `Must contain at most ${schema.maxItems} item(s).`));
    }
    if (schema.items) {
      normalized = normalized.map((item, index) => {
        const child = validateSchemaValue(item, schema.items, `${path}[${index}]`);
        details.push(...child.details);
        return child.value;
      });
    }
    if (schema.uniqueItems === true) {
      const duplicateIndex = normalized.findIndex((item, index) =>
        normalized.slice(0, index).some((candidate) => jsonSchemaValuesEqual(candidate, item))
      );
      if (duplicateIndex !== -1) {
        details.push(validationDetail(`${path}[${duplicateIndex}]`, "Must not duplicate another array item."));
      }
    }
  }

  if (schema.type === "object" && normalized && typeof normalized === "object" && !Array.isArray(normalized)) {
    const properties = schema.properties || {};
    normalized = { ...normalized };
    for (const required of schema.required || []) {
      if (!Object.hasOwn(normalized, required)) {
        details.push(validationDetail(`${path}.${required}`, "Required field is missing."));
      }
    }
    if (schema.additionalProperties === false) {
      for (const key of Object.keys(normalized)) {
        if (!Object.hasOwn(properties, key)) {
          details.push(validationDetail(`${path}.${key}`, "Unknown property."));
        }
      }
    }
    for (const [key, propertySchema] of Object.entries(properties)) {
      const child = validateSchemaValue(normalized[key], propertySchema, `${path}.${key}`);
      details.push(...child.details);
      if (child.value !== undefined || Object.hasOwn(normalized, key)) normalized[key] = child.value;
    }
    if (schema.anyOf?.length) {
      const branchLabels = schema.anyOf
        .map((branch) => (branch.required || []).join(" + "))
        .filter(Boolean);
      const branchMatches = schema.anyOf.filter((branch) =>
        (branch.required || []).every((key) => Object.hasOwn(normalized, key))
      );
      if (branchMatches.length === 0) {
        details.push(validationDetail(path, `Must include one target option: ${branchLabels.join(" or ")}.`));
      } else if (branchMatches.length > 1) {
        details.push(validationDetail(path, `Must include exactly one target option: ${branchLabels.join(" or ")}.`));
      }
    }
  }

  return { ok: details.length === 0, value: normalized, details };
}

function validateToolArguments(name, args = {}) {
  const tool = toolByName.get(name);
  if (!tool) return { ok: true, args };
  const result = validateSchemaValue(args || {}, tool.inputSchema || { type: "object" }, "$");
  if (result.ok) {
    const pairedFields = [
      ["relativePath", "preset"],
      ["destinationParentRelativePath", "destinationParentPreset"],
      ["parentRelativePath", "parentPreset"],
      ["remoteRelativePath", "remotePreset"]
    ];
    const pairDetails = pairedFields
      .filter(([relativeField, presetField]) => result.value?.[relativeField] !== undefined && !result.value?.[presetField])
      .map(([relativeField, presetField]) => validationDetail(`$.${relativeField}`, `${relativeField} requires ${presetField}.`));
    const officeKindDetails = [];
    if (name === "onedrive_office_batch_transform") {
      const schemas = { word: wordOperationsSchema, excel: excelOperationsSchema, powerpoint: powerpointOperationsSchema };
      for (const [index, item] of (result.value?.items || []).entries()) {
        const operationResult = validateSchemaValue(item.operations, schemas[item.kind] || {}, `$.items[${index}].operations`);
        if (!operationResult.ok) {
          officeKindDetails.push(validationDetail(`$.items[${index}].operations`, `Operations do not match kind ${item.kind}.`), ...operationResult.details);
        } else {
          item.operations = operationResult.value;
        }
      }
    }
    const semanticDetails = [...pairDetails, ...officeKindDetails];
    if (!semanticDetails.length) return { ok: true, args: result.value };
    return { ok: false, error: { error: "invalid_arguments", tool: name, details: semanticDetails } };
  }
  return {
    ok: false,
    error: {
      error: "invalid_arguments",
      tool: name,
      details: result.details
    }
  };
}

function readAdditionalLocalSyncRoots() {
  const raw = process.env.ONEDRIVE_ADDITIONAL_LOCAL_SYNC_ROOTS;
  if (!raw) return [];
  let paths;
  try {
    paths = JSON.parse(raw);
  } catch (error) {
    throw new Error(`ONEDRIVE_ADDITIONAL_LOCAL_SYNC_ROOTS must be a JSON array of absolute paths: ${error.message}`);
  }
  if (!Array.isArray(paths) || paths.some((path) => typeof path !== "string" || !path.trim() || !isAbsolute(path))) {
    throw new Error("ONEDRIVE_ADDITIONAL_LOCAL_SYNC_ROOTS must be a JSON array containing only non-empty absolute paths.");
  }
  return [...new Set(paths.map((path) => resolve(path)))].map((path) => ({ path, prefix: false }));
}

function readLocalConfig() {
  try {
    const parsed = JSON.parse(readFileSync(configPath, "utf8"));
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      localConfigReadError = `Could not read ${configPath}: top-level JSON value must be an object.`;
      return {};
    }
    return parsed;
  } catch (error) {
    if (error?.code !== "ENOENT") {
      localConfigReadError = `Could not read ${configPath}: ${error.message}`;
    }
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

function normalizeAuthContextToken(token = {}) {
  if (token.auth_context_id) return token;
  return { ...token, auth_context_id: randomUUID() };
}

function invalidateActiveStorageScope() {
  storageScopeGeneration += 1;
  activeStorageScopePromise = null;
  activeStorageScopeKey = null;
  activeStorageScopeGeneration = null;
  metadataCacheMemory = null;
  contentIndexMemory = null;
  metadataCacheMemoryGeneration = null;
  contentIndexMemoryGeneration = null;
  metadataCacheLoadPromise = null;
  contentIndexLoadPromise = null;
  metadataCacheLoadGeneration = null;
  contentIndexLoadGeneration = null;
  metadataCacheFileVersion = null;
  contentIndexFileVersion = null;
  previewTokens.clear();
  chatgptFetchSnapshots.clear();
  chatgptRevalidations.clear();
  chatgptRevalidationLastStartedAt.clear();
}

function accountContextChangedError(operation = "OneDrive operation") {
  const error = new Error(`${operation} was cancelled because the OneDrive authentication or storage scope changed while it was in flight.`);
  error.code = "ONEDRIVE_ACCOUNT_CONTEXT_CHANGED";
  return error;
}

function isAccountContextChangedError(error) {
  return error?.code === "ONEDRIVE_ACCOUNT_CONTEXT_CHANGED";
}

function assertToolAccountGeneration(operation = "OneDrive operation") {
  const store = toolCallContext.getStore();
  if (!store) return;
  if (store.authGeneration !== authGeneration || store.storageScopeGeneration !== storageScopeGeneration) {
    throw accountContextChangedError(operation);
  }
}

function adoptCurrentToolAccountGeneration() {
  const store = toolCallContext.getStore();
  if (!store) return;
  store.authGeneration = authGeneration;
  store.storageScopeGeneration = storageScopeGeneration;
}

function currentAuthContextId() {
  const requestContextId = toolCallContext.getStore()?.authContextId;
  if (requestContextId) return requestContextId;
  if (process.env.ONEDRIVE_TEST_ACCESS_TOKEN) return testAuthContextId;
  const cfg = config();
  const current = tokenCache || getKeychainToken(cfg);
  if (!current) return null;
  if (current.auth_context_id) return current.auth_context_id;
  const migrated = normalizeAuthContextToken(current);
  tokenCache = migrated;
  setKeychainToken(migrated, cfg);
  return migrated.auth_context_id;
}

function storageScopesEqual(left, right) {
  return Boolean(left?.authContextId && left?.driveId
    && left.authContextId === right?.authContextId
    && left.driveId === right?.driveId);
}

function storageScopeKey(scope) {
  return scope?.authContextId && scope?.driveId ? `${scope.authContextId}:${scope.driveId}` : null;
}

function scopedStatePath(root, scope) {
  const key = storageScopeKey(scope);
  if (!key) throw new Error("A complete authentication-context and drive scope is required for local state.");
  const opaqueScope = createHash("sha256").update(key).digest("hex");
  return join(root, `${opaqueScope}.json`);
}

function assertStorageScopeGuard(guard, operation = "OneDrive local-state operation") {
  assertToolAccountGeneration(operation);
  if (!guard?.scope || guard.generation !== storageScopeGeneration) throw accountContextChangedError(operation);
  const store = toolCallContext.getStore();
  const authContextId = currentAuthContextId();
  if (!authContextId || authContextId !== guard.scope.authContextId) throw accountContextChangedError(operation);
  if (store?.authMode === "oauth") {
    if (!storageScopesEqual(store.storageScope, guard.scope)) throw accountContextChangedError(operation);
    return guard;
  }
  if (activeStorageScopeGeneration !== guard.generation || activeStorageScopeKey !== storageScopeKey(guard.scope)) {
    throw accountContextChangedError(operation);
  }
  return guard;
}

async function captureStorageScopeGuard(operation = "OneDrive local-state operation") {
  assertToolAccountGeneration(operation);
  const generation = storageScopeGeneration;
  const scope = await activeStorageScope();
  return assertStorageScopeGuard({ generation, scope }, operation);
}

async function activeStorageScope() {
  assertToolAccountGeneration("OneDrive storage-scope resolution");
  const store = toolCallContext.getStore();
  if (store?.authMode === "oauth") {
    await getAccessToken();
    if (store.storageScope?.authContextId === store.authContextId && store.storageScope?.driveId) {
      return store.storageScope;
    }
    const drive = await graph("/me/drive?$select=id");
    if (!drive?.id) throw new Error("Microsoft Graph did not return the current OneDrive drive ID. Refusing to use account-scoped local state.");
    store.storageScope = { authContextId: store.authContextId, driveId: drive.id };
    return store.storageScope;
  }
  const generation = storageScopeGeneration;
  await getAccessToken();
  if (generation !== storageScopeGeneration) throw accountContextChangedError("OneDrive storage-scope resolution");
  assertToolAccountGeneration("OneDrive storage-scope resolution");
  const authContextId = currentAuthContextId();
  if (!authContextId) throw new Error("OneDrive authentication context is unavailable. Refusing to use account-scoped local state.");
  if (activeStorageScopePromise
    && activeStorageScopeGeneration === generation
    && activeStorageScopeKey?.startsWith(`${authContextId}:`)) {
    const scope = await activeStorageScopePromise;
    return assertStorageScopeGuard({ generation, scope }, "OneDrive storage-scope resolution").scope;
  }
  const promise = (async () => {
    const drive = await graph("/me/drive?$select=id");
    if (generation !== storageScopeGeneration || currentAuthContextId() !== authContextId) {
      throw accountContextChangedError("OneDrive storage-scope resolution");
    }
    if (!drive?.id) throw new Error("Microsoft Graph did not return the current OneDrive drive ID. Refusing to use account-scoped local state.");
    return { authContextId, driveId: drive.id };
  })();
  activeStorageScopePromise = promise;
  activeStorageScopeKey = `${authContextId}:pending`;
  activeStorageScopeGeneration = generation;
  try {
    const scope = await promise;
    if (activeStorageScopePromise !== promise
      || activeStorageScopeGeneration !== generation
      || generation !== storageScopeGeneration
      || currentAuthContextId() !== authContextId) {
      throw accountContextChangedError("OneDrive storage-scope resolution");
    }
    activeStorageScopeKey = storageScopeKey(scope);
    return assertStorageScopeGuard({ generation, scope }, "OneDrive storage-scope resolution").scope;
  } catch (error) {
    if (activeStorageScopePromise === promise) {
      activeStorageScopePromise = null;
      activeStorageScopeKey = null;
      activeStorageScopeGeneration = null;
    }
    throw error;
  }
}

function boolSetting(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  return ["1", "true", "yes", "on"].includes(String(value).toLowerCase());
}

function intSetting(value, defaultValue, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return defaultValue;
  return Math.min(Math.max(Math.trunc(number), min), max);
}

function listSetting(value, defaultValue) {
  if (Array.isArray(value)) return value;
  if (typeof value === "string") return value.split(/[,\s]+/).filter(Boolean);
  return defaultValue;
}

function pluginSettings() {
  const cfg = localConfig.settings || {};
  const indexing = localConfig.indexing || {};
  const supportedIndexedFileTypes = normalizeExtensions(listSetting(
    process.env.ONEDRIVE_INDEX_EXTENSIONS || indexing.supportedExtensions || cfg.supportedIndexedFileTypes,
    [".txt", ".md", ".csv", ".json", ".jsonl", ".xml", ".yaml", ".yml", ".html", ".css", ".js", ".mjs", ".ts", ".tsx", ".py", ".sql", ".log"]
  ));
  return {
    storageRoot,
    cacheRoot,
    cachePath,
    contentIndexPath,
    maxScanDepth: intSetting(process.env.ONEDRIVE_MAX_SCAN_DEPTH ?? cfg.maxScanDepth, 25, 0, 50),
    maxIndexedFileSize: intSetting(process.env.ONEDRIVE_MAX_INDEXED_FILE_SIZE ?? indexing.maxFileSize ?? cfg.maxIndexedFileSize, defaultMaxIndexedFileSize, 1024, textFileLimit),
    supportedIndexedFileTypes: [...supportedIndexedFileTypes],
    cacheTtlSeconds: intSetting(process.env.ONEDRIVE_CACHE_TTL_SECONDS ?? cfg.cacheTtlSeconds, 900, 0, 30 * 24 * 60 * 60),
    concurrencyLimit: intSetting(process.env.ONEDRIVE_CONCURRENCY_LIMIT ?? cfg.concurrencyLimit, 2, 1, 8),
    deltaSyncEnabled: boolSetting(process.env.ONEDRIVE_DELTA_SYNC_ENABLED ?? cfg.deltaSyncEnabled, true),
    contentIndexEnabled: boolSetting(process.env.ONEDRIVE_CONTENT_INDEX_ENABLED ?? indexing.enabled ?? cfg.contentIndexEnabled, true),
    includeOfficeTextExport: boolSetting(process.env.ONEDRIVE_INDEX_OFFICE_EXPORT ?? indexing.includeOfficeTextExport, false),
    fullReindexCommand: "onedrive_content_index_refresh({ force: true, refreshMetadata: true })"
  };
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
  return `${identityBaseUrl()}/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
}

function deviceCodeUrl(tenant) {
  return `${identityBaseUrl()}/${encodeURIComponent(tenant)}/oauth2/v2.0/devicecode`;
}

function identityBaseUrl() {
  return (process.env.ONEDRIVE_IDENTITY_BASE_URL || "https://login.microsoftonline.com").replace(/\/+$/, "");
}

function isMsaOnlyEndpointError(error) {
  const message = `${error?.body?.error || ""} ${error?.body?.error_description || ""} ${error?.message || ""}`;
  return message.includes("AADSTS9002331") || message.includes("/consumers endpoint");
}

function isMissingStoredAuthenticationError(error) {
  const message = String(error?.message || "");
  return message.includes("OneDrive is not authenticated")
    || message.includes("Stored token has no refresh token");
}

function isReauthenticationRequiredError(error) {
  const code = String(error?.body?.error || "").toLowerCase();
  return new Set([
    "invalid_grant",
    "interaction_required",
    "login_required",
    "consent_required"
  ]).has(code) || String(error?.message || "").includes("does not match the requested client, tenant, or scopes");
}

async function postFormWithConsumerFallback(kind, cfg, params) {
  try {
    return { body: await postForm(kind === "device" ? deviceCodeUrl(cfg.tenant) : tokenUrl(cfg.tenant), params), tenant: cfg.tenant };
  } catch (error) {
    if (cfg.tenant === "common" && isMsaOnlyEndpointError(error)) {
      return { body: await postForm(kind === "device" ? deviceCodeUrl("consumers") : tokenUrl("consumers"), params), tenant: "consumers" };
    }
    throw error;
  }
}

function keychainAccount() {
  return "tokens";
}

function authVault(cfg = config()) {
  return createAuthVault({
    account: keychainAccount(),
    service: cfg.keychainService,
    storageRoot
  });
}

function getKeychainToken(cfg = config()) {
  return authVault(cfg).read();
}

function setKeychainToken(token, cfg = config()) {
  authVault(cfg).write(token);
}

function deleteKeychainToken(cfg = config()) {
  return authVault(cfg).remove();
}

function publicConfig() {
  const cfg = config();
  const mcpAuth = oauthSettings();
  const vault = authVault(cfg);
  let stored = null;
  let authStoreError = null;
  try {
    stored = vault.read();
  } catch (error) {
    authStoreError = safeToolErrorMessage(error);
  }
  return {
    clientIdConfigured: Boolean(cfg.clientId),
    tenant: cfg.tenant,
    scopes: cfg.scopes,
    authenticationStore: vault.mode,
    storedCredentialConfigured: Boolean(stored?.refresh_token),
    ...(authStoreError ? { authenticationStoreError: authStoreError } : {}),
    keychainService: cfg.keychainService,
    keychainTokenConfigured: vault.mode === "keychain" && Boolean(stored?.refresh_token),
    mcpAuthentication: {
      mode: mcpAuth.mode,
      delegatedRequest: toolCallContext.getStore()?.authMode === "oauth",
      ...(mcpAuth.mode === "oauth" ? {
        resource: mcpAuth.resource,
        scope: mcpAuth.apiScope,
        authority: mcpAuth.authority
      } : {})
    },
    configPath,
    ...(localConfigReadError ? { configReadError: localConfigReadError } : {}),
    pathPresets: pathPresets(),
    settings: pluginSettings()
  };
}

function emptyMetadataCache(scope = null) {
  return {
    version: 4,
    scope,
    createdAt: new Date().toISOString(),
    updatedAt: null,
    deltaLink: null,
    deltaNextLink: null,
    deltaTarget: null,
    scanRoot: null,
    pathRootsById: {},
    itemCount: 0,
    itemsById: {},
    pathsByLower: {},
    searchTombstones: []
  };
}

async function ensurePrivateDirectory(path) {
  await mkdir(path, { recursive: true, mode: 0o700 });
  await chmod(path, 0o700);
}

async function hardenPrivateFile(path) {
  try {
    await chmod(path, 0o600);
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
}

async function writePrivateFile(path, data, options = {}) {
  await writeFile(path, data, { ...options, mode: 0o600 });
  await hardenPrivateFile(path);
}

async function writePrivateFileAtomic(path, data, options = {}) {
  const { beforeCommit, ...writeOptions } = options;
  await ensurePrivateDirectory(dirname(path));
  const temporaryPath = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
  let handle;
  try {
    handle = await open(temporaryPath, "wx", 0o600);
    await handle.writeFile(data, writeOptions);
    await handle.sync();
    await handle.close();
    handle = null;
    if (beforeCommit) await beforeCommit();
    await renameFile(temporaryPath, path);
    await hardenPrivateFile(path);
  } finally {
    await handle?.close().catch(() => null);
    await rm(temporaryPath, { force: true }).catch(() => null);
  }
}

function collisionPath(path, suffix) {
  const parsed = parse(path);
  return join(parsed.dir, `${parsed.name} (${suffix})${parsed.ext}`);
}

async function reserveLocalDestination(preferredPath, options = {}) {
  const preferred = resolve(preferredPath);
  await ensurePrivateDirectory(dirname(preferred));
  if (options.overwrite === true) return { path: preferred, reserved: false };
  const allowAlternate = options.allowAlternate === true;
  for (let suffix = 1; suffix <= 10_000; suffix += 1) {
    const candidate = suffix === 1 ? preferred : collisionPath(preferred, suffix);
    try {
      const handle = await open(candidate, "wx", 0o600);
      await handle.close();
      return { path: candidate, reserved: true };
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
      if (!allowAlternate) {
        throw new Error(`Local file already exists: ${preferred}. Pass overwrite: true to replace it.`);
      }
    }
  }
  throw new Error(`Could not reserve a unique local destination near ${preferred}.`);
}

async function localFileVersion(path) {
  try {
    const fileStat = await stat(path);
    return `${fileStat.dev}:${fileStat.ino}:${fileStat.size}:${fileStat.mtimeMs}`;
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  }
}

async function withFileLock(lockPath, fn, options = {}) {
  const timeoutMs = options.timeoutMs ?? 10_000;
  const staleMs = options.staleMs ?? 60_000;
  const startedAt = Date.now();
  await ensurePrivateDirectory(dirname(lockPath));
  let handle;
  while (!handle) {
    try {
      handle = await open(lockPath, "wx", 0o600);
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
      try {
        const lockStat = await stat(lockPath);
        if (Date.now() - lockStat.mtimeMs > staleMs) {
          await rm(lockPath, { force: true });
          continue;
        }
      } catch (statError) {
        if (statError.code !== "ENOENT") throw statError;
        continue;
      }
      if (Date.now() - startedAt >= timeoutMs) throw new Error(`Timed out waiting for local storage lock: ${lockPath}`);
      await sleep(Math.min(100, 10 + Math.floor((Date.now() - startedAt) / 10)));
    }
  }
  try {
    return await fn();
  } finally {
    await handle.close().catch(() => null);
    await rm(lockPath, { force: true }).catch(() => null);
  }
}

async function runSerialized(queueName, fn) {
  const previous = queueName === "metadata" ? metadataMutationQueue : contentIndexMutationQueue;
  let release;
  const current = new Promise((resolve) => { release = resolve; });
  if (queueName === "metadata") metadataMutationQueue = previous.catch(() => null).then(() => current);
  else contentIndexMutationQueue = previous.catch(() => null).then(() => current);
  await previous.catch(() => null);
  try {
    return await fn();
  } finally {
    release();
  }
}

function normalizeMetadataCache(parsed = {}, scope = null) {
  const cache = {
    ...emptyMetadataCache(scope),
    ...parsed,
    version: 4,
    scope,
    itemsById: parsed.itemsById || {},
    pathsByLower: parsed.pathsByLower || {},
    pathRootsById: parsed.pathRootsById || {},
    searchTombstones: Array.isArray(parsed.searchTombstones) ? parsed.searchTombstones : []
  };
  if ((cache.deltaLink || cache.deltaNextLink) && !parsed.deltaTarget) {
    cache.deltaLink = null;
    cache.deltaNextLink = null;
    cache.deltaTarget = null;
  }
  return cache;
}

async function readMetadataCacheFromDisk(existingGuard = null) {
  const guard = existingGuard || await captureStorageScopeGuard("metadata cache read");
  assertStorageScopeGuard(guard, "metadata cache read");
  await ensurePrivateDirectory(cacheRoot);
  try {
    const parsed = JSON.parse(await readFile(cachePath, "utf8"));
    await hardenPrivateFile(cachePath);
    assertStorageScopeGuard(guard, "metadata cache read");
    if (parsed.version !== 4 || !storageScopesEqual(parsed.scope, guard.scope)) return emptyMetadataCache(guard.scope);
    return normalizeMetadataCache(parsed, guard.scope);
  } catch (error) {
    if (error?.code === "ENOENT") {
      assertStorageScopeGuard(guard, "metadata cache read");
      return emptyMetadataCache(guard.scope);
    }
    throw error;
  }
}

async function loadMetadataCache() {
  const guard = await captureStorageScopeGuard("metadata cache load");
  if (metadataCacheLoadPromise && metadataCacheLoadGeneration !== guard.generation) {
    throw accountContextChangedError("metadata cache load");
  }
  if (!metadataCacheLoadPromise) {
    let loadPromise;
    loadPromise = (async () => {
      try {
        assertStorageScopeGuard(guard, "metadata cache load");
        const diskVersion = await localFileVersion(cachePath);
        assertStorageScopeGuard(guard, "metadata cache load");
        if (metadataCacheMemory
          && metadataCacheMemoryGeneration === guard.generation
          && storageScopesEqual(metadataCacheMemory.scope, guard.scope)
          && diskVersion === metadataCacheFileVersion) return metadataCacheMemory;
        let loaded = null;
        await withFileLock(metadataCacheLockPath, async () => {
          assertStorageScopeGuard(guard, "metadata cache load");
          const lockedVersion = await localFileVersion(cachePath);
          assertStorageScopeGuard(guard, "metadata cache load");
          if (metadataCacheMemory
            && metadataCacheMemoryGeneration === guard.generation
            && storageScopesEqual(metadataCacheMemory.scope, guard.scope)
            && lockedVersion === metadataCacheFileVersion) {
            loaded = metadataCacheMemory;
            return;
          }
          loaded = await readMetadataCacheFromDisk(guard);
          assertStorageScopeGuard(guard, "metadata cache publication");
          metadataCacheMemory = loaded;
          metadataCacheMemoryGeneration = guard.generation;
          metadataCacheFileVersion = await localFileVersion(cachePath);
        });
        assertStorageScopeGuard(guard, "metadata cache return");
        return loaded || metadataCacheMemory;
      } catch (error) {
        if (isAccountContextChangedError(error)) throw error;
        recordLocalWarning("metadata cache read", error);
        assertStorageScopeGuard(guard, "metadata cache fallback publication");
        metadataCacheMemory = emptyMetadataCache(guard.scope);
        metadataCacheMemoryGeneration = guard.generation;
        metadataCacheFileVersion = null;
      } finally {
        if (metadataCacheLoadPromise === loadPromise) {
          metadataCacheLoadPromise = null;
          metadataCacheLoadGeneration = null;
        }
      }
      assertStorageScopeGuard(guard, "metadata cache return");
      return metadataCacheMemory;
    })();
    metadataCacheLoadPromise = loadPromise;
    metadataCacheLoadGeneration = guard.generation;
  }
  const result = await metadataCacheLoadPromise;
  assertStorageScopeGuard(guard, "metadata cache return");
  if (!storageScopesEqual(result?.scope, guard.scope)) throw accountContextChangedError("metadata cache return");
  return result;
}

async function saveMetadataCache(cache, existingGuard = null) {
  const guard = existingGuard || await captureStorageScopeGuard("metadata cache write");
  assertStorageScopeGuard(guard, "metadata cache write");
  if (cache.scope && !storageScopesEqual(cache.scope, guard.scope)) {
    throw new Error("Metadata cache scope changed during this operation. Refusing to persist cross-account state.");
  }
  cache.version = 4;
  cache.scope = guard.scope;
  cache.updatedAt = new Date().toISOString();
  cache.itemCount = Object.keys(cache.itemsById || {}).length;
  await ensurePrivateDirectory(cacheRoot);
  assertStorageScopeGuard(guard, "metadata cache write");
  await writePrivateFileAtomic(cachePath, JSON.stringify(cache), {
    beforeCommit: async () => assertStorageScopeGuard(guard, "metadata cache commit")
  });
  assertStorageScopeGuard(guard, "metadata cache publication");
  metadataCacheMemory = cache;
  metadataCacheMemoryGeneration = guard.generation;
  metadataCacheFileVersion = await localFileVersion(cachePath);
  assertStorageScopeGuard(guard, "metadata cache return");
  const store = toolCallContext.getStore();
  if (store) store.metadataCacheWrites = (store.metadataCacheWrites || 0) + 1;
  return cache;
}

async function withMetadataCacheBatch(fn) {
  return await fn();
}

function cachePathKey(remotePath = "") {
  return cleanPath(remotePath).toLowerCase();
}

function mergeDefinedMetadata(previous = {}, incoming = {}) {
  const merged = { ...(previous || {}) };
  for (const [key, value] of Object.entries(incoming || {})) {
    if (value === undefined) continue;
    if ((key === "file" || key === "folder") && value && typeof value === "object") {
      merged[key] = mergeDefinedMetadata(previous?.[key] || {}, value);
    } else {
      merged[key] = value;
    }
  }
  return merged;
}

function remoteParentPath(remotePath = "") {
  const clean = cleanPath(remotePath);
  const separator = clean.lastIndexOf("/");
  return separator < 0 ? "" : clean.slice(0, separator);
}

function cachedParentRemotePath(cache, parentId) {
  if (!parentId) return null;
  const parent = cache.itemsById?.[parentId];
  if (parent?.remotePath !== undefined) return cleanPath(parent.remotePath);
  if (Object.hasOwn(cache.pathRootsById || {}, parentId)) return cleanPath(cache.pathRootsById[parentId] || "");
  if (cache.scanRoot?.id === parentId) return cleanPath(cache.scanRoot.remotePath || "");
  return null;
}

function graphParentPath(remotePath = "") {
  const clean = cleanPath(remotePath);
  return `/drive/root:${clean ? `/${clean}` : ""}`;
}

function resolveCachedItemPath(cache, incoming, previous) {
  if (incoming.remotePath !== undefined) return cleanPath(incoming.remotePath);
  const name = incoming.name || previous?.name;
  if (!name) return null;
  const parentId = incoming.parentId;
  const parentPath = cachedParentRemotePath(cache, parentId);
  if (parentPath !== null) return [parentPath, name].filter(Boolean).join("/");
  const sameKnownParent = parentId && previous?.parentId && parentId === previous.parentId;
  const noNewParentEvidence = !parentId;
  if (previous?.remotePath && (sameKnownParent || noNewParentEvidence)) {
    return [remoteParentPath(previous.remotePath), name].filter(Boolean).join("/");
  }
  return null;
}

function updateCachedDescendantPaths(cache, previousPath, nextPath) {
  const oldPath = cleanPath(previousPath || "");
  if (!oldPath) return [];
  const oldLower = cachePathKey(oldPath);
  const changes = [];
  for (const [id, cached] of Object.entries(cache.itemsById || {})) {
    const cachedPath = cleanPath(cached.remotePath || "");
    const cachedLower = cachePathKey(cachedPath);
    if (!cachedPath || !cachedLower.startsWith(`${oldLower}/`)) continue;
    const previous = { ...cached };
    delete cache.pathsByLower[cachedLower];
    if (nextPath === null) {
      delete cached.remotePath;
      delete cached.path;
    } else {
      const suffix = cachedPath.slice(oldPath.length).replace(/^\/+/, "");
      cached.remotePath = [cleanPath(nextPath), suffix].filter(Boolean).join("/");
      cached.path = graphParentPath(remoteParentPath(cached.remotePath));
      cache.pathsByLower[cachePathKey(cached.remotePath)] = id;
    }
    changes.push({ current: cached, previous, removed: [] });
  }
  return changes;
}

function cachePutSimplified(cache, item) {
  if (!item?.id) return { current: null, previous: null, removed: [], descendants: [] };
  if (item.deleted) {
    const previous = cache.itemsById?.[item.id] || null;
    const removed = cacheRemoveItemAndDescendants(cache, item);
    recordSearchTombstones(cache, [previous, ...removed, Object.hasOwn(item, "remotePath") ? item : simplifyItem(item)].filter(Boolean));
    return { current: null, previous, removed, descendants: [] };
  }
  const simplified = Object.hasOwn(item, "remotePath") ? item : simplifyItem(item);
  clearMatchingSearchTombstones(cache, simplified);
  const previous = cache.itemsById?.[simplified.id] || null;
  const current = mergeDefinedMetadata(previous || {}, simplified);
  const eTagChanged = simplified.eTag !== undefined
    && previous?.eTag !== undefined
    && simplified.eTag !== previous.eTag;
  if (eTagChanged && simplified.cTag === undefined) delete current.cTag;
  if (eTagChanged && simplified.file?.hashes === undefined && current.file) delete current.file.hashes;
  const resolvedPath = resolveCachedItemPath(cache, simplified, previous);
  if (resolvedPath === null) {
    delete current.remotePath;
    delete current.path;
  } else {
    current.remotePath = resolvedPath;
    current.path = simplified.path || graphParentPath(remoteParentPath(resolvedPath));
  }
  const removed = [];
  if (previous?.remotePath && previous.remotePath !== current.remotePath) {
    delete cache.pathsByLower[cachePathKey(previous.remotePath)];
  }
  const pathKey = current.remotePath ? cachePathKey(current.remotePath) : null;
  const existingIdForPath = pathKey ? cache.pathsByLower[pathKey] : null;
  if (existingIdForPath && existingIdForPath !== current.id) {
    const displaced = cache.itemsById[existingIdForPath];
    if (displaced) removed.push(...cacheRemoveItemAndDescendants(cache, displaced));
  }
  const descendants = previous?.folder && previous.remotePath !== current.remotePath
    ? updateCachedDescendantPaths(cache, previous.remotePath, current.remotePath ?? null)
    : [];
  cache.itemsById[current.id] = current;
  if (pathKey) cache.pathsByLower[pathKey] = current.id;
  return { current, previous, removed, descendants };
}

function pruneSearchTombstones(cache, now = Date.now()) {
  cache.searchTombstones = (cache.searchTombstones || []).filter((entry) => {
    const recordedAt = Date.parse(entry?.recordedAt || "");
    return entry?.id && Number.isFinite(recordedAt) && now - recordedAt <= searchTombstoneTtlMs;
  }).slice(-2000);
  return cache.searchTombstones;
}

function recordSearchTombstones(cache, items = []) {
  const tombstones = pruneSearchTombstones(cache);
  const byId = new Map(tombstones.map((entry) => [entry.id, entry]));
  const recordedAt = new Date().toISOString();
  for (const item of items) {
    if (!item?.id) continue;
    byId.set(item.id, {
      id: item.id,
      remotePath: cleanPath(item.remotePath || "") || null,
      recordedAt
    });
  }
  cache.searchTombstones = [...byId.values()].slice(-2000);
}

function clearMatchingSearchTombstones(cache, item = {}) {
  const itemPath = cachePathKey(item.remotePath || "");
  cache.searchTombstones = pruneSearchTombstones(cache).filter((entry) => {
    if (entry.id === item.id) return false;
    const tombstonePath = cachePathKey(entry.remotePath || "");
    return !(itemPath && tombstonePath && itemPath === tombstonePath);
  });
}

function isSearchTombstoned(item = {}, tombstones = []) {
  const itemPath = cachePathKey(item.remotePath || "");
  return tombstones.some((entry) => {
    if (entry.id === item.id) return true;
    const deletedPath = cachePathKey(entry.remotePath || "");
    return Boolean(deletedPath && itemPath && (itemPath === deletedPath || itemPath.startsWith(`${deletedPath}/`)));
  });
}

function cacheRemoveItemAndDescendants(cache, item) {
  const old = cache.itemsById?.[item.id] || (item.remotePath ? item : null);
  const remotePath = old?.remotePath || item.remotePath;
  const lowerPath = remotePath ? cachePathKey(remotePath) : null;
  const removed = [];
  const removedIds = new Set([item.id]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const [id, cached] of Object.entries(cache.itemsById || {})) {
      if (removedIds.has(id)) continue;
      const cachedLower = cached.remotePath ? cachePathKey(cached.remotePath) : "";
      if ((lowerPath && (cachedLower === lowerPath || cachedLower.startsWith(`${lowerPath}/`))) || removedIds.has(cached.parentId)) {
        removedIds.add(id);
        changed = true;
      }
    }
  }
  for (const id of removedIds) {
    const cached = cache.itemsById?.[id];
    if (!cached) continue;
    removed.push(cached);
    if (cached.remotePath) delete cache.pathsByLower[cachePathKey(cached.remotePath)];
    delete cache.itemsById[id];
  }
  if (lowerPath) delete cache.pathsByLower[lowerPath];
  return removed;
}

function resolveUnresolvedCachedPaths(cache) {
  const changes = [];
  const maxPasses = Math.max(1, Object.keys(cache.itemsById || {}).length);
  for (let pass = 0; pass < maxPasses; pass += 1) {
    let resolvedThisPass = 0;
    for (const item of Object.values(cache.itemsById || {})) {
      if (item.remotePath !== undefined || !item.name || !item.parentId) continue;
      const parentPath = cachedParentRemotePath(cache, item.parentId);
      if (parentPath === null) continue;
      const previous = { ...item };
      item.remotePath = [parentPath, item.name].filter(Boolean).join("/");
      item.path = graphParentPath(parentPath);
      const pathKey = cachePathKey(item.remotePath);
      const displacedId = cache.pathsByLower[pathKey];
      const removed = [];
      if (displacedId && displacedId !== item.id && cache.itemsById[displacedId]) {
        removed.push(...cacheRemoveItemAndDescendants(cache, cache.itemsById[displacedId]));
      }
      cache.pathsByLower[pathKey] = item.id;
      changes.push({ current: item, previous, removed });
      resolvedThisPass += 1;
    }
    if (!resolvedThisPass) break;
  }
  return changes;
}

async function cacheItems(items = [], metadata = {}, options = {}) {
  const guard = await captureStorageScopeGuard("metadata cache update");
  const lockTimeoutMs = options.lockTimeoutMs ?? (toolProfile === "chatgpt" ? 0 : 10_000);
  const hasMetadata = metadata.deltaLink !== undefined
    || metadata.deltaNextLink !== undefined
    || metadata.deltaTarget !== undefined
    || metadata.scanRoot !== undefined
    || metadata.pathRoot !== undefined;
  if (!items.length && !hasMetadata) return await loadMetadataCache();
  return await runSerialized("metadata", async () => await withFileLock(metadataCacheLockPath, async () => {
    assertStorageScopeGuard(guard, "metadata cache update");
    let cache;
    try {
      cache = await readMetadataCacheFromDisk(guard);
    } catch (error) {
      if (isAccountContextChangedError(error)) throw error;
      recordLocalWarning("metadata cache read", error);
      cache = metadataCacheMemory
        && metadataCacheMemoryGeneration === guard.generation
        && storageScopesEqual(metadataCacheMemory.scope, guard.scope)
        ? metadataCacheMemory
        : emptyMetadataCache(guard.scope);
    }
    if (metadata.pathRoot?.id) {
      cache.pathRootsById ||= {};
      cache.pathRootsById[metadata.pathRoot.id] = cleanPath(metadata.pathRoot.remotePath || "");
    }
    if (metadata.scanRoot !== undefined) {
      cache.scanRoot = metadata.scanRoot;
      if (metadata.scanRoot?.id) {
        cache.pathRootsById ||= {};
        cache.pathRootsById[metadata.scanRoot.id] = cleanPath(metadata.scanRoot.remotePath || "");
      }
    }
    const changes = [];
    for (const item of items) {
      const change = cachePutSimplified(cache, item);
      changes.push(change, ...(change.descendants || []));
    }
    changes.push(...resolveUnresolvedCachedPaths(cache));
    assertStorageScopeGuard(guard, "metadata/content index reconciliation");
    await reconcileContentIndexWithMetadata(changes, guard, { lockTimeoutMs });
    if (metadata.deltaLink !== undefined) {
      cache.deltaLink = metadata.deltaLink || null;
      if (metadata.deltaLink) cache.deltaNextLink = null;
    }
    if (metadata.deltaNextLink !== undefined) cache.deltaNextLink = metadata.deltaNextLink || null;
    if (metadata.deltaTarget !== undefined) cache.deltaTarget = metadata.deltaTarget || null;
    assertStorageScopeGuard(guard, "metadata cache update");
    return await saveMetadataCache(cache, guard);
  }, { timeoutMs: lockTimeoutMs }));
}

async function clearMetadataCache() {
  const guard = await captureStorageScopeGuard("metadata cache clear");
  return await runSerialized("metadata", async () => await withFileLock(metadataCacheLockPath, async () => {
    assertStorageScopeGuard(guard, "metadata cache clear");
    const cleared = emptyMetadataCache(guard.scope);
    metadataCacheLoadPromise = null;
    metadataCacheLoadGeneration = null;
    await ensurePrivateDirectory(cacheRoot);
    await writePrivateFileAtomic(cachePath, JSON.stringify(cleared), {
      beforeCommit: async () => assertStorageScopeGuard(guard, "metadata cache clear commit")
    });
    assertStorageScopeGuard(guard, "metadata cache clear publication");
    metadataCacheMemory = cleared;
    metadataCacheMemoryGeneration = guard.generation;
    metadataCacheFileVersion = await localFileVersion(cachePath);
    assertStorageScopeGuard(guard, "metadata cache clear return");
    return cleared;
  }));
}

function emptyContentIndex(scope = null) {
  return {
    version: 3,
    scope,
    createdAt: new Date().toISOString(),
    updatedAt: null,
    itemCount: 0,
    entriesById: {}
  };
}

async function loadContentIndex() {
  const guard = await captureStorageScopeGuard("content index load");
  if (contentIndexLoadPromise && contentIndexLoadGeneration !== guard.generation) {
    throw accountContextChangedError("content index load");
  }
  if (!contentIndexLoadPromise) {
    let loadPromise;
    loadPromise = (async () => {
      try {
        assertStorageScopeGuard(guard, "content index load");
        const diskVersion = await localFileVersion(contentIndexPath);
        assertStorageScopeGuard(guard, "content index load");
        if (contentIndexMemory
          && contentIndexMemoryGeneration === guard.generation
          && storageScopesEqual(contentIndexMemory.scope, guard.scope)
          && diskVersion === contentIndexFileVersion) return contentIndexMemory;
        let loaded = null;
        await withFileLock(contentIndexLockPath, async () => {
          assertStorageScopeGuard(guard, "content index load");
          const lockedVersion = await localFileVersion(contentIndexPath);
          assertStorageScopeGuard(guard, "content index load");
          if (contentIndexMemory
            && contentIndexMemoryGeneration === guard.generation
            && storageScopesEqual(contentIndexMemory.scope, guard.scope)
            && lockedVersion === contentIndexFileVersion) {
            loaded = contentIndexMemory;
            return;
          }
          loaded = await readContentIndexFromDisk(guard);
          assertStorageScopeGuard(guard, "content index publication");
          contentIndexMemory = loaded;
          contentIndexMemoryGeneration = guard.generation;
          contentIndexFileVersion = await localFileVersion(contentIndexPath);
        });
        assertStorageScopeGuard(guard, "content index return");
        return loaded || contentIndexMemory;
      } catch (error) {
        if (isAccountContextChangedError(error)) throw error;
        recordLocalWarning("content index read", error);
        assertStorageScopeGuard(guard, "content index fallback publication");
        contentIndexMemory = emptyContentIndex(guard.scope);
        contentIndexMemoryGeneration = guard.generation;
        contentIndexFileVersion = null;
      } finally {
        if (contentIndexLoadPromise === loadPromise) {
          contentIndexLoadPromise = null;
          contentIndexLoadGeneration = null;
        }
      }
      assertStorageScopeGuard(guard, "content index return");
      return contentIndexMemory;
    })();
    contentIndexLoadPromise = loadPromise;
    contentIndexLoadGeneration = guard.generation;
  }
  const result = await contentIndexLoadPromise;
  assertStorageScopeGuard(guard, "content index return");
  if (!storageScopesEqual(result?.scope, guard.scope)) throw accountContextChangedError("content index return");
  return result;
}

async function readContentIndexFromDisk(existingGuard = null) {
  const guard = existingGuard || await captureStorageScopeGuard("content index read");
  assertStorageScopeGuard(guard, "content index read");
  await ensurePrivateDirectory(cacheRoot);
  try {
    const parsed = JSON.parse(await readFile(contentIndexPath, "utf8"));
    await hardenPrivateFile(contentIndexPath);
    assertStorageScopeGuard(guard, "content index read");
    if (parsed.version !== 3 || !storageScopesEqual(parsed.scope, guard.scope)) return emptyContentIndex(guard.scope);
    return {
      ...emptyContentIndex(guard.scope),
      ...parsed,
      version: 3,
      scope: guard.scope,
      entriesById: parsed.entriesById || {}
    };
  } catch (error) {
    if (error?.code === "ENOENT") {
      assertStorageScopeGuard(guard, "content index read");
      return emptyContentIndex(guard.scope);
    }
    throw error;
  }
}

async function writeContentIndex(index, existingGuard = null) {
  const guard = existingGuard || await captureStorageScopeGuard("content index write");
  assertStorageScopeGuard(guard, "content index write");
  if (index.scope && !storageScopesEqual(index.scope, guard.scope)) {
    throw new Error("Content index scope changed during this operation. Refusing to persist cross-account state.");
  }
  index.version = 3;
  index.scope = guard.scope;
  index.updatedAt = new Date().toISOString();
  index.itemCount = Object.keys(index.entriesById || {}).length;
  await ensurePrivateDirectory(cacheRoot);
  await writePrivateFileAtomic(contentIndexPath, JSON.stringify(index), {
    beforeCommit: async () => assertStorageScopeGuard(guard, "content index commit")
  });
  assertStorageScopeGuard(guard, "content index publication");
  contentIndexMemory = index;
  contentIndexMemoryGeneration = guard.generation;
  contentIndexFileVersion = await localFileVersion(contentIndexPath);
  assertStorageScopeGuard(guard, "content index return");
  return index;
}

async function saveContentIndex(index, existingGuard = null) {
  const guard = existingGuard || await captureStorageScopeGuard("content index update");
  return await runSerialized("content", async () => await withFileLock(contentIndexLockPath, async () => {
    assertStorageScopeGuard(guard, "content index update");
    let latest;
    try {
      latest = await readContentIndexFromDisk(guard);
    } catch (error) {
      if (isAccountContextChangedError(error)) throw error;
      recordLocalWarning("content index read", error);
      latest = contentIndexMemory
        && contentIndexMemoryGeneration === guard.generation
        && storageScopesEqual(contentIndexMemory.scope, guard.scope)
        ? contentIndexMemory
        : emptyContentIndex(guard.scope);
    }
    const merged = {
      ...latest,
      entriesById: {
        ...(latest.entriesById || {}),
        ...(index.entriesById || {})
      }
    };
    assertStorageScopeGuard(guard, "content index update");
    return await writeContentIndex(merged, guard);
  }));
}

async function clearContentIndex() {
  const guard = await captureStorageScopeGuard("content index clear");
  return await runSerialized("content", async () => await withFileLock(contentIndexLockPath, async () => {
    assertStorageScopeGuard(guard, "content index clear");
    const cleared = emptyContentIndex(guard.scope);
    contentIndexLoadPromise = null;
    contentIndexLoadGeneration = null;
    await ensurePrivateDirectory(cacheRoot);
    await writePrivateFileAtomic(contentIndexPath, JSON.stringify(cleared), {
      beforeCommit: async () => assertStorageScopeGuard(guard, "content index clear commit")
    });
    assertStorageScopeGuard(guard, "content index clear publication");
    contentIndexMemory = cleared;
    contentIndexMemoryGeneration = guard.generation;
    contentIndexFileVersion = await localFileVersion(contentIndexPath);
    assertStorageScopeGuard(guard, "content index clear return");
    return cleared;
  }));
}

function indexedItemMetadataChanged(left = {}, right = {}) {
  return left.id !== right.id
    || left.name !== right.name
    || left.remotePath !== right.remotePath
    || left.path !== right.path
    || left.webUrl !== right.webUrl
    || left.size !== right.size
    || left.createdDateTime !== right.createdDateTime
    || left.lastModifiedDateTime !== right.lastModifiedDateTime
    || left.eTag !== right.eTag
    || left.cTag !== right.cTag
    || left.parentId !== right.parentId
    || left.driveId !== right.driveId
    || Boolean(left.deleted) !== Boolean(right.deleted);
}

async function reconcileContentIndexWithMetadata(changes = [], existingGuard = null, options = {}) {
  if (!changes.length) return;
  const guard = existingGuard || await captureStorageScopeGuard("metadata/content index reconciliation");
  const lockTimeoutMs = options.lockTimeoutMs ?? (toolProfile === "chatgpt" ? 0 : 10_000);
  await runSerialized("content", async () => await withFileLock(contentIndexLockPath, async () => {
    assertStorageScopeGuard(guard, "metadata/content index reconciliation");
    let index;
    try {
      index = await readContentIndexFromDisk(guard);
    } catch (error) {
      if (isAccountContextChangedError(error)) throw error;
      recordLocalWarning("content index read", error);
      index = contentIndexMemory
        && contentIndexMemoryGeneration === guard.generation
        && storageScopesEqual(contentIndexMemory.scope, guard.scope)
        ? contentIndexMemory
        : emptyContentIndex(guard.scope);
    }
    let changed = false;
    for (const change of changes) {
      for (const removed of change?.removed || []) {
        if (removed?.id && Object.hasOwn(index.entriesById, removed.id)) {
          delete index.entriesById[removed.id];
          changed = true;
        }
      }
      const current = change?.current;
      if (!current?.id) continue;
      const entry = index.entriesById[current.id];
      if (!entry) continue;
      if (!contentIndexEntryFresh(entry, current)) {
        delete index.entriesById[current.id];
        changed = true;
        continue;
      }
      if (indexedItemMetadataChanged(entry.item, current)) {
        entry.item = current;
        changed = true;
      }
    }
    assertStorageScopeGuard(guard, "content index reconciliation publication");
    if (changed) {
      await writeContentIndex(index, guard);
    } else {
      contentIndexMemory = index;
      contentIndexMemoryGeneration = guard.generation;
    }
  }, { timeoutMs: lockTimeoutMs }));
}

function contentIndexEntryFresh(entry, item) {
  if (!entry) return false;
  if (entry.cTag && item.cTag) return entry.cTag === item.cTag;
  return entry.eTag === item.eTag
    && entry.lastModifiedDateTime === item.lastModifiedDateTime
    && entry.size === item.size;
}

function officeExportIndexable(item = {}) {
  const extension = extname(item.name || "").toLowerCase();
  const mimeType = (item.file?.mimeType || "").toLowerCase();
  return Object.values(officeKinds).some((kind) => kind.extensions.has(extension) || kind.mimeTypes.has(mimeType));
}

function contentIndexableReason(item = {}, args = {}) {
  if (!item?.file || item.folder) return { ok: false, reason: "not-file" };
  const settings = pluginSettings();
  const maxBytes = clampInteger(args.maxBytesPerFile, settings.maxIndexedFileSize, 1024, textFileLimit);
  const isOffice = officeExportIndexable(item);
  const openXmlKind = officePackageKindFromName(item.name);
  if (args.officeOnly === true && !isOffice) return { ok: false, reason: "not-office" };
  if (isOffice && openXmlKind && args.includeOfficeStructured !== false) {
    const officeIndexLimit = clampInteger(args.maxOfficePackageBytes, 50 * 1024 * 1024, 1024, maxOfficePackageBytes);
    if (item.size && item.size > officeIndexLimit) return { ok: false, reason: "office-package-too-large" };
    return { ok: true, source: "office-openxml", kind: openXmlKind };
  }
  if (item.size && item.size > maxBytes) return { ok: false, reason: "too-large" };
  const allowedExtensions = normalizeExtensions(args.extensions?.length ? args.extensions : settings.supportedIndexedFileTypes);
  const extension = extname(item.name || "").toLowerCase();
  if (allowedExtensions.has(extension) || isLikelyTextItem(item)) return { ok: true, source: "text-read" };
  if ((args.includeOfficeExport === true || settings.includeOfficeTextExport) && officeExportIndexable(item)) {
    return { ok: true, source: "graph-text-export" };
  }
  return { ok: false, reason: "unsupported-type" };
}

function officeIndexSegments(document, maxSegments = 50_000) {
  const segments = [];
  const add = (text, anchor) => {
    const value = String(text ?? "").trim();
    if (!value || segments.length >= maxSegments) return;
    segments.push({ text: value, anchor });
  };
  if (document.kind === "word") {
    for (const paragraph of document.paragraphs || []) add(paragraph.text, { type: "paragraph", part: paragraph.part, index: paragraph.index, style: paragraph.style || null });
    for (const table of document.tables || []) {
      for (const [rowIndex, row] of (table.rows || []).entries()) {
        for (const [columnIndex, value] of row.entries()) add(value, { type: "tableCell", part: table.part, tableIndex: table.index, rowIndex, columnIndex });
      }
    }
    for (const control of document.contentControls || []) add(control.text, { type: "contentControl", part: control.part, index: control.index, id: control.id, tag: control.tag });
    for (const comment of document.comments || []) add(comment.text, { type: "comment", id: comment.id, author: comment.author });
  } else if (document.kind === "excel") {
    for (const sheet of document.sheets || []) {
      for (const cell of sheet.cells || []) {
        add(cell.value, { type: "cell", sheet: sheet.name, address: cell.address, formula: cell.formula || null });
        if (cell.formula) add(cell.formula, { type: "formula", sheet: sheet.name, address: cell.address });
      }
      for (const table of sheet.tables || []) add(`${table.name || table.displayName || "table"} ${table.ref || ""}`, { type: "table", sheet: sheet.name, name: table.name, ref: table.ref });
    }
    for (const name of document.definedNames || []) add(`${name.name || ""} ${name.value || ""}`, { type: "definedName", name: name.name });
  } else if (document.kind === "powerpoint") {
    const addShape = (shape, slideIndex, parentShapeId = null) => {
      add(shape.text, { type: "shape", slideIndex, shapeId: shape.id, name: shape.name, parentShapeId });
      for (const [rowIndex, row] of (shape.table?.rows || []).entries()) {
        for (const [columnIndex, value] of row.entries()) add(value, { type: "tableCell", slideIndex, shapeId: shape.id, rowIndex, columnIndex });
      }
      for (const child of shape.children || []) addShape(child, slideIndex, shape.id);
    };
    for (const slide of document.slides || []) {
      for (const shape of slide.shapes || []) addShape(shape, slide.index);
      add(slide.notes, { type: "notes", slideIndex: slide.index });
    }
  }
  return { segments, truncated: segments.length >= maxSegments };
}

async function extractIndexText(item, args = {}) {
  const settings = pluginSettings();
  const maxBytes = clampInteger(args.maxBytesPerFile, settings.maxIndexedFileSize, 1024, textFileLimit);
  const indexable = contentIndexableReason(item, args);
  if (!indexable.ok) throw new Error(indexable.reason);
  if (indexable.source === "office-openxml") {
    const document = await inspectRemoteOfficePackage({
      itemId: item.id,
      maxParagraphs: clampInteger(args.maxOfficeParagraphs, 10_000, 1, 10_000),
      maxCells: clampInteger(args.maxOfficeCells, 50_000, 1, 50_000),
      maxSlides: clampInteger(args.maxOfficeSlides, 5_000, 1, 5_000),
      includeCells: true,
      strictRelationships: true
    }, indexable.kind, "inspect");
    const structured = officeIndexSegments(document, clampInteger(args.maxOfficeSegments, 50_000, 1, 100_000));
    const text = structured.segments.map((entry) => entry.text).join("\n");
    return {
      text,
      segments: structured.segments,
      structuredKind: document.kind,
      bytesRead: Number(item.size || 0),
      truncated: structured.truncated || Boolean(document.truncated),
      source: "office-openxml"
    };
  }
  const target = { itemId: item.id };
  const sourcePath = indexable.source === "graph-text-export"
    ? `${contentPath(target)}?${new URLSearchParams({ format: "text" }).toString()}`
    : contentPath(target);
  const limited = await graphLimitedBuffer(sourcePath, maxBytes);
  if (indexable.source === "text-read") assertNoBinaryNulls(limited.buffer);
  return {
    text: limited.buffer.toString("utf8").replace(/\uFFFD$/u, ""),
    bytesRead: limited.bytesRead,
    truncated: limited.truncated,
    source: indexable.source
  };
}

function contentIndexTextFields(text = "") {
  return {
    normalizedText: normalizeFindText(text),
    tokens: findTokens(text)
  };
}

function contentSearchContext(query) {
  return {
    query,
    normalizedQuery: normalizeFindText(query),
    tokens: findImportantTokens(query)
  };
}

function compareContentMatches(left, right) {
  if (right.score !== left.score) return right.score - left.score;
  return String(left.item?.name || "").localeCompare(String(right.item?.name || ""));
}

function insertTopContentMatch(matches, candidate, maxResults) {
  matches.push(candidate);
  matches.sort(compareContentMatches);
  if (matches.length > maxResults) matches.pop();
}

function contentMatchForQuery(entry, queryOrContext) {
  const context = typeof queryOrContext === "object" && queryOrContext
    ? queryOrContext
    : contentSearchContext(queryOrContext);
  const text = String(entry.text || "");
  const normalizedText = entry.normalizedText || normalizeFindText(text);
  const normalizedQuery = context.normalizedQuery;
  const tokens = context.tokens || [];
  const textTokenSet = new Set(Array.isArray(entry.tokens) ? entry.tokens : findTokens(text));
  const tokenMatches = tokens.filter((token) => textTokenSet.has(token));
  if (!normalizedQuery && !tokenMatches.length) return null;
  const exactPhrase = normalizedQuery
    ? (/\s|[-.]/.test(normalizedQuery) ? normalizedText.includes(normalizedQuery) : textTokenSet.has(normalizedQuery))
    : false;
  const exactIndex = exactPhrase ? normalizedText.indexOf(normalizedQuery) : -1;
  if (exactIndex < 0 && tokenMatches.length === 0) return null;

  const lowerText = text.toLowerCase();
  const rawNeedle = exactIndex >= 0 ? normalizedQuery : tokenMatches[0];
  const rawIndex = rawNeedle ? lowerText.indexOf(rawNeedle.toLowerCase()) : -1;
  const start = rawIndex >= 0 ? Math.max(0, rawIndex - 80) : 0;
  const end = rawIndex >= 0 ? Math.min(text.length, rawIndex + rawNeedle.length + 120) : Math.min(text.length, 200);
  const snippet = text.slice(start, end).replace(/\s+/g, " ").trim();
  const coverage = tokens.length ? tokenMatches.length / tokens.length : 0;
  const segmentMatches = [];
  for (const segment of entry.segments || []) {
    const segmentText = String(segment.text || "");
    const normalizedSegment = normalizeFindText(segmentText);
    const segmentTokens = new Set(findTokens(segmentText));
    const phrase = normalizedQuery && normalizedSegment.includes(normalizedQuery);
    const matched = tokens.filter((token) => segmentTokens.has(token));
    if (!phrase && !matched.length) continue;
    segmentMatches.push({
      anchor: segment.anchor,
      snippet: segmentText.slice(0, 500),
      score: (phrase ? 70 : 35) + Math.round((tokens.length ? matched.length / tokens.length : 0) * 35)
    });
  }
  segmentMatches.sort((a, b) => b.score - a.score);
  return {
    score: (exactIndex >= 0 ? 70 : 35) + Math.round(coverage * 35),
    matchedTokens: tokenMatches.length,
    snippet: segmentMatches[0]?.snippet || snippet,
    anchor: segmentMatches[0]?.anchor || null,
    segmentMatches: segmentMatches.slice(0, 10),
    reasons: [
      exactIndex >= 0 ? "indexed content phrase" : "indexed content tokens",
      ...(tokenMatches.length ? [`content token matches: ${tokenMatches.slice(0, 5).join(", ")}`] : [])
    ]
  };
}

async function contentSearch(args = {}) {
  const startedAt = Date.now();
  const query = String(args.query || "").trim();
  if (!query) throw new Error("query is required.");
  const maxResults = clampInteger(args.maxResults, 20, 1, 100);
  const index = await loadContentIndex();
  const context = contentSearchContext(query);
  const matches = [];
  let matched = 0;
  for (const entry of Object.values(index.entriesById || {})) {
    if (args.officeOnly === true && entry.source !== "office-openxml") continue;
    const match = contentMatchForQuery(entry, context);
    if (!match) continue;
    matched += 1;
    insertTopContentMatch(matches, {
      item: entry.item,
      score: match.score,
      matchedTokens: match.matchedTokens,
      snippet: match.snippet,
      anchor: match.anchor,
      segmentMatches: match.segmentMatches,
      reasons: match.reasons,
      source: entry.source,
      indexedAt: entry.indexedAt,
      truncated: entry.truncated
    }, maxResults);
  }
  return {
    query,
    indexPath: contentIndexPath,
    itemCount: index.itemCount || 0,
    matched,
    returned: matches.length,
    durationMs: elapsedMs(startedAt),
    items: matches.map((match) => ({
      ...formatSimplifiedItem(match.item, args.format),
      score: match.score,
      matchedTokens: match.matchedTokens,
      reasons: match.reasons,
      snippet: match.snippet,
      anchor: match.anchor,
      segmentMatches: match.segmentMatches,
      source: match.source,
      indexedAt: match.indexedAt,
      truncated: match.truncated
    }))
  };
}

async function officeIndexRefresh(args = {}) {
  return await contentIndexRefresh({
    ...args,
    refreshMetadata: args.refreshMetadata !== false,
    metadataMode: args.metadataMode || "auto",
    includeOfficeStructured: true,
    officeOnly: true
  });
}

async function officeContentSearch(args = {}) {
  let result = await contentSearch({ ...args, officeOnly: true });
  if (!result.items.length) return result;
  const staleIds = [];
  for (const item of result.items) {
    try {
      const index = await loadContentIndex();
      const entry = index.entriesById?.[item.id];
      const current = simplifyItem(await getRawInfo({ itemId: item.id }));
      if (entry && !contentIndexEntryFresh(entry, current)) staleIds.push(item.id);
      await bestEffortLocalWrite("Office search metadata freshness update", async () => await cacheItems([current]));
    } catch (error) {
      staleIds.push(item.id);
      await bestEffortLocalWrite("Office search stale index removal", async () => {
        const index = await loadContentIndex();
        if (index.entriesById?.[item.id]) {
          delete index.entriesById[item.id];
          await saveContentIndex(index);
        }
      });
      recordLocalWarning("Office search metadata freshness check", error);
    }
  }
  if (staleIds.length) {
    result = await contentSearch({ ...args, officeOnly: true });
    result.staleEntriesRemoved = [...new Set(staleIds)];
  }
  return result;
}

async function contentIndexRefresh(args = {}) {
  const startedAt = Date.now();
  const settings = pluginSettings();
  if (!settings.contentIndexEnabled) throw new Error("Content indexing is disabled by configuration.");
  if (args.refreshMetadata === true) {
    await cacheRefresh({
      ...args,
      mode: args.metadataMode || "auto",
      maxItems: clampInteger(args.scanMaxItems, 10000, 1, 50000),
      maxFolders: clampInteger(args.scanMaxFolders, 2000, 1, 10000),
      maxDepth: clampInteger(args.scanMaxDepth, settings.maxScanDepth, 0, 50)
    });
  }

  const maxFiles = clampInteger(args.maxFiles, 100, 1, 1000);
  let sourceItems = [];
  if (args.itemId || args.path || args.preset) {
    sourceItems = [simplifyItem(await getRawInfo(args))];
  } else {
    sourceItems = await cachedItems();
  }

  const index = await loadContentIndex();
  const eligibleItems = sourceItems
    .filter((item) => item?.id && !item.deleted)
    .filter((item) => contentIndexableReason(item, args).ok);
  const candidates = eligibleItems
    .map((item) => ({
      item,
      fresh: Boolean(args.force !== true && contentIndexEntryFresh(index.entriesById[item.id], item))
    }))
    .sort((left, right) => Number(left.fresh) - Number(right.fresh))
    .slice(0, maxFiles)
    .map((entry) => entry.item);
  const results = {
    indexPath: contentIndexPath,
    considered: sourceItems.length,
    eligible: eligibleItems.length,
    selected: candidates.length,
    indexed: 0,
    reused: 0,
    skipped: sourceItems.length - candidates.length,
    graphContentReadsAttempted: 0,
    failed: 0,
    failures: []
  };

  await mapWithConcurrency(candidates, clampInteger(args.concurrencyLimit, settings.concurrencyLimit, 1, 8), async (item) => {
    const existing = index.entriesById[item.id];
    if (args.force !== true && contentIndexEntryFresh(existing, item)) {
      index.entriesById[item.id] = { ...existing, item };
      results.reused += 1;
      return;
    }
    try {
      results.graphContentReadsAttempted += 1;
      const extracted = await extractIndexText(item, args);
      const textFields = contentIndexTextFields(extracted.text);
      index.entriesById[item.id] = {
        item,
        text: extracted.text,
        normalizedText: textFields.normalizedText,
        tokens: textFields.tokens,
        indexedAt: new Date().toISOString(),
        source: extracted.source,
        bytesRead: extracted.bytesRead,
        textBytes: Buffer.byteLength(extracted.text, "utf8"),
        truncated: extracted.truncated,
        segments: extracted.segments || [],
        structuredKind: extracted.structuredKind || null,
        eTag: item.eTag,
        cTag: item.cTag,
        lastModifiedDateTime: item.lastModifiedDateTime,
        size: item.size
      };
      results.indexed += 1;
    } catch (error) {
      results.failed += 1;
      if (results.failures.length < 10) {
        results.failures.push({ id: item.id, name: item.name, error: safeToolErrorMessage(error) });
      }
    }
  });

  const updatedEntries = Object.fromEntries(
    candidates
      .filter((item) => index.entriesById[item.id])
      .map((item) => [item.id, index.entriesById[item.id]])
  );
  const persistedIndex = await saveContentIndex({ entriesById: updatedEntries });
  return {
    ...results,
    itemCount: persistedIndex.itemCount,
    durationMs: elapsedMs(startedAt),
    settings: {
      maxBytesPerFile: clampInteger(args.maxBytesPerFile, settings.maxIndexedFileSize, 1024, textFileLimit),
      supportedIndexedFileTypes: args.extensions?.length ? [...normalizeExtensions(args.extensions)] : settings.supportedIndexedFileTypes,
      includeOfficeExport: args.includeOfficeExport === true || settings.includeOfficeTextExport,
      includeOfficeStructured: args.includeOfficeStructured !== false,
      concurrencyLimit: clampInteger(args.concurrencyLimit, settings.concurrencyLimit, 1, 8)
    },
    note: "Content indexing reads file bodies only during this explicit refresh. Normal find/search calls reuse the local index and do not fetch content."
  };
}

async function cacheMovedOrRenamedItem(previous, current) {
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([current]));
}

function metadataCacheAgeSeconds(cache) {
  if (!cache?.updatedAt) return false;
  const updatedAt = Date.parse(cache.updatedAt);
  if (!Number.isFinite(updatedAt)) return false;
  return Math.max(0, (Date.now() - updatedAt) / 1000);
}

function metadataCacheFresh(cache, settings = pluginSettings()) {
  const ageSeconds = metadataCacheAgeSeconds(cache);
  if (ageSeconds === false) return false;
  return settings.cacheTtlSeconds === 0 || ageSeconds <= settings.cacheTtlSeconds;
}

function metadataCacheWithinStaleWindow(cache, maxAgeSeconds = chatgptStaleCacheMaxAgeSeconds) {
  const ageSeconds = metadataCacheAgeSeconds(cache);
  return ageSeconds !== false && ageSeconds <= maxAgeSeconds;
}

function unresolvedPathItems(cache) {
  return Object.values(cache.itemsById || {}).filter((item) => item?.id && !item.deleted && item.remotePath === undefined);
}

async function syncStatus(args = {}) {
  const cache = await loadMetadataCache();
  const contentIndex = await loadContentIndex();
  const settings = pluginSettings();
  const items = Object.values(cache.itemsById || {});
  const contentEntries = Object.values(contentIndex.entriesById || {});
  const unresolvedItems = unresolvedPathItems(cache);
  const cacheAgeSeconds = cache.updatedAt ? Math.max(0, Math.round((Date.now() - Date.parse(cache.updatedAt)) / 1000)) : null;
  return {
    cachePath,
    contentIndexPath,
    downloadRoot,
    updateRoot,
    backupRoot,
    storageRoot,
    itemCount: items.length,
    updatedAt: cache.updatedAt,
    cacheAgeSeconds,
    cacheFresh: metadataCacheFresh(cache, settings),
    deltaLinkAvailable: Boolean(cache.deltaLink),
    deltaNextLinkAvailable: Boolean(cache.deltaNextLink),
    deltaTarget: cache.deltaTarget || null,
    scanRoot: cache.scanRoot,
    unresolvedPathCount: unresolvedItems.length,
    unresolvedPathSamples: args.includeSamples
      ? unresolvedItems.slice(0, 10).map((item) => ({ id: item.id, name: item.name, parentId: item.parentId }))
      : undefined,
    contentIndex: {
      itemCount: contentEntries.length,
      structuredOfficeCount: contentEntries.filter((entry) => entry.source === "office-openxml").length,
      segmentCount: contentEntries.reduce((sum, entry) => sum + (entry.segments?.length || 0), 0),
      updatedAt: contentIndex.updatedAt,
      enabled: settings.contentIndexEnabled
    },
    settings,
    samples: args.includeSamples ? items.slice(0, 10).map((item) => formatSimplifiedItem(item, "compact")) : undefined
  };
}

async function cachedItems() {
  const cache = await loadMetadataCache();
  return Object.values(cache.itemsById || {});
}

async function cachedItemByPath(path) {
  const cache = await loadMetadataCache();
  const id = cache.pathsByLower?.[cachePathKey(path)];
  return id ? cache.itemsById[id] : null;
}

async function cachedItemById(id) {
  const cache = await loadMetadataCache();
  return id ? cache.itemsById[id] || null : null;
}

async function mapWithConcurrency(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    for (;;) {
      const index = next;
      next += 1;
      if (index >= items.length) return;
      results[index] = await fn(items[index], index);
    }
  });
  await Promise.all(workers);
  return results;
}

async function doctor(args = {}) {
  const checks = [];
  const addCheck = (name, status, details = {}) => {
    checks.push({ name, status, details });
  };
  const runCheck = async (name, fn) => {
    try {
      addCheck(name, "pass", await fn());
    } catch (error) {
      addCheck(name, "fail", { error: safeToolErrorMessage(error) });
    }
  };

  const cfgStatus = publicConfig();
  addCheck("config", cfgStatus.clientIdConfigured && !cfgStatus.configReadError ? "pass" : "fail", {
    clientIdConfigured: cfgStatus.clientIdConfigured,
    tenant: cfgStatus.tenant,
    scopes: cfgStatus.scopes,
    authenticationStore: cfgStatus.authenticationStore,
    storedCredentialConfigured: cfgStatus.storedCredentialConfigured,
    keychainTokenConfigured: cfgStatus.keychainTokenConfigured,
    configPath: cfgStatus.configPath,
    ...(cfgStatus.configReadError ? { configReadError: cfgStatus.configReadError } : {})
  });

  await runCheck("access token", async () => {
    await getAccessToken();
    return { accessTokenAvailable: true };
  });
  await runCheck("profile", async () => {
    const profile = await graph("/me");
    return {
      displayName: profile.displayName,
      userPrincipalName: profile.userPrincipalName,
      mail: profile.mail
    };
  });
  await runCheck("drive", async () => {
    const drive = await graph("/me/drive");
    return {
      id: drive.id,
      name: drive.name,
      driveType: drive.driveType,
      quotaState: drive.quota?.state
    };
  });
  if (args.checkPresets === false) {
    addCheck("presets", "pass", {
      pathPresets: cfgStatus.pathPresets,
      validated: false,
      note: "Preset resolution checks were explicitly skipped."
    });
  } else {
    const presetEntries = Object.entries(cfgStatus.pathPresets || {});
    const presetChecks = await mapWithConcurrency(
      presetEntries,
      Math.min(pluginSettings().concurrencyLimit, 4),
      async ([name, path]) => {
        try {
          const item = await getRawInfo({ path, cacheResults: false });
          return {
            name,
            path,
            available: true,
            item: {
              id: item.id,
              name: item.name,
              type: item.folder ? "folder" : item.file ? "file" : "item"
            }
          };
        } catch (error) {
          return {
            name,
            path,
            available: false,
            graphStatus: error?.graphStatus,
            error: safeToolErrorMessage(error)
          };
        }
      }
    );
    const missing = presetChecks.filter((entry) => entry.graphStatus === 404);
    const failed = presetChecks.filter((entry) => !entry.available && entry.graphStatus !== 404);
    addCheck("presets", failed.length ? "fail" : missing.length ? "warn" : "pass", {
      pathPresets: cfgStatus.pathPresets,
      validated: true,
      availableCount: presetChecks.filter((entry) => entry.available).length,
      missingCount: missing.length,
      failedCount: failed.length,
      presets: presetChecks,
      ...(missing.length ? {
        recommendation: "Update pathPresets in config.json so every preset points to an existing root-relative OneDrive folder."
      } : {})
    });
  }

  if (args.checkRootList !== false) {
    await runCheck("root list", async () => {
      const result = await list({ limit: Math.min(args.rootListLimit ?? 5, 20) });
      return {
        returned: result.items.length,
        nextLink: Boolean(result.nextLink),
        sample: result.items.slice(0, 3).map((item) => ({ name: item.name, type: item.type }))
      };
    });
  }

  const failCount = checks.filter((check) => check.status === "fail").length;
  const warnCount = checks.filter((check) => check.status === "warn").length;
  return {
    ok: failCount === 0,
    status: failCount ? "fail" : warnCount ? "warn" : "pass",
    summary: { total: checks.length, pass: checks.length - failCount - warnCount, warn: warnCount, fail: failCount },
    checks,
    note: failCount
      ? "At least one OneDrive health check failed. Fix the failed check before relying on file tools."
      : warnCount
        ? "OneDrive plugin health checks completed with warnings. Review the warning checks before relying on affected workflows."
        : "OneDrive plugin health checks passed."
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

function normalizedScopeSet(scopes = "") {
  return [...new Set(String(scopes).trim().split(/\s+/).filter(Boolean))].sort().join(" ");
}

function tokenMatchesAuthConfig(token = {}, cfg = config()) {
  if (token.auth_client_id && token.auth_client_id !== cfg.clientId) return false;
  const tenantMatches = !token.auth_tenant
    || token.auth_tenant === cfg.tenant
    || (cfg.tenant === "common" && token.auth_tenant === "consumers");
  if (!tenantMatches) return false;
  if (token.auth_scopes && normalizedScopeSet(token.auth_scopes) !== normalizedScopeSet(cfg.scopes)) return false;
  const defaultCfg = config();
  const requestedOverride = cfg.clientId !== defaultCfg.clientId
    || cfg.tenant !== defaultCfg.tenant
    || normalizedScopeSet(cfg.scopes) !== normalizedScopeSet(defaultCfg.scopes);
  return !requestedOverride || Boolean(token.auth_client_id && token.auth_tenant && token.auth_scopes);
}

async function postForm(url, params) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
    signal: timeoutSignal()
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
  let reauthenticationReason = args.forceReauth === true ? "explicitly-forced" : "no-stored-credential";
  if (args.forceReauth !== true) {
    try {
      await getAccessToken(cfg);
      const vault = authVault(config());
      const storedCredentialConfigured = Boolean(getKeychainToken(config())?.refresh_token);
      return {
        authenticated: true,
        alreadyAuthenticated: true,
        deviceCodeIssued: false,
        authenticationStore: vault.mode,
        storedCredentialConfigured,
        keychainTokenConfigured: vault.mode === "keychain" && storedCredentialConfigured,
        message: "Existing OneDrive authentication is healthy. No device code was issued."
      };
    } catch (error) {
      if (isMissingStoredAuthenticationError(error)) {
        reauthenticationReason = "no-stored-credential";
      } else if (isReauthenticationRequiredError(error)) {
        reauthenticationReason = "stored-credential-rejected";
      } else {
        throw new Error(
          `Existing OneDrive authentication could not be verified, so device login was not started. ${safeToolErrorMessage(error)} `
          + "Retry the token check after the temporary problem clears, or use forceReauth: true only when the user explicitly wants to sign in again."
        );
      }
    }
  }
  const generation = authGeneration;
  const deviceGeneration = ++deviceLoginGeneration;
  const { body: result, tenant } = await postFormWithConsumerFallback("device", cfg, {
    client_id: cfg.clientId,
    scope: cfg.scopes
  });
  if (generation !== authGeneration || deviceGeneration !== deviceLoginGeneration) {
    throw new Error("OneDrive authentication state changed while device login was starting. Start login again.");
  }
  pendingDevice = { ...result, tenant, scopes: cfg.scopes, startedAt: Date.now(), deviceGeneration };
  return {
    userCode: result.user_code,
    verificationUri: result.verification_uri,
    verificationUriComplete: result.verification_uri_complete,
    expiresIn: result.expires_in,
    interval: result.interval,
    message: result.message,
    authTenant: tenant,
    reauthenticationReason,
    deviceCodeStoredInMemory: true
  };
}

async function pollDeviceLogin(args = {}) {
  const generation = authGeneration;
  const pending = pendingDevice;
  const deviceCode = args.deviceCode || pending?.device_code;
  if (!deviceCode) throw new Error("No pending device code. Run onedrive_auth_device_start first.");
  if (pending && args.deviceCode && args.deviceCode !== pending.device_code) {
    throw new Error("The supplied device code does not match the active OneDrive device-login session. Start login again.");
  }
  const deviceGeneration = pending?.deviceGeneration ?? deviceLoginGeneration;
  const cfg = config({ tenant: pending?.tenant, scopes: pending?.scopes });
  requireClientId(cfg);
  const tokenParams = {
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    client_id: cfg.clientId,
    device_code: deviceCode
  };
  const tokenResponse = await postFormWithConsumerFallback("token", cfg, tokenParams).catch((error) => {
    if (error.body?.error === "authorization_pending") {
      return { authorizationPending: true, message: "Authorization is still pending. Try again after the user completes browser sign-in." };
    }
    if (error.body?.error === "slow_down") {
      return { authorizationPending: true, slowDown: true, message: "Microsoft asked polling to slow down. Try again in a few more seconds." };
    }
    throw error;
  });
  const result = tokenResponse.body || tokenResponse;
  if (generation !== authGeneration
    || deviceGeneration !== deviceLoginGeneration
    || (pending && pendingDevice?.deviceGeneration !== deviceGeneration)) {
    throw new Error("OneDrive authentication state changed while device login was pending. Start login again.");
  }
  if (result.authorizationPending) return result;
  authGeneration += 1;
  invalidateActiveStorageScope();
  tokenRefreshPromise = null;
  tokenCache = normalizeToken({
    ...result,
    auth_client_id: cfg.clientId,
    auth_tenant: tokenResponse.tenant || cfg.tenant,
    auth_scopes: cfg.scopes,
    auth_context_id: randomUUID()
  });
  setKeychainToken(tokenCache, cfg);
  pendingDevice = null;
  adoptCurrentToolAccountGeneration();
  return {
    authenticated: true,
    authTenant: tokenResponse.tenant || cfg.tenant,
    tokenType: tokenCache.token_type,
    expiresAt: tokenCache.expires_at ? new Date(tokenCache.expires_at).toISOString() : null,
    authenticationStore: authVault(cfg).mode,
    refreshTokenStored: Boolean(tokenCache.refresh_token),
    refreshTokenStoredInKeychain: authVault(cfg).mode === "keychain" && Boolean(tokenCache.refresh_token)
  };
}

async function refreshAccessToken(refreshToken, cfg = config(), generation = authGeneration, authContextId = null) {
  requireClientId(cfg);
  const { body: result, tenant } = await postFormWithConsumerFallback("token", cfg, {
    grant_type: "refresh_token",
    client_id: cfg.clientId,
    refresh_token: refreshToken,
    scope: cfg.scopes
  });
  if (generation !== authGeneration) {
    throw new Error("OneDrive authentication state changed while the access token was refreshing. Try again if access is still intended.");
  }
  tokenCache = normalizeToken({
    ...result,
    auth_client_id: cfg.clientId,
    auth_tenant: tenant,
    auth_scopes: cfg.scopes,
    refresh_token: result.refresh_token || refreshToken,
    auth_context_id: authContextId || randomUUID()
  });
  setKeychainToken(tokenCache, cfg);
  return tokenCache;
}

async function getAccessToken(cfg = config()) {
  const requestToken = toolCallContext.getStore()?.graphAccessToken;
  if (requestToken) return requestToken;
  if (process.env.ONEDRIVE_TEST_ACCESS_TOKEN) return process.env.ONEDRIVE_TEST_ACCESS_TOKEN;
  requireClientId(cfg);
  let current = tokenCache || getKeychainToken(cfg);
  if (!current?.refresh_token && !current?.access_token) {
    throw new Error("OneDrive is not authenticated. Run onedrive_auth_device_start, complete browser login, then run onedrive_auth_device_poll.");
  }
  if (!tokenMatchesAuthConfig(current, cfg)) {
    throw new Error("Stored OneDrive authentication does not match the requested client, tenant, or scopes. Start device-code login for the requested authentication context.");
  }
  if (!current.auth_context_id) {
    current = normalizeAuthContextToken(current);
    tokenCache = current;
    setKeychainToken(current, cfg);
  }
  const expiresAt = current.expires_at || 0;
  if (current.access_token && expiresAt - Date.now() > 60_000) {
    tokenCache = current;
    return current.access_token;
  }
  if (!current.refresh_token) throw new Error("Stored token has no refresh token. Run device-code login again.");
  if (!tokenRefreshPromise) {
    const generation = authGeneration;
    tokenRefreshPromise = refreshAccessToken(current.refresh_token, cfg, generation, current.auth_context_id);
  }
  const refresh = tokenRefreshPromise;
  try {
    const refreshed = await refresh;
    return refreshed.access_token;
  } finally {
    if (tokenRefreshPromise === refresh) tokenRefreshPromise = null;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function elapsedMs(startedAt) {
  return Math.max(0, Date.now() - startedAt);
}

function fetchTimeoutMs() {
  const value = Number(process.env.ONEDRIVE_FETCH_TIMEOUT_MS || 60_000);
  return Number.isFinite(value) && value > 0 ? value : 60_000;
}

function timeoutSignal(timeoutMs = fetchTimeoutMs()) {
  if (AbortSignal.timeout) return AbortSignal.timeout(timeoutMs);
  const controller = new AbortController();
  setTimeout(() => controller.abort(), timeoutMs).unref?.();
  return controller.signal;
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
  if (method === "GET" || method === "HEAD") return explicitMaxRetries ?? 4;
  if (explicitMaxRetries !== undefined && isReplayableBody(fetchOptions.body)) return explicitMaxRetries;
  return 0;
}

function shouldDefaultJsonContentType(body) {
  return body
    && typeof body === "string"
    && body.trim().startsWith("{");
}

function contentLength(response) {
  const value = Number(response.headers.get("content-length"));
  return Number.isFinite(value) && value >= 0 ? value : null;
}

async function parseResponseBody(response) {
  if (response.status === 204) return null;
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) return await response.json().catch(() => ({}));
  const buffer = await response.arrayBuffer();
  if (contentType.startsWith("text/") || (!response.ok && buffer.byteLength <= 4096)) {
    return new TextDecoder().decode(buffer);
  }
  return buffer;
}

async function fetchWithRetry(url, options = {}, retryOptions = {}) {
  const maxRetries = retryOptions.maxRetries ?? 3;
  const timeoutMs = retryOptions.timeoutMs ?? fetchTimeoutMs();
  for (let attempt = 0; ; attempt += 1) {
    let response;
    try {
      response = await fetch(url, { ...options, signal: options.signal || timeoutSignal(timeoutMs) });
    } catch (error) {
      if (attempt >= maxRetries) {
        if (error.name === "AbortError" || error.name === "TimeoutError") {
          throw new Error(`Request timed out after ${timeoutMs}ms: ${url}`);
        }
        const transportError = new Error(`Microsoft Graph transport error after ${attempt + 1} attempts: ${error.message || String(error)}`);
        transportError.cause = error;
        throw transportError;
      }
      await sleep(Math.min(1000 * 2 ** attempt, 8000));
      continue;
    }
    if (!shouldRetryResponse(response) || attempt >= maxRetries) return response;
    await parseResponseBody(response).catch(() => null);
    await sleep(retryDelayMs(response, attempt));
  }
}

function graphRequestId(headers, body) {
  return headers?.get?.("request-id")
    || headers?.get?.("x-ms-request-id")
    || headers?.get?.("client-request-id")
    || body?.error?.innerError?.["request-id"]
    || body?.error?.innerError?.requestId
    || null;
}

function currentGraphRequestId(options = {}) {
  const store = toolCallContext.getStore();
  return options.mutation === true
    ? store?.lastMutationGraphRequestId || null
    : store?.lastGraphRequestId || null;
}

function rememberGraphRequestId(requestId, options = {}) {
  const store = toolCallContext.getStore();
  if (!store) return;
  store.lastGraphRequestId = requestId || null;
  if (options.mutation === true) store.lastMutationGraphRequestId = requestId || null;
}

function microsoftGraphError(body, response) {
  const code = body?.error?.code ? `${body.error.code}: ` : "";
  const textBody = typeof body === "string" ? body.trim() : "";
  const message = body?.error?.message || textBody || `${response.status} ${response.statusText}`;
  const requestId = graphRequestId(response.headers, body);
  const suffix = requestId ? ` (request-id: ${requestId})` : "";
  const error = new Error(`Microsoft Graph error: ${code}${message}${suffix}`);
  error.graphRequestId = requestId;
  error.graphStatus = response.status;
  return error;
}

function graphBaseUrl() {
  return process.env.ONEDRIVE_GRAPH_BASE_URL || "https://graph.microsoft.com/v1.0";
}

function graphUrl(path, options = {}) {
  const base = new URL(graphBaseUrl());
  if (!String(path).startsWith("http")) return `${base.toString().replace(/\/+$/, "")}${path}`;
  const target = new URL(path);
  if (options.skipAuth) return target.toString();
  const basePath = base.pathname.replace(/\/+$/, "");
  const targetAllowed = target.origin === base.origin
    && (target.pathname === basePath || target.pathname.startsWith(`${basePath}/`));
  if (!targetAllowed) {
    throw new Error(`Refusing to send an authenticated Microsoft Graph request to an untrusted URL: ${target.origin}${target.pathname}`);
  }
  return target.toString();
}

function assertTrustedCopyMonitorUrl(monitorUrl) {
  const base = new URL(graphBaseUrl());
  const target = new URL(monitorUrl, base);
  if (target.origin === base.origin) return target.toString();
  const host = target.hostname.toLowerCase();
  const trustedMicrosoftHost = host === "my.microsoftpersonalcontent.com"
    || host.endsWith(".microsoftpersonalcontent.com")
    || host.endsWith(".sharepoint.com");
  if (target.protocol !== "https:" || !trustedMicrosoftHost) {
    throw new Error(`Refusing to poll untrusted copy monitor URL: ${target.origin}${target.pathname}`);
  }
  return target.toString();
}

function assertTrustedUploadSessionUrl(uploadUrl) {
  const base = new URL(graphBaseUrl());
  const target = new URL(uploadUrl, base);
  if (target.origin === base.origin) return target.toString();
  const host = target.hostname.toLowerCase();
  const trustedMicrosoftHost = host.endsWith(".up.1drv.com")
    || host.endsWith(".up.1drv.ms")
    || host.endsWith(".sharepoint.com")
    || host.endsWith(".microsoftpersonalcontent.com");
  if (target.protocol !== "https:" || !trustedMicrosoftHost) {
    throw new Error(`Refusing to upload file contents to an untrusted upload session URL: ${target.origin}${target.pathname}`);
  }
  return target.toString();
}

function captureGraphAccountContext(operation = "Microsoft Graph request") {
  assertToolAccountGeneration(operation);
  return {
    authGeneration,
    storageScopeGeneration,
    authContextId: currentAuthContextId()
  };
}

function assertGraphAccountContext(snapshot, operation = "Microsoft Graph request") {
  assertToolAccountGeneration(operation);
  if (snapshot.authGeneration !== authGeneration || snapshot.storageScopeGeneration !== storageScopeGeneration) {
    throw accountContextChangedError(operation);
  }
  if (snapshot.authContextId && currentAuthContextId() !== snapshot.authContextId) {
    throw accountContextChangedError(operation);
  }
}

async function graph(path, options = {}) {
  const { returnResponse = false, maxRetries, skipAuth = false, isMutation, ...fetchOptions } = options;
  const method = String(fetchOptions.method || "GET").toUpperCase();
  const mutationRequest = isMutation ?? !["GET", "HEAD", "OPTIONS"].includes(method);
  const accountContext = captureGraphAccountContext("Microsoft Graph request");
  const accessToken = skipAuth ? null : await getAccessToken();
  if (!skipAuth && !accountContext.authContextId) accountContext.authContextId = currentAuthContextId();
  assertGraphAccountContext(accountContext, "Microsoft Graph request");
  const url = graphUrl(path, { skipAuth });
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
  assertGraphAccountContext(accountContext, "Microsoft Graph response");
  const requestId = graphRequestId(retriedResponse.headers, body);
  rememberGraphRequestId(requestId, { mutation: mutationRequest });
  if (returnResponse) {
    return { body, headers: retriedResponse.headers, status: retriedResponse.status, ok: retriedResponse.ok, graphRequestId: requestId };
  }
  if (!retriedResponse.ok) {
    throw microsoftGraphError(body, retriedResponse);
  }
  return body;
}

async function graphDownloadToFile(path, target, options = {}) {
  const accountContext = captureGraphAccountContext("Microsoft Graph download");
  const accessToken = await getAccessToken();
  if (!accountContext.authContextId) accountContext.authContextId = currentAuthContextId();
  assertGraphAccountContext(accountContext, "Microsoft Graph download");
  const url = graphUrl(path);
  const response = await fetchWithRetry(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers || {})
    }
  }, { maxRetries: 3 });
  assertGraphAccountContext(accountContext, "Microsoft Graph download response");
  if (!response.ok) {
    const body = await parseResponseBody(response);
    throw microsoftGraphError(body, response);
  }
  await mkdir(dirname(target), { recursive: true, mode: 0o700 });
  const temp = `${target}.part-${process.pid}-${randomUUID()}`;
  let bytesWritten = 0;
  let published = false;
  try {
    if (response.body) {
      const counter = new TransformStream({
        transform(chunk, controller) {
          bytesWritten += chunk.byteLength;
          controller.enqueue(chunk);
        }
      });
      await pipeline(Readable.fromWeb(response.body.pipeThrough(counter)), createWriteStream(temp, { flags: "wx", mode: 0o600 }));
    } else {
      const buffer = Buffer.from(await response.arrayBuffer());
      bytesWritten = buffer.length;
      await writePrivateFile(temp, buffer);
    }
    assertGraphAccountContext(accountContext, "Microsoft Graph download publication");
    await renameFile(temp, target);
    published = true;
    assertGraphAccountContext(accountContext, "Microsoft Graph download publication");
    await hardenPrivateFile(target);
    assertGraphAccountContext(accountContext, "Microsoft Graph download completion");
  } catch (error) {
    await rm(temp, { force: true });
    if ((published && isAccountContextChangedError(error)) || options.reserved === true) await rm(target, { force: true });
    throw error;
  }
  return { bytesWritten: bytesWritten || contentLength(response) || 0 };
}

async function graphLimitedBuffer(path, maxBytes, options = {}) {
  const accountContext = captureGraphAccountContext("Microsoft Graph bounded content read");
  const accessToken = await getAccessToken();
  if (!accountContext.authContextId) accountContext.authContextId = currentAuthContextId();
  assertGraphAccountContext(accountContext, "Microsoft Graph bounded content read");
  const url = graphUrl(path);
  const limit = Math.max(1, maxBytes);
  const response = await fetchWithRetry(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Range: `bytes=0-${limit}`
    }
  }, { maxRetries: 3 });
  assertGraphAccountContext(accountContext, "Microsoft Graph bounded content response");
  if (!response.ok && response.status !== 206) {
    const body = await parseResponseBody(response);
    throw microsoftGraphError(body, response);
  }

  if (!response.body?.getReader) {
    const buffer = Buffer.from(await response.arrayBuffer());
    assertGraphAccountContext(accountContext, "Microsoft Graph bounded content return");
    return {
      buffer: buffer.subarray(0, limit),
      bytesRead: buffer.length,
      truncated: buffer.length > limit
    };
  }

  const reader = response.body.getReader();
  const chunks = [];
  let bytesRead = 0;
  let truncated = false;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = Buffer.from(value);
      bytesRead += chunk.length;
      const remaining = limit + 1 - chunks.reduce((sum, entry) => sum + entry.length, 0);
      if (remaining > 0) chunks.push(chunk.subarray(0, remaining));
      if (bytesRead > limit) {
        truncated = true;
        await reader.cancel().catch(() => {});
        break;
      }
    }
  } finally {
    reader.releaseLock?.();
  }

  const buffer = Buffer.concat(chunks);
  assertGraphAccountContext(accountContext, "Microsoft Graph bounded content return");
  return {
    buffer: buffer.subarray(0, limit),
    bytesRead,
    truncated: truncated || bytesRead > limit || buffer.length > limit
  };
}

function cleanPath(path = "") {
  return String(path).replace(/^\/+/, "").replace(/\/+$/, "");
}

function clampInteger(value, defaultValue, min, max) {
  const number = Number(value ?? defaultValue);
  if (!Number.isInteger(number)) return defaultValue;
  return Math.min(Math.max(number, min), max);
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

function assertSafeItemName(name = "", label = "name") {
  if (typeof name !== "string") throw new Error(`${label} must be a string.`);
  const value = String(name);
  const trimmed = value.trim();
  if (!trimmed) throw new Error(`${label} must not be empty.`);
  if (trimmed === "." || trimmed === "..") throw new Error(`${label} must not be "." or "..".`);
  if (/[\/\\]/.test(value)) throw new Error(`${label} must be a single item name, not a path.`);
  if (/[\u0000-\u001F]/.test(value)) throw new Error(`${label} contains control characters.`);
  return trimmed;
}

function assertAtMostOneSelector(args = {}, label, selectors = []) {
  const provided = selectors.filter((selector) =>
    selector.keys.every((key) => Object.hasOwn(args, key) && args[key] !== undefined && args[key] !== null && args[key] !== "")
  );
  if (provided.length > 1) {
    throw new Error(`${label} must use at most one selector: ${selectors.map((selector) => selector.label).join(" or ")}.`);
  }
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
  if (args.remoteRelativePath !== undefined && !args.remotePreset) {
    throw new Error("remoteRelativePath requires remotePreset.");
  }
  if (args.remotePreset && !args.remoteRelativePath) {
    throw new Error("remoteRelativePath is required when remotePreset is used, and must include the destination filename.");
  }
  const resolved = resolvePresetPath(args, {
    pathField: "remotePath",
    presetField: "remotePreset",
    relativeField: "remoteRelativePath",
    allowEmpty: false
  });
  if (args.remotePreset) splitRemotePath(resolved);
  return resolved;
}

function normalizeLocalPathForCompare(path) {
  return resolve(path).replace(/\/+$/g, "");
}

function isLocalOneDriveSyncPath(path) {
  const normalized = normalizeLocalPathForCompare(path);
  return localOneDriveSyncRoots.some((root) => {
    const normalizedRoot = normalizeLocalPathForCompare(root.path);
    return normalized === normalizedRoot
      || normalized.startsWith(`${normalizedRoot}/`)
      || (root.prefix && normalized.startsWith(normalizedRoot));
  });
}

function assertNotLocalOneDriveSyncPath(path, operation, args = {}) {
  if (args.allowLocalOneDriveSyncPath === true) return;
  if (!isLocalOneDriveSyncPath(path)) return;
  throw new Error(`${operation} refuses to use a local OneDrive sync folder by default: ${path}. Use remote OneDrive plugin tools instead, or pass allowLocalOneDriveSyncPath: true only when you explicitly need local sync-folder access.`);
}

async function existingRealPath(path) {
  try {
    return await realpath(path);
  } catch {
    return null;
  }
}

async function realPathForPotentialWrite(path) {
  const target = resolve(path);
  const existingTarget = await existingRealPath(target);
  if (existingTarget) return existingTarget;

  let parent = dirname(target);
  while (parent && parent !== dirname(parent)) {
    const existingParent = await existingRealPath(parent);
    if (existingParent) return join(existingParent, relative(parent, target));
    parent = dirname(parent);
  }
  return target;
}

async function assertNotLocalOneDriveSyncPathForRead(path, operation, args = {}) {
  assertNotLocalOneDriveSyncPath(path, operation, args);
  const real = await existingRealPath(path);
  if (real) assertNotLocalOneDriveSyncPath(real, operation, args);
}

async function assertNotLocalOneDriveSyncPathForWrite(path, operation, args = {}) {
  assertNotLocalOneDriveSyncPath(path, operation, args);
  const real = await realPathForPotentialWrite(path);
  assertNotLocalOneDriveSyncPath(real, operation, args);
}

function formatDriveItem(item, format = "compact") {
  const simplified = simplifyItem(item);
  return formatSimplifiedItem(simplified, format);
}

function formatSimplifiedItem(simplified, format = "compact") {
  if (!simplified || format === "full") return simplified;
  return {
    id: simplified.id,
    name: simplified.name,
    remotePath: simplified.remotePath,
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

function itemIdBase(itemId) {
  return `/me/drive/items/${encodeURIComponent(itemId)}`;
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

function hasExpectedIdentity(args = {}) {
  return Boolean(args.expectedName || args.expectedId);
}

function stablePreviewValue(value) {
  if (Array.isArray(value)) return value.map(stablePreviewValue);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, stablePreviewValue(value[key])])
      .filter(([, child]) => child !== undefined)
  );
}

function previewProofHash(tool, proof = {}) {
  return createHash("sha256")
    .update(tool)
    .update("\n")
    .update(JSON.stringify(stablePreviewValue(proof)))
    .digest("hex");
}

function previewProofItemIds(value, key = "") {
  const ids = new Set();
  const visit = (child, childKey) => {
    if (Array.isArray(child)) return child.forEach((entry) => visit(entry, childKey));
    if (child && typeof child === "object") return Object.entries(child).forEach(([nestedKey, entry]) => visit(entry, nestedKey));
    if (typeof child === "string" && /(?:^|_)(?:item|source|draft|target)?id$/i.test(childKey)) ids.add(child);
  };
  visit(value, key);
  return [...ids];
}

function cleanupPreviewTokens(now = Date.now()) {
  for (const [token, entry] of previewTokens.entries()) {
    if (!entry || entry.expiresAt <= now) previewTokens.delete(token);
  }
}

function issuePreviewToken(tool, proof = {}) {
  assertToolAccountGeneration(`${tool} preview issuance`);
  cleanupPreviewTokens();
  const scope = toolCallContext.getStore()?.storageScope;
  if (!scope?.authContextId || !scope?.driveId) {
    throw new Error("OneDrive storage scope is unavailable. Refusing to issue an unscoped preview token.");
  }
  const token = randomUUID();
  const expiresAt = Date.now() + previewTokenTtlMs;
  previewTokens.set(token, {
    tool,
    proofHash: previewProofHash(tool, proof),
    scopeKey: storageScopeKey(scope),
    authGeneration,
    storageScopeGeneration,
    itemIds: previewProofItemIds(proof),
    expiresAt
  });
  return {
    previewToken: token,
    previewTokenExpiresAt: new Date(expiresAt).toISOString()
  };
}

function previewWithToken(preview, tool, proof = {}) {
  if (preview?.dryRun !== true) return preview;
  return {
    ...preview,
    ...issuePreviewToken(tool, proof)
  };
}

function consumePreviewToken(tool, proof = {}, token) {
  assertToolAccountGeneration(`${tool} preview consumption`);
  cleanupPreviewTokens();
  if (!token) return { ok: false, reason: "missing" };
  const entry = previewTokens.get(token);
  if (!entry) return { ok: false, reason: "not_found_or_expired" };
  const currentScopeKey = storageScopeKey(toolCallContext.getStore()?.storageScope);
  if (!currentScopeKey
    || entry.scopeKey !== currentScopeKey
    || entry.authGeneration !== authGeneration
    || entry.storageScopeGeneration !== storageScopeGeneration) {
    return { ok: false, reason: "scope_mismatch" };
  }
  if (entry.tool !== tool || entry.proofHash !== previewProofHash(tool, proof)) {
    return { ok: false, reason: "mismatch" };
  }
  previewTokens.delete(token);
  return { ok: true };
}

function previewTokenRequiredResult(preview, tool, proof, token, requiredField) {
  const result = consumePreviewToken(tool, proof, token);
  if (result.ok) return null;
  return {
    ...preview,
    dryRun: false,
    confirmed: true,
    previewTokenRequired: true,
    previewTokenStatus: result.reason,
    [requiredField]: "Run a dry-run preview for this exact operation and pass the returned previewToken with dryRun: false and confirmed: true."
  };
}

function batchMutationWarnings(extra = []) {
  return [partialBatchMutationWarning, ...extra.filter(Boolean)];
}

function itemAuditSummary(item) {
  const simplified = item?.remotePath !== undefined ? item : simplifyItem(item);
  if (!simplified) return null;
  return {
    id: simplified.id,
    name: simplified.name,
    remotePath: simplified.remotePath,
    type: simplified.folder ? "folder" : simplified.file ? "file" : "item",
    size: simplified.size,
    lastModifiedDateTime: simplified.lastModifiedDateTime
  };
}

function itemMutationBase(rawItem) {
  if (!rawItem?.id) throw new Error("Resolved item is missing an ID; refusing live mutation.");
  return itemIdBase(rawItem.id);
}

function mutationMatchHeaders(rawItem) {
  return rawItem?.eTag ? { "If-Match": rawItem.eTag } : {};
}

function itemVersionProof(rawItem = {}) {
  return {
    id: rawItem.id,
    name: rawItem.name,
    eTag: rawItem.eTag || null,
    cTag: rawItem.cTag || null,
    size: Number.isFinite(rawItem.size) ? rawItem.size : null,
    lastModifiedDateTime: rawItem.lastModifiedDateTime || null
  };
}

function permissionAuditSummary(permission = {}) {
  return {
    id: permission.id,
    roles: permission.roles,
    link: permission.link ? {
      type: permission.link.type,
      scope: permission.link.scope,
      preventsDownload: permission.link.preventsDownload
    } : undefined,
    grantedToPresent: Boolean(permission.grantedTo || permission.grantedToV2),
    grantedToIdentityCount: Array.isArray(permission.grantedToIdentities) ? permission.grantedToIdentities.length : undefined,
    grantedToIdentityV2Count: Array.isArray(permission.grantedToIdentitiesV2) ? permission.grantedToIdentitiesV2.length : undefined,
    invitationPresent: Boolean(permission.invitation),
    invitationSignInRequired: permission.invitation?.signInRequired,
    inheritedFrom: permission.inheritedFrom ? itemAuditSummary(permission.inheritedFrom) : undefined,
    expirationDateTime: permission.expirationDateTime,
    hasPassword: permission.hasPassword
  };
}

function permissionDiffAuditSummary(diff = {}) {
  return {
    added: (diff.added || []).map(permissionAuditSummary),
    removed: (diff.removed || []).map(permissionAuditSummary),
    unchangedCount: diff.unchangedCount || 0,
    beforeCount: diff.beforeCount || 0,
    afterCount: diff.afterCount || 0
  };
}

function safeErrorInfo(error) {
  return {
    message: redactAuditText(error?.message || String(error)),
    graphRequestId: error?.graphRequestId || currentGraphRequestId({ mutation: true }) || currentGraphRequestId() || undefined,
    graphStatus: error?.graphStatus
  };
}

function redactAuditText(text = "") {
  return String(text)
    .replace(/https?:\/\/[^\s")]+/gi, "[redacted-url]")
    .replace(/\bBearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [redacted]")
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, "[redacted-email]")
    .replace(/\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi, "[redacted-id]")
    .replace(/\b(access[_-]?token|refresh[_-]?token|id[_-]?token|secret)\s*[:=]\s*([^\s,;)]+)/gi, "$1=[redacted]")
    .replace(/\b(email|alias|objectId|recipient)\s*[:=]\s*([^\s,;)]+)/gi, "$1=[redacted]");
}

function safeToolErrorMessage(error) {
  return redactAuditText(error?.message || String(error));
}

function recordLocalWarning(operation, error) {
  const store = toolCallContext.getStore();
  if (!store) return;
  store.localWarnings.push({ operation, error: safeToolErrorMessage(error) });
}

async function bestEffortLocalWrite(operation, fn) {
  try {
    return await fn();
  } catch (error) {
    recordLocalWarning(operation, error);
    return null;
  }
}

function sanitizeAuditValue(value) {
  if (Array.isArray(value)) return value.map(sanitizeAuditValue);
  if (!value || typeof value !== "object") return value;
  const clean = {};
  const forbidden = new Set([
    "access_token", "refresh_token", "id_token", "Authorization", "authorization",
    "content", "body", "uploadUrl", "monitorUrl", "resourceLocation", "webUrl",
    "password", "recipients", "email", "alias", "objectId"
  ]);
  for (const [key, entry] of Object.entries(value)) {
    if (forbidden.has(key)) continue;
    clean[key] = sanitizeAuditValue(entry);
  }
  return clean;
}

async function writeMutationAudit(tool, entry) {
  const record = sanitizeAuditValue({
    timestamp: new Date().toISOString(),
    tool,
    ...entry,
    graphRequestId: entry.graphRequestId || currentGraphRequestId({ mutation: true }) || currentGraphRequestId() || undefined
  });
  return await bestEffortLocalWrite("mutation audit write", async () => {
    await withFileLock(auditLockPath, async () => {
      await ensurePrivateDirectory(auditRoot);
      await appendFile(auditPath, `${JSON.stringify(record)}\n`, { encoding: "utf8", mode: 0o600 });
      await hardenPrivateFile(auditPath);
    });
    return record;
  });
}

async function auditRecent(args = {}) {
  let text = "";
  await withFileLock(auditLockPath, async () => {
    await ensurePrivateDirectory(auditRoot);
    try {
      text = await readFile(auditPath, "utf8");
      await hardenPrivateFile(auditPath);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  });
  const limit = clampInteger(args.limit, 50, 1, 500);
  const since = args.since ? Date.parse(args.since) : null;
  const until = args.until ? Date.parse(args.until) : null;
  if (args.since && Number.isNaN(since)) throw new Error("since must be an ISO timestamp.");
  if (args.until && Number.isNaN(until)) throw new Error("until must be an ISO timestamp.");
  const pathNeedle = args.pathContains ? String(args.pathContains).toLowerCase() : "";
  const entryPathText = (entry) => JSON.stringify({
    target: entry.target,
    before: entry.before,
    after: entry.after
  }).toLowerCase();
  let entries = text.trim()
    ? text.trim().split("\n").map((line) => {
        try {
          return JSON.parse(line);
        } catch {
          return { malformed: true, line };
        }
      })
    : [];
  entries = entries.filter((entry) => {
    if (args.tool && entry.tool !== args.tool) return false;
    if (args.status && entry.status !== args.status) return false;
    const timestamp = entry.timestamp ? Date.parse(entry.timestamp) : null;
    if (since !== null && (!timestamp || timestamp < since)) return false;
    if (until !== null && (!timestamp || timestamp > until)) return false;
    if (pathNeedle && !entryPathText(entry).includes(pathNeedle)) return false;
    return true;
  });
  if (args.newestFirst !== false) entries.reverse();
  entries = entries.slice(0, limit);
  return {
    auditPath,
    count: entries.length,
    filters: {
      tool: args.tool || null,
      status: args.status || null,
      pathContains: args.pathContains || null,
      since: args.since || null,
      until: args.until || null,
      newestFirst: args.newestFirst !== false
    },
    entries
  };
}

async function auditExport(args = {}) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const preferredTarget = args.localPath ? resolve(args.localPath) : join(auditRoot, `export-${stamp}-${randomUUID()}.jsonl`);
  await assertNotLocalOneDriveSyncPathForWrite(preferredTarget, "Audit export", args);
  const reservation = await reserveLocalDestination(preferredTarget, {
    overwrite: args.overwrite === true,
    allowAlternate: !args.localPath
  });
  const target = reservation.path;
  try {
    await withFileLock(auditLockPath, async () => {
      try {
        await copyFile(auditPath, target);
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
        await writePrivateFile(target, "");
      }
    });
    await hardenPrivateFile(target);
    const written = await stat(target);
    return { auditPath, localPath: target, bytesWritten: written.size };
  } catch (error) {
    if (reservation.reserved) await rm(target, { force: true });
    throw error;
  }
}

async function auditClear(args = {}) {
  if (args.confirmed !== true) {
    return {
      confirmed: false,
      auditPath,
      requiredToClear: "Set confirmed: true after explicit user confirmation to clear the local mutation audit log."
    };
  }
  await withFileLock(auditLockPath, async () => await rm(auditPath, { force: true }));
  return { confirmed: true, auditPath, cleared: true };
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
  const source = item.remoteItem ? {
    ...item.remoteItem,
    id: item.remoteItem.id || item.id,
    name: item.remoteItem.name || item.name,
    webUrl: item.remoteItem.webUrl || item.webUrl,
    parentReference: item.remoteItem.parentReference || item.parentReference
  } : item;
  const extension = extname(source.name || "").toLowerCase();
  const reportedMimeType = source.file?.mimeType;
  const normalizedMimeType = canonicalTextMimeTypes.get(extension)
    || (reportedMimeType === "application/octet-stream" ? undefined : reportedMimeType);
  const numericSize = Number(source.size);
  return {
    id: source.id,
    name: source.name,
    remotePath: itemRemotePath(source),
    path: source.parentReference?.path ? decodeGraphPath(source.parentReference.path) : source.parentReference?.path,
    parentId: source.parentReference?.id,
    driveId: source.parentReference?.driveId ? String(source.parentReference.driveId).toUpperCase() : source.parentReference?.driveId,
    webUrl: source.webUrl,
    size: Number.isFinite(numericSize) && numericSize >= 0 ? numericSize : null,
    createdDateTime: source.createdDateTime,
    lastModifiedDateTime: source.lastModifiedDateTime,
    eTag: source.eTag,
    cTag: source.cTag,
    deleted: source.deleted,
    folder: source.folder ? { childCount: source.folder.childCount } : undefined,
    file: source.file ? {
      mimeType: normalizedMimeType || reportedMimeType,
      ...(source.file.hashes !== undefined ? { hashes: source.file.hashes } : {})
    } : undefined
  };
}

function itemRemotePath(item) {
  if (!item?.name) return undefined;
  if (!item.parentReference && !item.root) return undefined;
  const parentPath = item.parentReference?.path;
  if (!parentPath && item.parentReference?.id) return undefined;
  const rootMatch = String(parentPath || "").match(/^\/(?:drive|drives\/[^/]+)\/root:(.*)$/);
  const parentRemotePath = rootMatch
    ? decodeGraphPath(rootMatch[1].replace(/^\/+|\/+$/g, ""))
    : "";
  return [parentRemotePath, item.name].filter(Boolean).join("/");
}

function decodeGraphPath(path = "") {
  try {
    return decodeURIComponent(path);
  } catch {
    return path;
  }
}

async function list(args = {}) {
  const params = new URLSearchParams();
  params.set("$top", String(clampInteger(args.limit, 100, 1, 200)));
  params.set("$select", args.select || defaultSelect);
  const result = await graph(`${childrenPath(args)}?${params.toString()}`);
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(result.value || []));
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), nextLink: result["@odata.nextLink"] || null };
}

function isDeltaCursor(value) {
  if (!value) return false;
  try {
    const base = new URL(graphBaseUrl());
    const target = new URL(String(value), base);
    const basePath = base.pathname.replace(/\/+$/, "");
    const trustedPath = target.origin === base.origin
      && (target.pathname === basePath || target.pathname.startsWith(`${basePath}/`));
    return trustedPath && target.pathname.split("/").some((segment) => segment === "delta" || segment.startsWith("delta("));
  } catch {
    return false;
  }
}

async function collectPages(firstPath, maxItems, format = "compact", formatter = formatDriveItem, options = {}) {
  return await withMetadataCacheBatch(async () => {
    const items = [];
    let nextPath = firstPath;
    let nextLink = null;
    let deltaLink = null;
    let truncated = false;
    const seenPages = new Set();
    let pagesFetched = 0;
    let unsafePageTruncation = false;
    const pendingCacheItems = [];
    let pendingCursorMetadata = {};
    const safetyMaxPages = Math.max(1, Math.ceil(maxItems / 1) + 100);
    const requestedMaxPages = options.maxPages === undefined ? null : clampInteger(options.maxPages, 1, 1, 100);
    let maxPagesReached = false;
    while (nextPath && items.length < maxItems) {
      if (seenPages.has(nextPath)) throw new Error(`Microsoft Graph pagination cycle detected at ${safeDisplayPath(nextPath)}.`);
      if (pagesFetched >= safetyMaxPages) throw new Error(`Microsoft Graph pagination exceeded ${safetyMaxPages} pages before reaching the item limit.`);
      seenPages.add(nextPath);
      pagesFetched += 1;
      let result;
      try {
        result = await graph(nextPath);
      } catch (error) {
        error.graphPagesFetched = pagesFetched;
        throw error;
      }
      const pageItems = result.value || [];
      const remaining = maxItems - items.length;
      const acceptedItems = pageItems.slice(0, remaining);
      const pageTruncated = pageItems.length > remaining;
      const pageNextLink = result["@odata.nextLink"] || null;
      const pageDeltaLink = result["@odata.deltaLink"] || null;
      const cursorMetadata = options.persistDeltaCursor === true && !pageTruncated && (pageDeltaLink || pageNextLink) ? {
        deltaLink: pageDeltaLink || undefined,
        deltaNextLink: pageNextLink && !pageDeltaLink ? pageNextLink : undefined,
        deltaTarget: options.deltaTarget
      } : {};
      const preparedItems = options.prepareItems ? await options.prepareItems(acceptedItems) : acceptedItems;
      const cacheableItems = options.prepareCacheItems
        ? await options.prepareCacheItems(preparedItems)
        : preparedItems;
      if (options.cacheResults !== false) pendingCacheItems.push(...cacheableItems);
      if (Object.keys(cursorMetadata).length) pendingCursorMetadata = cursorMetadata;
      const outputFormatter = options.formatter || formatter;
      items.push(...preparedItems.map((item) => outputFormatter(item, format)));
      if (pageTruncated) {
        unsafePageTruncation = true;
        nextLink = null;
      } else {
        nextLink = pageNextLink;
      }
      deltaLink = pageTruncated ? null : pageDeltaLink;
      nextPath = !pageTruncated && nextLink && items.length < maxItems ? nextLink : null;
      truncated = truncated || pageTruncated || (Boolean(nextLink) && items.length >= maxItems);
      if (requestedMaxPages !== null && pagesFetched >= requestedMaxPages && nextPath) {
        maxPagesReached = true;
        truncated = true;
        nextPath = null;
      }
    }
    if (pendingCacheItems.length || Object.keys(pendingCursorMetadata).length) {
      const deduplicatedCacheItems = [...new Map(pendingCacheItems.filter((item) => item?.id).map((item) => [item.id, item])).values()];
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(deduplicatedCacheItems, pendingCursorMetadata));
    }
    return {
      items,
      nextLink,
      deltaLink,
      truncated,
      unsafePageTruncation,
      pagesFetched,
      ...(requestedMaxPages !== null ? { maxPagesReached } : {}),
      count: items.length
    };
  });
}

function safeDisplayPath(value = "") {
  try {
    const url = new URL(value, graphBaseUrl());
    return `${url.origin}${url.pathname}`;
  } catch {
    return String(value).split("?")[0];
  }
}

async function listAll(args = {}) {
  const maxItems = clampInteger(args.maxItems, 1000, 1, 5000);
  const params = new URLSearchParams();
  params.set("$top", String(Math.min(clampInteger(args.pageSize, 200, 1, 200), maxItems)));
  params.set("$select", args.select || defaultSelect);
  return await collectPages(`${childrenPath(args)}?${params.toString()}`, maxItems, args.format, formatDriveItem, {
    cacheResults: args.cacheResults !== false
  });
}

function normalizeExtensions(extensions = []) {
  return new Set(
    extensions
      .map((extension) => String(extension || "").trim().toLowerCase())
      .filter(Boolean)
      .map((extension) => extension.startsWith(".") ? extension : `.${extension}`)
  );
}

function scanItemMatches(item, args = {}, extensionFilter = new Set()) {
  const isFolder = Boolean(item.folder);
  const isFile = Boolean(item.file);
  if (isFile && args.includeFiles === false) return false;
  if (isFolder && args.includeFolders === false) return false;
  if (!isFile && !isFolder && (args.includeFiles === false || args.includeFolders === false)) return false;

  if (args.nameContains) {
    const needle = String(args.nameContains).toLowerCase();
    if (!String(item.name || "").toLowerCase().includes(needle)) return false;
  }

  if (extensionFilter.size && isFile) {
    const extension = extname(item.name || "").toLowerCase();
    if (!extensionFilter.has(extension)) return false;
  }

  if (extensionFilter.size && isFolder && args.includeFolders === false) return false;
  return true;
}

async function resolveScanRoot(args = {}) {
  if (args.itemId) {
    const folder = await getRawInfo({ itemId: args.itemId, cacheResults: args.cacheResults });
    if (!folder.folder && !folder.root) throw new Error(`Scan target is not a folder: ${folder.name}`);
    return { id: folder.id, name: folder.name || "root", remotePath: itemRemotePath(folder) || "", target: `itemId:${folder.id}` };
  }

  const resolvedPath = resolvePresetPath(args);
  if (resolvedPath) {
    const folder = await getRawInfo({ path: resolvedPath, cacheResults: args.cacheResults });
    if (!folder.folder && !folder.root) throw new Error(`Scan target is not a folder: ${folder.name}`);
    return { id: folder.id, name: folder.name || resolvedPath, remotePath: resolvedPath, target: resolvedPath };
  }

  const root = await graph("/me/drive/root");
  return { id: root.id, name: "root", remotePath: "", target: "root" };
}

async function scan(args = {}) {
  return await withMetadataCacheBatch(async () => {
    const pageSize = clampInteger(args.pageSize, 200, 1, 200);
    const maxItems = clampInteger(args.maxItems, 10000, 1, 50000);
    const maxResults = clampInteger(args.maxResults, 500, 1, 5000);
    const maxDepth = clampInteger(args.maxDepth, 25, 0, 50);
    const maxFolders = clampInteger(args.maxFolders, 1000, 1, 10000);
    const scanConcurrency = clampInteger(args.scanConcurrency, Math.min(pluginSettings().concurrencyLimit, 4), 1, 4);
    const extensionFilter = normalizeExtensions(args.extensions || []);
    const skipFolderIds = new Set((args._skipFolderIds || []).filter(Boolean));
    const root = args._resolvedRoot || await resolveScanRoot(args);
    const params = new URLSearchParams();
    params.set("$top", String(pageSize));
    params.set("$select", args.select || defaultSelect);

    const queue = [{
      id: root.id,
      name: root.name,
      remotePath: root.remotePath,
      depth: 0,
      nextPath: null,
      counted: false,
      pagesFetched: 0,
      seenPages: new Set(),
      discoveredFolders: []
    }];
    const results = [];
    const pendingCacheItems = [];
    const counters = {
      itemsScanned: 0,
      filesScanned: 0,
      foldersScanned: 0,
      foldersVisited: 0,
      foldersSkipped: 0,
      matched: 0
    };
    let truncatedReason = null;

    const scanFolderPage = async (folder) => {
      const nextPath = folder.nextPath || `/me/drive/items/${encodeURIComponent(folder.id)}/children?${params.toString()}`;
      const maxPagesPerFolder = Math.max(1, maxItems + 100);
      if (folder.seenPages.has(nextPath)) {
        throw new Error(`Microsoft Graph pagination cycle detected while scanning ${folder.remotePath || folder.name || folder.id}.`);
      }
      if (folder.pagesFetched >= maxPagesPerFolder) {
        throw new Error(`Microsoft Graph pagination exceeded ${maxPagesPerFolder} pages while scanning ${folder.remotePath || folder.name || folder.id}.`);
      }
      return { folder, requestedPath: nextPath, page: await graph(nextPath) };
    };

    while (queue.length) {
      if (counters.itemsScanned >= maxItems) {
        truncatedReason = "maxItems";
        break;
      }
      if (counters.foldersVisited >= maxFolders && !queue[0]?.counted) {
        truncatedReason = "maxFolders";
        break;
      }

      const batch = [];
      let newFoldersInBatch = 0;
      while (queue.length && batch.length < scanConcurrency) {
        const candidate = queue[0];
        if (!candidate.counted && counters.foldersVisited + newFoldersInBatch >= maxFolders) break;
        queue.shift();
        if (!candidate.counted) newFoldersInBatch += 1;
        batch.push(candidate);
      }
      if (!batch.length) {
        truncatedReason = "maxFolders";
        break;
      }
      counters.foldersVisited += newFoldersInBatch;
      const scannedPages = await Promise.all(batch.map(scanFolderPage));
      const continuations = [];
      const discoveredFolders = [];

      for (const { folder, requestedPath, page } of scannedPages) {
        const cacheableItems = [];
        const pageDiscoveredFolders = [];
        for (const item of page.value || []) {
          if (counters.itemsScanned >= maxItems) {
            truncatedReason = "maxItems";
            break;
          }
          cacheableItems.push(item);
          counters.itemsScanned += 1;
          if (item.file) counters.filesScanned += 1;
          if (item.folder) counters.foldersScanned += 1;
          if (typeof args.onItem === "function") await args.onItem(item);

          if (scanItemMatches(item, args, extensionFilter)) {
            counters.matched += 1;
            if (results.length < maxResults) {
              results.push(formatDriveItem(item, args.format));
            }
            if (args.stopAfterResults === true && results.length >= maxResults) {
              truncatedReason = "maxResults";
              break;
            }
          }

          if (item.folder && folder.depth < maxDepth) {
            if (skipFolderIds.has(item.id)) {
              counters.foldersSkipped += 1;
            } else {
              pageDiscoveredFolders.push({
                id: item.id,
                name: item.name,
                remotePath: itemRemotePath(item),
                depth: folder.depth + 1,
                nextPath: null,
                counted: false,
                pagesFetched: 0,
                seenPages: new Set(),
                discoveredFolders: []
              });
            }
          }
        }
        if (args.cacheResults !== false) pendingCacheItems.push(...cacheableItems);
        if (truncatedReason) break;
        const nextPath = page["@odata.nextLink"] || null;
        const accumulatedDiscoveredFolders = [...(folder.discoveredFolders || []), ...pageDiscoveredFolders];
        if (nextPath) {
          continuations.push({
            ...folder,
            nextPath,
            counted: true,
            pagesFetched: folder.pagesFetched + 1,
            seenPages: new Set([...folder.seenPages, requestedPath]),
            discoveredFolders: accumulatedDiscoveredFolders
          });
        } else {
          discoveredFolders.push(...accumulatedDiscoveredFolders);
        }
      }

      if (truncatedReason) break;
      if (continuations.length) queue.unshift(...continuations);
      if (discoveredFolders.length) queue.push(...discoveredFolders);
    }

    if (!truncatedReason && queue.length) truncatedReason = "queueRemaining";
    if (args.cacheResults !== false && pendingCacheItems.length) {
      const deduplicatedCacheItems = [...new Map(pendingCacheItems.filter((item) => item?.id).map((item) => [item.id, item])).values()];
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(deduplicatedCacheItems));
    }
    return {
      root,
      filters: {
        nameContains: args.nameContains || null,
        extensions: [...extensionFilter],
        includeFiles: args.includeFiles !== false,
        includeFolders: args.includeFolders !== false,
        maxDepth,
        maxItems,
        maxFolders,
        maxResults,
        scanConcurrency
      },
      summary: {
        ...counters,
        returned: results.length,
        resultTruncated: counters.matched > results.length,
        traversalTruncated: Boolean(truncatedReason),
        truncatedReason,
        foldersQueued: new Set(queue.map((entry) => entry.id)).size
      },
      items: results,
      note: truncatedReason
        ? `Scan stopped at ${truncatedReason}. Increase the relevant cap or narrow the scan.`
        : "Recursive scan completed within the requested caps."
    };
  });
}

async function search(args = {}) {
  const query = String(args.query);
  const escaped = query.replace(/'/g, "''");
  const params = new URLSearchParams();
  params.set("$top", String(clampInteger(args.limit, 50, 1, 200)));
  params.set("$select", defaultSelect);
  let result = await graph(`/me/drive/root/search(q='${encodeURIComponent(escaped)}')?${params.toString()}`);
  let extensionFallback = null;
  const suffixMatch = query.trim().match(/^(.+?)\.([A-Za-z0-9]{1,10})$/);
  if (!(result.value || []).length && suffixMatch) {
    const fallbackQuery = suffixMatch[1].trim();
    const fallbackEscaped = fallbackQuery.replace(/'/g, "''");
    const fallbackResult = await graph(`/me/drive/root/search(q='${encodeURIComponent(fallbackEscaped)}')?${params.toString()}`);
    const expectedSuffix = `.${suffixMatch[2].toLowerCase()}`;
    result = {
      ...fallbackResult,
      value: (fallbackResult.value || []).filter((item) => String(item.name || "").toLowerCase().endsWith(expectedSuffix))
    };
    extensionFallback = fallbackQuery;
  }
  const prepared = await prepareSearchItems(result.value || [], { useCache: args.cacheResults !== false });
  if (args.cacheResults !== false) {
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(prepared.items));
  }
  return {
    items: prepared.items.map((item) => formatSimplifiedItem(item, args.format)),
    nextLink: result["@odata.nextLink"] || null,
    staleItemsFiltered: prepared.staleItemsFiltered,
    ...(extensionFallback ? { extensionFallback } : {})
  };
}

async function searchAll(args = {}) {
  const escaped = String(args.query).replace(/'/g, "''");
  const maxItems = clampInteger(args.maxItems, 1000, 1, 5000);
  const params = new URLSearchParams();
  params.set("$top", String(Math.min(clampInteger(args.pageSize, 200, 1, 200), maxItems)));
  params.set("$select", defaultSelect);
  return await collectPages(
    `/me/drive/root/search(q='${encodeURIComponent(escaped)}')?${params.toString()}`,
    maxItems,
    args.format,
    formatDriveItem,
    {
      cacheResults: args.cacheResults !== false,
      prepareItems: async (items) => (await prepareSearchItems(items, { useCache: args.cacheResults !== false })).items,
      formatter: formatSimplifiedItem
    }
  );
}

function preferCachedSearchItem(cached = {}, incoming = {}) {
  if (!cached?.id || cached.id !== incoming?.id) return false;
  const cachedTime = Date.parse(cached.lastModifiedDateTime || "");
  const incomingTime = Date.parse(incoming.lastModifiedDateTime || "");
  if (Number.isFinite(cachedTime) && Number.isFinite(incomingTime) && cachedTime > incomingTime) return true;
  return cached.eTag && incoming.eTag && cached.eTag !== incoming.eTag
    && (!Number.isFinite(cachedTime) || !Number.isFinite(incomingTime) || cachedTime >= incomingTime);
}

async function prepareSearchItems(items = [], options = {}) {
  if (options.useCache === false) {
    return { items: items.map((item) => simplifyItem(item)), staleItemsFiltered: 0 };
  }
  const cache = await loadMetadataCache();
  const tombstones = pruneSearchTombstones(cache);
  let staleItemsFiltered = 0;
  const prepared = [];
  for (const rawItem of items) {
    const incoming = simplifyItem(rawItem);
    if (isSearchTombstoned(incoming, tombstones)) {
      staleItemsFiltered += 1;
      continue;
    }
    const cached = cache.itemsById?.[incoming.id];
    prepared.push(preferCachedSearchItem(cached, incoming) ? cached : incoming);
  }
  return { items: prepared, staleItemsFiltered };
}

const findStopWords = new Set([
  "a", "an", "and", "by", "can", "could", "did", "do", "find", "for", "from", "gave", "get", "had", "has",
  "i", "in", "is", "it", "left", "locate", "me", "my", "named", "of", "on", "or", "please", "show", "that",
  "the", "they", "this", "to", "was", "were", "what", "where", "which", "who", "with"
]);
const findGenericWords = new Set([
  "called", "file", "files", "folder", "folders", "named", "document", "documents", "summary",
  "codex", "copy", "onedrive", "paper", "paperwork", "plugin", "something", "stuff", "test", "thing"
]);
const findKindHints = [
  {
    kind: "presentation",
    words: ["deck", "decks", "presentation", "presentations", "powerpoint", "ppt", "pptx", "slides", "slideshow"],
    extensions: [".pptx", ".ppt", ".pptm", ".ppsx", ".odp"]
  },
  {
    kind: "spreadsheet",
    words: ["spreadsheet", "spreadsheets", "sheet", "sheets", "excel", "xlsx", "xls", "csv", "workbook"],
    extensions: [".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods"]
  },
  {
    kind: "word",
    words: ["doc", "docs", "docx", "word", "letter", "memo"],
    extensions: [".docx", ".doc", ".docm", ".rtf", ".odt"]
  },
  {
    kind: "pdf",
    words: ["pdf"],
    extensions: [".pdf"]
  },
  {
    kind: "image",
    words: ["image", "images", "photo", "photos", "picture", "pictures", "screenshot", "screenshots", "png", "jpg", "jpeg"],
    extensions: [".png", ".jpg", ".jpeg", ".gif", ".webp", ".heic", ".tif", ".tiff"]
  }
];

// Keep semantic discovery deterministic and bounded. These aliases are only
// used to plan extra Graph/content-index searches; they never rewrite the
// user's query or claim that a result matched a concept without evidence from
// one of those searches.
const findConceptFamilies = [
  {
    id: "hvac",
    label: "HVAC",
    kind: "domain",
    triggers: [
      "hvac", "heating", "cooling", "air conditioning", "air conditioner", "ac repair",
      "furnace", "heat pump", "thermostat", "boiler", "ventilation"
    ],
    expansions: [
      "heating", "cooling", "air conditioning", "furnace", "heat pump", "thermostat", "boiler", "ventilation"
    ],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home", "House", "Home Maintenance"]
  },
  {
    id: "plumbing",
    label: "plumbing",
    kind: "domain",
    triggers: ["plumbing", "plumber", "leak", "drain", "water heater", "sewer", "faucet", "toilet"],
    expansions: ["plumbing", "leak repair", "drain service", "pipe repair", "water heater", "sewer"],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home", "House", "Home Maintenance"]
  },
  {
    id: "electrical",
    label: "electrical",
    kind: "domain",
    triggers: ["electrical", "electrician", "wiring", "breaker", "electrical panel", "outlet", "generator service"],
    expansions: ["electrical", "electrician", "wiring", "breaker panel", "outlet repair", "generator service"],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home", "House", "Home Maintenance"]
  },
  {
    id: "roofing",
    label: "roofing",
    kind: "domain",
    triggers: ["roof", "roofing", "roofer", "shingle", "gutter", "roof leak"],
    expansions: ["roofing", "roof repair", "shingles", "gutter service", "roof inspection"],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home", "House", "Home Maintenance"]
  },
  {
    id: "pest-control",
    label: "pest control",
    kind: "domain",
    triggers: ["pest", "exterminator", "termite", "rodent", "insects", "bed bugs"],
    expansions: ["pest control", "exterminator", "termite treatment", "rodent service", "pest inspection"],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home", "House", "Home Maintenance"]
  },
  {
    id: "appliance-repair",
    label: "appliance repair",
    kind: "domain",
    triggers: ["appliance", "refrigerator", "fridge", "washer", "dryer", "dishwasher", "oven"],
    expansions: ["appliance repair", "refrigerator service", "washer repair", "dryer repair", "dishwasher service"],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home", "House", "Home Maintenance"]
  },
  {
    id: "vehicle-service",
    label: "vehicle service",
    kind: "domain",
    triggers: ["car", "vehicle", "mechanic", "dealership", "oil change", "tire", "auto repair"],
    expansions: ["vehicle service", "auto repair", "mechanic", "oil change", "tire service", "dealership"],
    folderHints: ["Documents/Vehicles", "Personal/Documents/Vehicles", "Vehicles", "Auto"]
  },
  {
    id: "medical-record",
    label: "medical record",
    kind: "domain",
    triggers: ["doctor", "medical", "clinic", "hospital", "patient", "lab result", "prescription", "visit summary"],
    expansions: ["medical record", "visit summary", "after visit", "lab results", "patient report", "prescription"],
    folderHints: ["Personal/Documents/Health", "Documents/Health", "Health", "Medical"]
  },
  {
    id: "tax-record",
    label: "tax record",
    kind: "domain",
    triggers: ["tax", "taxes", "tax return", "w-2", "w2", "1099", "irs", "deduction"],
    expansions: ["tax return", "W-2", "1099", "IRS", "tax worksheet", "deductions"],
    folderHints: ["Documents/Taxes", "Personal/Documents/Taxes", "Taxes"]
  },
  {
    id: "insurance-record",
    label: "insurance record",
    kind: "domain",
    triggers: ["insurance", "insurance policy", "insurance claim", "claim number", "claim adjuster", "premium notice", "coverage letter", "deductible"],
    expansions: ["insurance policy", "claim", "coverage", "premium notice", "deductible", "declarations page"],
    folderHints: ["Documents/Insurance", "Personal/Documents/Insurance", "Insurance"]
  },
  {
    id: "travel-record",
    label: "travel record",
    kind: "domain",
    triggers: ["travel", "flight", "hotel", "booking", "reservation", "itinerary"],
    expansions: ["travel itinerary", "flight confirmation", "hotel reservation", "booking confirmation", "boarding pass"],
    folderHints: ["Documents/Travel", "Personal/Documents/Travel", "Travel"]
  },
  {
    id: "employment-record",
    label: "employment record",
    kind: "domain",
    triggers: ["resume", "résumé", "cv", "job offer", "offer letter", "pay stub", "paycheck", "employment"],
    expansions: ["resume", "curriculum vitae", "offer letter", "pay stub", "employment agreement"],
    folderHints: ["Documents/Career", "Personal/Documents/Career", "Career", "Employment"]
  },
  {
    id: "agreement",
    label: "agreement",
    kind: "document",
    triggers: ["contract", "agreement", "lease", "landlord", "tenant", "rental", "terms", "signed copy"],
    expansions: ["contract", "agreement", "signed document", "lease", "terms and conditions"],
    folderHints: ["Documents", "Personal/Documents", "Contracts", "Legal"]
  },
  {
    id: "service-record",
    label: "service record",
    kind: "document",
    triggers: [
      "work order", "service order", "service report", "maintenance report", "repair report",
      "job ticket", "invoice", "estimate", "inspection", "inspection report", "receipt", "technician", "contractor",
      "repair paperwork", "service paperwork", "proof of service"
    ],
    expansions: [
      "service order", "service report", "maintenance record", "repair invoice", "job ticket", "inspection report"
    ],
    folderHints: ["Documents/Home", "Personal/Documents/Home", "Home Maintenance", "Receipts"]
  }
];

function normalizeFindText(value = "") {
  return String(value)
    .toLowerCase()
    .replace(/[_]+/g, " ")
    .replace(/[^\p{L}\p{N}.-]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function findTokens(value = {}) {
  const normalized = normalizeFindText(value);
  const rawTokens = normalized.match(/[\p{L}\p{N}]+(?:[-.][\p{L}\p{N}]+)*/gu) || [];
  const expanded = [];
  for (const token of rawTokens) {
    expanded.push(token);
    if (/[-.]/.test(token)) expanded.push(...token.split(/[-.]/).filter(Boolean));
  }
  return [...new Set(expanded.filter((token) => token.length > 1 && !findStopWords.has(token)))];
}

function findDateTokens(query) {
  return [...new Set(String(query).match(/\b\d{4}[-_. ]\d{1,2}[-_. ]\d{1,2}\b/g) || [])]
    .map((token) => token.trim().replace(/[_. ]/g, "-"));
}

function findImportantTokens(query) {
  return findTokens(query).filter((token) => !findGenericWords.has(token));
}

function findConceptTriggerMatches(query, family) {
  const normalized = ` ${normalizeFindText(query)} `;
  return family.triggers.filter((trigger) => normalized.includes(` ${normalizeFindText(trigger)} `));
}

function findQueryConcepts(query) {
  return findConceptFamilies.flatMap((family) => {
    const triggers = findConceptTriggerMatches(query, family);
    return triggers.length ? [{ ...family, matchedTriggers: triggers }] : [];
  });
}

function buildFindSearchPlan(query, maxSearchTerms = 8) {
  const plan = [];
  const add = (term, metadata = {}) => {
    const clean = String(term || "").replace(/\s+/g, " ").trim();
    if (clean && !plan.some((existing) => existing.term.toLowerCase() === clean.toLowerCase())) {
      plan.push({ term: clean, ...metadata });
    }
  };
  const dateTokens = findDateTokens(query);
  const dateParts = new Set(dateTokens.flatMap((token) => [token, ...token.split("-")]));
  const important = findImportantTokens(query).filter((token) => !dateParts.has(token));
  const concepts = findQueryConcepts(query);
  const conceptVocabulary = new Set(concepts.flatMap((concept) => [
    ...concept.triggers.flatMap((value) => findTokens(value)),
    ...concept.expansions.flatMap((value) => findTokens(value))
  ]));
  const intentActionWords = new Set([
    "after", "checked", "checking", "fixed", "fixing", "inspection", "maintenance", "order", "record",
    "repair", "report", "service", "someone", "technician", "work", "worked"
  ]);
  const specificTokens = important.filter((token) => !conceptVocabulary.has(token) && !intentActionWords.has(token));

  add(query, { kind: "canonical" });
  for (const dateToken of dateTokens) add(dateToken, { kind: "literal" });
  if (specificTokens.length) add(specificTokens.join(" "), { kind: "specific-literal" });
  for (const concept of concepts) {
    for (const expansion of concept.expansions) {
      if (concept.matchedTriggers.some((trigger) => normalizeFindText(trigger) === normalizeFindText(expansion))) continue;
      add(expansion, {
        kind: "semantic",
        semanticExpansion: true,
        semanticConceptId: concept.id,
        semanticConceptLabel: concept.label,
        semanticKind: concept.kind
      });
    }
  }
  if (important.length) add(important.join(" "), { kind: "literal" });
  if (important.length >= 3) add(important.slice(-3).join(" "), { kind: "literal" });
  if (important.length >= 2) {
    for (let index = 0; index < important.length - 1; index += 1) {
      add(`${important[index]} ${important[index + 1]}`, { kind: "literal" });
    }
  }
  for (const token of [...important].sort((a, b) => b.length - a.length)) {
    if (token.length >= 4) add(token, { kind: "literal" });
  }

  return plan.slice(0, Math.min(maxSearchTerms, 12));
}

function buildFindSearchTerms(query, maxSearchTerms = 8) {
  return buildFindSearchPlan(query, maxSearchTerms).map((entry) => entry.term);
}

function inferFindExtensions(query, explicitExtensions = []) {
  const explicit = normalizeExtensions(explicitExtensions);
  const inferred = new Set();
  const strictInferred = new Set();
  const tokens = new Set(findTokens(query));
  const normalized = normalizeFindText(query);
  const matchedKinds = [];
  for (const hint of findKindHints) {
    if (hint.words.some((word) => tokens.has(word) || normalized.includes(word))) {
      matchedKinds.push(hint.kind);
      for (const extension of hint.extensions) inferred.add(extension);
      const strictMatch = hint.words.some((word) => tokens.has(word) && (word.startsWith(".") || hint.extensions.includes(`.${word}`) || ["pdf", "ppt", "pptx", "powerpoint", "excel", "xlsx", "xls", "csv", "doc", "docx", "word", "png", "jpg", "jpeg"].includes(word)));
      if (strictMatch) {
        for (const extension of hint.extensions) strictInferred.add(extension);
      }
    }
  }
  return {
    explicit,
    inferred,
    strictInferred,
    effectiveForScan: explicit.size ? explicit : strictInferred,
    matchedKinds
  };
}

function findItemExtension(item = {}) {
  return extname(item.name || "").toLowerCase();
}

function findItemType(item = {}) {
  return item.folder ? "folder" : item.file ? "file" : "item";
}

function findCandidateKey(item = {}) {
  return item.id || item.remotePath || `${item.name || "unnamed"}:${item.size || 0}:${item.lastModifiedDateTime || ""}`;
}

function shouldIncludeFindItem(item, args = {}, extensionInfo = {}) {
  const type = findItemType(item);
  if (type === "folder" && args.includeFolders !== true) return false;
  if (extensionInfo.explicit?.size && type === "file" && !extensionInfo.explicit.has(findItemExtension(item))) return false;
  return true;
}

function scoreFindCandidate(item, context = {}) {
  const queryText = normalizeFindText(context.query);
  const queryTokens = findImportantTokens(context.query);
  const dateTokens = findDateTokens(context.query);
  const nameText = normalizeFindText(item.name || "");
  const pathText = normalizeFindText(item.remotePath || item.path || "");
  const nameTokenSet = new Set(findTokens(item.name || ""));
  const pathTokenSet = new Set(findTokens(item.remotePath || item.path || ""));
  const extension = findItemExtension(item);
  const reasons = [];
  let score = 0;
  let matchedTokens = 0;

  if (context.source === "exactPath") {
    score += 100;
    reasons.push("exact path");
  } else if (context.source === "search") {
    score += 18;
    reasons.push(context.canonicalSearch === true
      ? "canonical Graph content/metadata match"
      : `Graph search: ${context.term}`);
  } else if (context.source === "scan") {
    score += 12;
    reasons.push(`scan: ${context.folder || "root"}`);
  } else if (context.source === "cache") {
    score += 10;
    reasons.push("metadata cache");
  } else if (context.source === "metadataConfirm") {
    score += 12;
    reasons.push("metadata cache confirmed live");
  } else if (context.source === "contentIndex") {
    score += Math.min(55, Math.max(20, context.contentScore || 20));
    matchedTokens += context.contentMatchedTokens || 0;
    reasons.push("indexed content");
    if (context.contentReason) reasons.push(context.contentReason);
  }

  if (context.semanticExpansion === true) {
    score += context.semanticKind === "domain" ? 26 : 16;
    reasons.push(`concept match: ${context.semanticConceptLabel || context.semanticConceptId || context.term}`);
    const semanticText = normalizeFindText(context.term);
    const semanticTokens = findImportantTokens(context.term);
    const semanticMetadataMatch = (semanticText && (nameText.includes(semanticText) || pathText.includes(semanticText)))
      || semanticTokens.some((token) => nameTokenSet.has(token) || pathTokenSet.has(token));
    if (semanticMetadataMatch) {
      score += 16;
      reasons.push(`concept visible in metadata: ${context.term}`);
    }
  }

  if (queryText && nameText === queryText) {
    score += 90;
    reasons.push("exact filename");
  } else if (queryText && nameText.includes(queryText)) {
    score += 65;
    reasons.push("filename contains full query");
  } else if (queryText && pathText.includes(queryText)) {
    score += 35;
    reasons.push("path contains full query");
  }

  for (const token of queryTokens) {
    if (nameTokenSet.has(token) || nameText.includes(token)) {
      matchedTokens += 1;
      score += token.length >= 5 ? 9 : 5;
    } else if (pathTokenSet.has(token) || pathText.includes(token)) {
      matchedTokens += 1;
      score += token.length >= 5 ? 4 : 2;
    }
  }

  if (queryTokens.length) {
    const coverage = matchedTokens / queryTokens.length;
    score += Math.round(coverage * 45);
    if (coverage >= 0.85) reasons.push("strong token match");
    else if (coverage >= 0.5) reasons.push("partial token match");
  }

  for (const dateToken of dateTokens) {
    const compact = dateToken.replace(/-/g, "");
    if (nameText.includes(dateToken) || pathText.includes(dateToken) || nameText.replace(/[-_. ]/g, "").includes(compact)) {
      score += 18;
      reasons.push(`date match: ${dateToken}`);
    }
  }

  if (context.extensionInfo?.explicit?.size && context.extensionInfo.explicit.has(extension)) {
    score += 35;
    reasons.push(`requested extension: ${extension}`);
  } else if (context.extensionInfo?.inferred?.size && context.extensionInfo.inferred.has(extension)) {
    score += 24;
    reasons.push(`likely file type: ${extension}`);
  }

  if (findItemType(item) === "folder") score -= 15;
  if (item.remotePath && context.scoringFolderHints?.some((hint) => {
    const cleanHint = cleanPath(hint).toLowerCase();
    const cleanItemPath = cleanPath(item.remotePath).toLowerCase();
    return cleanHint && (cleanItemPath === cleanHint || cleanItemPath.startsWith(`${cleanHint}/`));
  })) {
    score += 8;
    reasons.push("folder hint");
  }

  return {
    score: Math.max(0, Math.round(score)),
    matchedTokens,
    reasons: [...new Set(reasons)].slice(0, 6)
  };
}

function defaultFindFolderHints(query = "") {
  const normalized = normalizeFindText(query);
  const hints = [
    "Personal/Documents",
    "Documents",
    "Personal",
    "Microsoft Copilot Chat Files",
    "Desktop"
  ];
  if (normalized.includes("screenshot") || normalized.includes("photo") || normalized.includes("picture")) {
    hints.unshift("Pictures/Screenshots", "Pictures");
  }
  if (normalized.includes("health") || normalized.includes("eval") || normalized.includes("evaluation")) {
    hints.unshift("Personal/Documents/Health");
  }
  for (const concept of findQueryConcepts(query)) hints.unshift(...(concept.folderHints || []));
  return [...new Set(hints)];
}

function normalizeFolderHintKey(hint = "") {
  return cleanPath(hint).toLowerCase();
}

function pruneFolderHints(hints = []) {
  const pruned = [];
  const seen = new Set();
  for (const hint of hints) {
    const clean = cleanPath(hint || "");
    const key = normalizeFolderHintKey(clean);
    if (seen.has(key)) continue;
    seen.add(key);
    if (!key) {
      pruned.push(clean);
      continue;
    }
    const overlapsEarlierHint = pruned.some((earlier) => {
      const earlierKey = normalizeFolderHintKey(earlier);
      if (!earlierKey) return false;
      return key.startsWith(`${earlierKey}/`) || earlierKey.startsWith(`${key}/`);
    });
    if (!overlapsEarlierHint) pruned.push(clean);
  }
  return pruned;
}

function findScanNeedle(query = "") {
  const tokens = findImportantTokens(query)
    .filter((token) => !["deck", "presentation", "powerpoint", "spreadsheet", "sheet", "pdf", "doc", "word"].includes(token))
    .sort((a, b) => b.length - a.length);
  return tokens.find((token) => token.length >= 4) || findDateTokens(query)[0] || tokens[0] || "";
}

function scanNeedleIsWeak(needle = "", query = "") {
  if (!needle) return true;
  const important = findImportantTokens(query);
  if (!important.includes(needle)) return true;
  if (needle.length <= 3) return true;
  return false;
}

function addFindCandidate(candidates, item, context) {
  if (!shouldIncludeFindItem(item, context.args, context.extensionInfo)) return;
  const key = findCandidateKey(item);
  const scored = scoreFindCandidate(item, context);
  const queryTokens = findImportantTokens(context.query);
  const contentCoverage = queryTokens.length && context.contentMatchedTokens
    ? context.contentMatchedTokens / queryTokens.length
    : 0;
  const strongContentIndexMatch = context.source === "contentIndex"
    && (context.contentExactPhrase === true || (context.contentMatchedTokens || 0) >= 2 || contentCoverage >= 0.5);
  const canonicalGraphMatch = context.source === "search" && context.canonicalSearch === true;
  const semanticSearchMatch = context.semanticExpansion === true
    && ["search", "contentIndex"].includes(context.source);
  const hasQueryRelevance = scored.matchedTokens > 0 || scored.reasons.some((reason) =>
    reason.startsWith("exact filename")
    || reason.startsWith("filename contains")
    || reason.startsWith("path contains")
    || reason.startsWith("date match")
    || reason.startsWith("requested extension")
    || reason.startsWith("likely file type")
  ) || strongContentIndexMatch || canonicalGraphMatch || semanticSearchMatch;
  if (context.source !== "exactPath" && !hasQueryRelevance) return;
  const existing = candidates.get(key);
  const semanticEvidence = context.semanticExpansion === true
    ? [`${context.semanticConceptId || "concept"}:${normalizeFindText(context.term)}`]
    : [];
  if (!existing || scored.score > existing.score) {
    const previousEvidence = existing?.semanticEvidence || [];
    const mergedEvidence = [...new Set([...previousEvidence, ...semanticEvidence])];
    const baseScore = Math.max(existing?.baseScore || 0, scored.score);
    candidates.set(key, {
      item,
      baseScore,
      score: baseScore + Math.min(40, Math.max(0, mergedEvidence.length - 1) * 16),
      reasons: [...new Set([...(existing?.reasons || []), ...scored.reasons])].slice(0, 6),
      snippets: [...new Set([...(existing?.snippets || []), ...(context.snippet ? [context.snippet] : [])])],
      semanticEvidence: mergedEvidence,
      sources: [...(existing?.sources || []), { source: context.source, term: context.term, folder: context.folder }]
    });
  } else {
    existing.sources.push({ source: context.source, term: context.term, folder: context.folder });
    existing.reasons = [...new Set([...existing.reasons, ...scored.reasons])].slice(0, 6);
    if (context.snippet && !existing.snippets.includes(context.snippet)) existing.snippets.push(context.snippet);
    existing.semanticEvidence = [...new Set([...(existing.semanticEvidence || []), ...semanticEvidence])];
    existing.baseScore = Math.max(existing.baseScore || existing.score, scored.score);
    existing.score = existing.baseScore + Math.min(40, Math.max(0, existing.semanticEvidence.length - 1) * 16);
  }
}

function rankedFindCandidates(candidates) {
  return [...candidates.values()].sort((left, right) => {
    if (right.score !== left.score) return right.score - left.score;
    return String(left.item.name || "").localeCompare(String(right.item.name || ""));
  });
}

function candidateMatchesConceptMetadata(candidate, concept) {
  const metadataText = normalizeFindText(`${candidate.item?.name || ""} ${candidate.item?.remotePath || candidate.item?.path || ""}`);
  const metadataTokens = new Set(findTokens(metadataText));
  return [...concept.triggers, ...concept.expansions].some((value) => {
    const normalized = normalizeFindText(value);
    if (!normalized) return false;
    if (/\s|[-.]/.test(normalized)) return ` ${metadataText} `.includes(` ${normalized} `);
    return metadataTokens.has(normalized);
  });
}

function candidateHasConceptEvidence(candidate, queryConcepts = []) {
  if (!queryConcepts.length) return true;
  if (candidate.sources.some((source) => source.source === "exactPath" || source.source === "contentIndex")) return true;
  const requiredConcepts = queryConcepts.some((concept) => concept.kind === "domain")
    ? queryConcepts.filter((concept) => concept.kind === "domain")
    : queryConcepts;
  if (requiredConcepts.some((concept) => candidateMatchesConceptMetadata(candidate, concept))) return true;
  const requiredIds = new Set(requiredConcepts.map((concept) => concept.id));
  const corroboratingEvidence = (candidate.semanticEvidence || [])
    .filter((evidence) => requiredIds.has(String(evidence).split(":", 1)[0]));
  return new Set(corroboratingEvidence).size >= 2;
}

function conceptRelevantFindCandidates(candidates, queryConcepts = []) {
  return rankedFindCandidates(candidates).filter((candidate) => candidateHasConceptEvidence(candidate, queryConcepts));
}

function candidateHasLiveSource(candidate) {
  return candidate.sources.some((source) => ["exactPath", "search", "scan", "metadataConfirm"].includes(source.source));
}

function cacheConfirmationEligible(candidate, query = "") {
  const normalizedQuery = normalizeFindText(query);
  const nameText = normalizeFindText(candidate.item?.name || "");
  const pathText = normalizeFindText(candidate.item?.remotePath || candidate.item?.path || "");
  const queryLooksPathLike = String(query).includes("/");
  if (normalizedQuery && nameText === normalizedQuery) return true;
  if (queryLooksPathLike && normalizedQuery && pathText.includes(normalizedQuery)) return true;
  const extension = findItemExtension(candidate.item);
  return Boolean(extension && normalizedQuery.endsWith(extension) && nameText.includes(normalizedQuery));
}

async function confirmCachedFindCandidates(candidates, ranked, context = {}) {
  const maxConfirmations = clampInteger(context.args?.cacheConfirmMaxItems, 5, 1, 20);
  const minScore = clampInteger(context.args?.cacheConfirmMinScore, 60, 0, 200);
  const targets = ranked
    .filter((candidate) => !candidateHasLiveSource(candidate))
    .filter((candidate) => candidate.item?.id && candidate.score >= minScore)
    .filter((candidate) => candidate.sources.some((source) => source.source === "cache"))
    .filter((candidate) => cacheConfirmationEligible(candidate, context.query))
    .slice(0, maxConfirmations);
  if (!targets.length) return { attempted: 0, confirmed: 0, errors: 0 };
  let confirmed = 0;
  let errors = 0;
  let confirmedItems = [];
  try {
    const result = await batchGetInfo({
      items: targets.map((candidate) => ({ itemId: candidate.item.id })),
      format: "full"
    });
    for (const item of result.items || []) {
      if (item.error) {
        errors += 1;
        continue;
      }
      confirmedItems.push(item);
    }
  } catch {
    for (const candidate of targets) {
      try {
        confirmedItems.push(simplifyItem(await getRawInfo({ itemId: candidate.item.id })));
      } catch {
        errors += 1;
      }
    }
  }
  for (const item of confirmedItems) {
    if (item.error) {
      errors += 1;
      continue;
    }
    confirmed += 1;
    addFindCandidate(candidates, item, {
      args: context.args,
      query: context.query,
      source: "metadataConfirm",
      term: "batch-get-info",
      extensionInfo: context.extensionInfo,
      scoringFolderHints: context.scoringFolderHints
    });
  }
  return { attempted: targets.length, confirmed, errors };
}

function scanRootKey(root = {}, fallbackFolder = "") {
  return root.id || normalizeFolderHintKey(root.remotePath || fallbackFolder);
}

async function find(args = {}) {
  const startedAt = Date.now();
  const query = String(args.query || "").trim();
  if (!query) throw new Error("query is required.");

  const maxResults = Math.min(args.maxResults ?? 10, args.maxResultsLimit ?? 50);
  const maxSearchTerms = Math.min(args.maxSearchTerms ?? 8, 12);
  const searchPlanEntries = buildFindSearchPlan(query, maxSearchTerms);
  const searchTerms = searchPlanEntries.map((entry) => entry.term);
  const queryConcepts = findQueryConcepts(query);
  const extensionInfo = inferFindExtensions(query, args.extensions || []);
  const scoringFolderHints = pruneFolderHints(args.folderHints || []);
  const explicitFolderHintKeys = new Set(scoringFolderHints.map(normalizeFolderHintKey));
  const folderHints = pruneFolderHints([...scoringFolderHints, ...defaultFindFolderHints(query), ""]);
  const candidates = new Map();
  const searchRuns = [];
  const scanRuns = [];
  const pendingSearchCacheItems = [];
  const searchConcurrency = clampInteger(args.searchConcurrency, 2, 1, 4);
  const searchConfidenceThreshold = args.minConfidenceForSearchOnly ?? 78;
  const executedSearchTerms = [];
  const skippedSearchTerms = [];
  let searchStopReason = null;
  let cacheCandidateCount = 0;
  let contentIndexCandidateCount = 0;
  let contentIndexDurationMs = 0;
  let liveSearchDurationMs = 0;
  let scanDurationMs = 0;
  let cacheConfirmDurationMs = 0;
  let cacheFresh = false;
  let cacheAgeSeconds = null;
  let cacheWithinStaleWindow = false;
  let usedFreshLocalFastPath = false;
  let usedStaleLocalFastPath = false;
  let cacheConfirmations = { attempted: 0, confirmed: 0, errors: 0 };

  if (args.useCache !== false) {
    const cache = await loadMetadataCache();
    const cacheList = Object.values(cache.itemsById || {});
    cacheFresh = metadataCacheFresh(cache);
    cacheAgeSeconds = metadataCacheAgeSeconds(cache);
    cacheWithinStaleWindow = !cacheFresh
      && args.preferStaleLocalResults === true
      && metadataCacheWithinStaleWindow(cache, clampInteger(args.staleLocalMaxAgeSeconds, chatgptStaleCacheMaxAgeSeconds, 1, 30 * 24 * 60 * 60));
    for (const item of cacheList) {
      addFindCandidate(candidates, item, {
        args,
        query,
        source: "cache",
        term: "metadata-cache",
        extensionInfo,
        scoringFolderHints
      });
    }
    cacheCandidateCount = cacheList.length;
  }

  if (args.useContentIndex !== false && (args.contentMaxResults ?? 10) > 0) {
    const contentStartedAt = Date.now();
    const contentQueries = searchPlanEntries
      .filter((entry) => entry.kind === "canonical" || entry.semanticExpansion === true)
      .slice(0, clampInteger(args.contentMaxQueries, 4, 1, 8));
    const indexedRuns = await Promise.all(contentQueries.map(async (entry) => ({
      entry,
      result: await contentSearch({
        query: entry.term,
        maxResults: clampInteger(args.contentMaxResults, 10, 0, 100) || 1,
        format: "full"
      })
    })));
    contentIndexDurationMs += elapsedMs(contentStartedAt);
    contentIndexCandidateCount = indexedRuns.reduce((sum, run) => sum + (run.result.items?.length || 0), 0);
    for (const { entry, result: indexed } of indexedRuns) {
      for (const match of indexed.items || []) {
        addFindCandidate(candidates, match, {
          args,
          query,
          source: "contentIndex",
          term: entry.term,
          semanticExpansion: entry.semanticExpansion,
          semanticConceptId: entry.semanticConceptId,
          semanticConceptLabel: entry.semanticConceptLabel,
          semanticKind: entry.semanticKind,
          extensionInfo,
          scoringFolderHints,
          contentScore: match.score,
          contentMatchedTokens: match.matchedTokens || 1,
          contentExactPhrase: match.reasons?.includes("indexed content phrase"),
          contentReason: match.reasons?.[0],
          snippet: match.snippet
        });
      }
    }
  }

  if (query.includes("/")) {
    try {
      const item = simplifyItem(await getRawInfo({ path: query, cacheResults: args.useCache !== false }));
      addFindCandidate(candidates, item, {
        args,
        query,
        source: "exactPath",
        extensionInfo,
        scoringFolderHints
      });
    } catch (error) {
      searchRuns.push({ strategy: "exactPath", path: query, error: safeToolErrorMessage(error) });
    }
  }

  if ((args.preferFreshLocalResults === true && cacheFresh) || cacheWithinStaleWindow) {
    const bestLocal = conceptRelevantFindCandidates(candidates, queryConcepts)[0];
    const localConfidenceThreshold = cacheFresh
      ? clampInteger(args.freshLocalMinConfidence, 60, 0, 200)
      : clampInteger(args.staleLocalMinConfidence, 75, 0, 200);
    if (bestLocal && bestLocal.score >= localConfidenceThreshold) {
      if (cacheFresh) {
        usedFreshLocalFastPath = true;
        searchStopReason = "fresh-local-cache";
      } else {
        usedStaleLocalFastPath = true;
        searchStopReason = "stale-local-cache";
      }
    }
  }

  const executeSearchTerm = async ({ term, index, stage, ...planMetadata }) => {
    const termStartedAt = Date.now();
    executedSearchTerms.push(term);
    try {
      const result = await searchAll({
        query: term,
        pageSize: args.searchPageSize ?? 50,
        maxItems: args.searchMaxItemsPerTerm ?? 100,
        format: "full",
        cacheResults: false
      });
      if (args.useCache !== false) pendingSearchCacheItems.push(...(result.items || []));
      for (const [resultIndex, item] of (result.items || []).entries()) {
        addFindCandidate(candidates, item, {
          args,
          query,
          source: "search",
          term,
          canonicalSearch: index === 0,
          ...planMetadata,
          searchResultRank: resultIndex + 1,
          extensionInfo,
          scoringFolderHints
        });
      }
      return {
        term,
        stage,
        executed: true,
        count: result.count,
        truncated: result.truncated,
        unsafePageTruncation: result.unsafePageTruncation,
        graphSearchCalls: result.pagesFetched || 0,
        durationMs: elapsedMs(termStartedAt)
      };
    } catch (error) {
      return {
        term,
        stage,
        executed: true,
        graphSearchCalls: error.graphPagesFetched || 0,
        durationMs: elapsedMs(termStartedAt),
        error: safeToolErrorMessage(error)
      };
    }
  };

  const executeSearchWave = async (entries) => {
    if (!entries.length) return;
    const waveStartedAt = Date.now();
    const runs = await mapWithConcurrency(entries, searchConcurrency, executeSearchTerm);
    liveSearchDurationMs += elapsedMs(waveStartedAt);
    searchRuns.push(...runs);
  };

  if (searchTerms.length && !usedFreshLocalFastPath && !usedStaleLocalFastPath) {
    const initialSearchTermCount = clampInteger(args.initialSearchTermCount, 1, 1, searchConcurrency);
    const initialWave = searchPlanEntries
      .slice(0, initialSearchTermCount)
      .map((entry, index) => ({ ...entry, index, stage: index === 0 ? "canonical" : "initial-expansion" }));
    await executeSearchWave(initialWave);
    let nextSearchTermIndex = initialWave.length;
    let bestLiveSearchScore = conceptRelevantFindCandidates(candidates, queryConcepts).find(candidateHasLiveSource)?.score || 0;
    if (args.executeAllSearchTerms !== true && bestLiveSearchScore >= searchConfidenceThreshold && nextSearchTermIndex < searchTerms.length) {
      searchStopReason = "high-confidence-canonical";
    } else {
      while (nextSearchTermIndex < searchPlanEntries.length) {
        const wave = searchPlanEntries
          .slice(nextSearchTermIndex, nextSearchTermIndex + searchConcurrency)
          .map((entry, offset) => ({ ...entry, index: nextSearchTermIndex + offset, stage: "expansion" }));
        nextSearchTermIndex += wave.length;
        await executeSearchWave(wave);
        bestLiveSearchScore = conceptRelevantFindCandidates(candidates, queryConcepts).find(candidateHasLiveSource)?.score || 0;
        if (args.executeAllSearchTerms !== true && bestLiveSearchScore >= searchConfidenceThreshold && nextSearchTermIndex < searchTerms.length) {
          searchStopReason = "high-confidence-expansion";
          break;
        }
      }
    }

  }
  for (const [index, term] of searchTerms.slice(executedSearchTerms.length).entries()) {
      skippedSearchTerms.push(term);
      searchRuns.push({
        term,
        stage: executedSearchTerms.length + index === 0 ? "canonical" : "expansion",
        executed: false,
        skipped: searchStopReason || "adaptive-stop"
      });
  }
  if (!searchStopReason) searchStopReason = "all-terms-executed";

  if (args.useCache !== false && pendingSearchCacheItems.length) {
    const uniqueSearchCacheItems = [...new Map(pendingSearchCacheItems.filter((item) => item?.id).map((item) => [item.id, item])).values()];
    await bestEffortLocalWrite("metadata cache update", async () => await withMetadataCacheBatch(async () => {
      await cacheItems(uniqueSearchCacheItems);
    }));
  }

  let ranked = conceptRelevantFindCandidates(candidates, queryConcepts);
  if (args.confirmCacheCandidates !== false) {
    const confirmStartedAt = Date.now();
    cacheConfirmations = await confirmCachedFindCandidates(candidates, ranked, {
      args,
      query,
      extensionInfo,
      scoringFolderHints
    });
    cacheConfirmDurationMs += elapsedMs(confirmStartedAt);
    ranked = conceptRelevantFindCandidates(candidates, queryConcepts);
  }
  const bestLiveScore = ranked.find(candidateHasLiveSource)?.score || 0;
  const shouldScan = args.scanFallback !== false && bestLiveScore < searchConfidenceThreshold;

  if (shouldScan) {
    const scanNeedle = findScanNeedle(query);
    const scanConcurrency = clampInteger(args.scanConcurrency, 2, 1, 4);
    const scanPlans = [
      { nameContains: scanNeedle, extensions: [...extensionInfo.effectiveForScan], reason: "targeted" }
    ];
    if (extensionInfo.effectiveForScan.size || scanNeedleIsWeak(scanNeedle, query)) {
      scanPlans.push({ nameContains: scanNeedle, extensions: [], reason: "no-extension-filter" });
    }
    if (scanNeedle && scanNeedleIsWeak(scanNeedle, query) && extensionInfo.inferred.size) {
      scanPlans.push({ nameContains: "", extensions: [...extensionInfo.strictInferred], reason: "type-only" });
    }

    for (const plan of scanPlans) {
      let remainingItems = Math.min(args.scanMaxItems ?? 1500, 10000);
      let remainingFolders = Math.min(args.scanMaxFolders ?? 250, 2000);
      const scannedFolderKeys = new Set();
      const completedScanRootIds = new Set();
      let folderIndex = 0;
      while (folderIndex < folderHints.length && remainingItems > 0 && remainingFolders > 0) {
        const targetBatch = [];
        while (folderIndex < folderHints.length && targetBatch.length < scanConcurrency) {
          const folder = folderHints[folderIndex];
          folderIndex += 1;
          // Keep the broad root fallback in its own wave so it can exclude every
          // hinted subtree that earlier waves already traversed completely.
          if (!folder && targetBatch.length > 0) {
            folderIndex -= 1;
            break;
          }
          try {
            const root = await resolveScanRoot({
              ...(folder ? { path: folder } : {}),
              useCache: args.useCache !== false,
              cacheResults: args.useCache !== false
            });
            const key = scanRootKey(root, folder);
            if (scannedFolderKeys.has(key)) {
              scanRuns.push({ folder: folder || "root", reason: plan.reason, skipped: "duplicate-root", root: { id: root.id, remotePath: root.remotePath } });
              continue;
            }
            scannedFolderKeys.add(key);
            targetBatch.push({ folder, root });
          } catch (error) {
            const errorMessage = safeToolErrorMessage(error);
            if (folder && !explicitFolderHintKeys.has(normalizeFolderHintKey(folder)) && /\b(?:itemNotFound|notFound)\b/i.test(errorMessage)) {
              scanRuns.push({ folder, reason: plan.reason, skipped: "missing-optional-folder" });
            } else {
              scanRuns.push({ folder: folder || "root", reason: plan.reason, error: errorMessage });
            }
          }
        }
        if (!targetBatch.length) continue;
        const perScanMaxItems = Math.max(1, Math.ceil(remainingItems / targetBatch.length));
        const perScanMaxFolders = Math.max(1, Math.ceil(remainingFolders / targetBatch.length));
        const scanStartedAt = Date.now();
        const batchResults = await mapWithConcurrency(targetBatch, scanConcurrency, async (target) => {
          try {
            const result = await scan({
              path: target.folder,
              _resolvedRoot: target.root,
              nameContains: plan.nameContains,
              extensions: plan.extensions,
              includeFiles: true,
              includeFolders: args.includeFolders === true,
              maxItems: perScanMaxItems,
              maxFolders: perScanMaxFolders,
              maxDepth: args.scanMaxDepth ?? 20,
              maxResults: Math.max(maxResults * 2, 20),
              stopAfterResults: true,
              format: "full",
              cacheResults: args.useCache !== false,
              _skipFolderIds: [...completedScanRootIds].filter((id) => id !== target.root.id)
            });
            return { target, result };
          } catch (error) {
            return { target, error };
          }
        });
        scanDurationMs += elapsedMs(scanStartedAt);
        for (const entry of batchResults) {
          const folderLabel = entry.target.folder || "root";
          if (entry.error) {
            scanRuns.push({ folder: folderLabel, reason: plan.reason, error: safeToolErrorMessage(entry.error) });
            continue;
          }
          const result = entry.result;
          scanRuns.push({ folder: folderLabel, reason: plan.reason, nameContains: plan.nameContains || null, extensions: plan.extensions, summary: result.summary });
          if (!result.summary.traversalTruncated && entry.target.root.id) {
            completedScanRootIds.add(entry.target.root.id);
          }
          remainingItems -= result.summary.itemsScanned || 0;
          remainingFolders -= result.summary.foldersVisited || 0;
          for (const item of result.items || []) {
            addFindCandidate(candidates, item, {
              args,
              query,
              source: "scan",
              folder: folderLabel,
              extensionInfo,
              scoringFolderHints
            });
          }
        }
        ranked = conceptRelevantFindCandidates(candidates, queryConcepts);
        if ((ranked[0]?.score || 0) >= searchConfidenceThreshold && ranked.length >= maxResults) break;
      }
      ranked = conceptRelevantFindCandidates(candidates, queryConcepts);
      if (ranked.length > 0) break;
    }
  }

  ranked = conceptRelevantFindCandidates(candidates, queryConcepts);
  if (shouldScan) {
    ranked = ranked.filter(candidateHasLiveSource);
  }
  const noteParts = [];
  if (cacheCandidateCount > 0) noteParts.push("Used the local metadata cache as a candidate source.");
  if (contentIndexCandidateCount > 0) noteParts.push("Used matching entries from the local content index.");
  if (executedSearchTerms.length > 0) {
    noteParts.push(`Ran ${executedSearchTerms.length} live Graph search ${executedSearchTerms.length === 1 ? "term" : "terms"}.`);
  }
  if (skippedSearchTerms.length > 0) {
    noteParts.push(`Skipped ${skippedSearchTerms.length} expansion ${skippedSearchTerms.length === 1 ? "term" : "terms"} after reaching the confidence threshold.`);
  }
  if (shouldScan) {
    noteParts.push("Used bounded remote recursive scan fallback because live-search confidence remained below the threshold.");
  } else if (args.scanFallback === false) {
    noteParts.push("Recursive scan fallback was disabled.");
  } else {
    noteParts.push("Recursive scan fallback was not required by the configured confidence threshold.");
  }
  return {
    query,
    strategy: "cache-assisted-remote-first",
    searchTerms,
    searchPlan: {
      mode: "adaptive-staged",
      concurrency: searchConcurrency,
      confidenceThreshold: searchConfidenceThreshold,
      planned: searchTerms,
      executed: executedSearchTerms,
      skipped: skippedSearchTerms,
      stopReason: searchStopReason
    },
    inferred: {
      extensions: [...extensionInfo.inferred],
      strictExtensions: [...extensionInfo.strictInferred],
      matchedKinds: extensionInfo.matchedKinds,
      explicitExtensions: [...extensionInfo.explicit]
    },
    summary: {
      candidates: ranked.length,
      returned: Math.min(ranked.length, maxResults),
      bestScore: ranked[0]?.score || 0,
      usedScanFallback: shouldScan,
      scanRuns: scanRuns.length,
      localIndexUsed: contentIndexCandidateCount > 0,
      contentIndexCandidates: contentIndexCandidateCount,
      persistentCacheUsed: cacheCandidateCount > 0,
      cacheFresh,
      cacheAgeSeconds,
      cacheWithinStaleWindow,
      usedFreshLocalFastPath,
      usedStaleLocalFastPath,
      cacheCandidates: cacheCandidateCount,
      durationMs: elapsedMs(startedAt),
      contentIndexDurationMs,
      liveSearchDurationMs,
      scanDurationMs,
      cacheConfirmDurationMs,
      cacheConfirmations,
      graphSearchCalls: searchRuns.reduce((sum, run) => sum + (run.graphSearchCalls || 0), 0),
      searchTermsPlanned: searchTerms.length,
      searchTermsExecuted: executedSearchTerms.length,
      searchTermsSkipped: skippedSearchTerms.length,
      searchStopReason,
      scanAttempts: scanRuns.length,
      metadataCacheWrites: toolCallContext.getStore()?.metadataCacheWrites || 0
    },
    searchRuns,
    scanRuns,
    items: ranked.slice(0, maxResults).map((candidate) => ({
      ...formatSimplifiedItem(candidate.item, args.format),
      score: candidate.score,
      reasons: candidate.reasons,
      snippets: candidate.snippets?.slice(0, 2),
      sources: candidate.sources.slice(0, 5)
    })),
    note: noteParts.join(" ")
  };
}

function broadFindFolderHints(query = "", userHints = []) {
  const presets = pathPresets();
  const common = [
    presets.documents,
    presets.desktop,
    presets.pictures,
    presets.screenshots,
    "Personal/Documents",
    "Personal",
    "Microsoft Copilot Chat Files",
    ""
  ];
  return pruneFolderHints([
    ...(userHints || []),
    ...defaultFindFolderHints(query),
    ...common
  ].filter((hint) => hint !== undefined && hint !== null));
}

async function findAll(args = {}) {
  const query = String(args.query || "").trim();
  if (!query) throw new Error("query is required.");
  const result = await find({
    ...args,
    maxResults: Math.min(args.maxResults ?? 50, 200),
    maxResultsLimit: 200,
    folderHints: broadFindFolderHints(query, args.folderHints || []),
    scanFallback: true,
    scanMaxItems: Math.min(args.scanMaxItems ?? 10000, 50000),
    scanMaxFolders: Math.min(args.scanMaxFolders ?? 2000, 10000),
    scanMaxDepth: Math.min(args.scanMaxDepth ?? 25, 50),
    searchPageSize: Math.min(args.searchPageSize ?? 100, 200),
    searchMaxItemsPerTerm: Math.min(args.searchMaxItemsPerTerm ?? 250, 1000),
    minConfidenceForSearchOnly: args.minConfidenceForSearchOnly ?? 78,
    executeAllSearchTerms: true
  });
  return {
    ...result,
    strategy: "broad-cache-assisted-remote-first",
    folderPlan: broadFindFolderHints(query, args.folderHints || []).map((folder) => folder || "root")
  };
}

function formatDeltaItem(item, format = "compact") {
  const formatted = formatDriveItem(item, format);
  if (format === "full") return formatted ? { ...formatted, deleted: item.deleted } : formatted;
  return formatted ? { ...formatted, deleted: item.deleted ? item.deleted : undefined } : formatted;
}

async function resolveDeltaPathRoot(args = {}) {
  try {
    if (args.itemId) return await resolveScanRoot({ itemId: args.itemId, cacheResults: false });
    const resolvedPath = resolvePresetPath(args);
    if (resolvedPath) return await resolveScanRoot({ path: resolvedPath, cacheResults: false });
    return await resolveScanRoot({ cacheResults: false });
  } catch (error) {
    recordLocalWarning("delta path-root resolution", error);
    return null;
  }
}

async function hydrateDeltaParentItems(items = [], pathRoot = null) {
  if (!items.length) return items;
  const cache = await loadMetadataCache();
  const combined = new Map(items.filter((item) => item?.id).map((item) => [item.id, item]));
  const supplemental = new Map();
  const knownRootIds = new Set([
    ...Object.keys(cache.pathRootsById || {}),
    ...(cache.scanRoot?.id ? [cache.scanRoot.id] : []),
    ...(pathRoot?.id ? [pathRoot.id] : [])
  ]);
  const maxHydratedParents = Math.min(100, Math.max(10, items.length * 2));

  while (supplemental.size < maxHydratedParents) {
    const missingParentIds = [];
    for (const item of combined.values()) {
      if (!item?.parentReference?.id || itemRemotePath(item) !== undefined) continue;
      const parentId = item.parentReference.id;
      if (knownRootIds.has(parentId)) continue;
      const cachedParent = cache.itemsById?.[parentId];
      if (cachedParent?.remotePath !== undefined) continue;
      if (combined.has(parentId)) continue;
      if (!missingParentIds.includes(parentId)) missingParentIds.push(parentId);
      if (supplemental.size + missingParentIds.length >= maxHydratedParents) break;
    }
    if (!missingParentIds.length) break;
    const fetched = await mapWithConcurrency(missingParentIds, 4, async (parentId) => {
      try {
        return await graph(`/me/drive/items/${encodeURIComponent(parentId)}?$select=${encodeURIComponent(defaultSelect)}`);
      } catch (error) {
        recordLocalWarning("delta parent path resolution", error);
        return null;
      }
    });
    let added = 0;
    for (const parent of fetched) {
      if (!parent?.id || combined.has(parent.id)) continue;
      combined.set(parent.id, parent);
      supplemental.set(parent.id, parent);
      added += 1;
    }
    if (!added) break;
  }
  return [...supplemental.values(), ...items];
}

async function delta(args = {}) {
  const cursorCount = Number(Boolean(args.nextLink)) + Number(Boolean(args.deltaLink));
  const targetSelectorCount = Number(Boolean(args.itemId))
    + Number(Boolean(args.path))
    + Number(Boolean(args.preset));
  if (cursorCount > 1) throw new Error("Provide only one of nextLink or deltaLink.");
  if (targetSelectorCount > 1) throw new Error("Provide only one delta target: itemId, path, or preset.");
  if (cursorCount && targetSelectorCount) throw new Error("Delta cursors cannot be combined with itemId, path, or preset targets.");
  const maxItems = clampInteger(args.maxItems, 1000, 1, 5000);
  let firstPath = args.nextLink || args.deltaLink;
  let target = args._deltaTarget || (args.nextLink ? "nextLink" : args.deltaLink ? "deltaLink" : "root");
  let pathRoot = null;
  if (firstPath && !isDeltaCursor(firstPath)) {
    throw new Error("nextLink and deltaLink must be Microsoft Graph delta cursor URLs. Refusing a non-delta pagination cursor.");
  }
  if (!firstPath) {
    const params = new URLSearchParams();
    params.set("$top", String(Math.min(clampInteger(args.pageSize, 200, 1, 200), maxItems)));
    if (args.itemId) {
      firstPath = `/me/drive/items/${encodeURIComponent(args.itemId)}/delta?${params.toString()}`;
      target = `itemId:${args.itemId}`;
    } else {
      const resolvedPath = resolvePresetPath(args);
      if (resolvedPath) {
        pathRoot = await resolveDeltaPathRoot(args);
        if (!pathRoot?.id) throw new Error(`Could not resolve delta target folder: ${resolvedPath}`);
        firstPath = `/me/drive/items/${encodeURIComponent(pathRoot.id)}/delta?${params.toString()}`;
        target = resolvedPath;
      } else {
        firstPath = `/me/drive/root/delta?${params.toString()}`;
      }
    }
    if (!pathRoot) pathRoot = await resolveDeltaPathRoot(args);
    if (pathRoot) {
      await bestEffortLocalWrite("metadata cache path-root update", async () => await cacheItems([], { pathRoot }));
    }
  }
  const result = await collectPages(firstPath, maxItems, args.format, formatDeltaItem, {
    cacheResults: true,
    persistDeltaCursor: args._persistCursor === true,
    deltaTarget: target,
    maxPages: args.maxPages,
    prepareCacheItems: async (items) => await hydrateDeltaParentItems(items, pathRoot)
  });
  const cache = await loadMetadataCache();
  const unresolvedPathCount = unresolvedPathItems(cache).length;
  const resolvedItems = result.items.map((item) => {
    if (!item?.id || item.deleted) return item;
    const cached = cache.itemsById?.[item.id];
    return cached ? formatSimplifiedItem(cached, args.format) : item;
  });
  return {
    ...result,
    items: resolvedItems,
    target,
    unresolvedPathCount,
    note: result.unsafePageTruncation
      ? "Microsoft Graph returned more items than could be accepted safely from one page. No continuation was returned because it would skip items; rerun with a larger maxItems."
      : result.deltaLink
        ? "Save deltaLink to ask for changes since this point later."
        : "Use nextLink to continue this delta scan before saving a deltaLink."
  };
}

async function cacheRefresh(args = {}) {
  const startedAt = Date.now();
  const existing = await loadMetadataCache();
  const invalidCursorMetadata = {};
  if (existing.deltaLink && !isDeltaCursor(existing.deltaLink)) {
    existing.deltaLink = null;
    invalidCursorMetadata.deltaLink = null;
  }
  if (existing.deltaNextLink && !isDeltaCursor(existing.deltaNextLink)) {
    existing.deltaNextLink = null;
    invalidCursorMetadata.deltaNextLink = null;
  }
  if ((!existing.deltaLink && !existing.deltaNextLink) && existing.deltaTarget) {
    existing.deltaTarget = null;
    invalidCursorMetadata.deltaTarget = null;
  }
  if (Object.keys(invalidCursorMetadata).length) {
    await cacheItems([], invalidCursorMetadata);
  }
  const settings = pluginSettings();
  const mode = args.mode || "auto";
  const progress = [];
  const addProgress = (stage, details = {}) => {
    progress.push({ stage, elapsedMs: elapsedMs(startedAt), ...details });
  };
  const requestedTarget = args.itemId
    ? `itemId:${args.itemId}`
    : (resolvePresetPath(args) || "root");
  const cachedTarget = existing.deltaTarget || existing.scanRoot?.target || "root";
  const targetMatchesCache = requestedTarget === cachedTarget
    && (!existing.deltaTarget || existing.deltaTarget === requestedTarget);
  if ((mode === "delta" || mode === "auto") && settings.deltaSyncEnabled && (existing.deltaNextLink || existing.deltaLink) && targetMatchesCache) {
    const continuingNextLink = Boolean(existing.deltaNextLink);
    addProgress(continuingNextLink ? "delta-resume-nextLink" : "delta-start", {
      cursor: continuingNextLink ? "deltaNextLink" : "deltaLink"
    });
    const result = await delta({
      ...(continuingNextLink ? { nextLink: existing.deltaNextLink } : { deltaLink: existing.deltaLink }),
      pageSize: clampInteger(args.pageSize, 200, 1, 200),
      maxItems: clampInteger(args.maxItems, 10000, 1, 50000),
      maxPages: clampInteger(args.maxPages, 10, 1, 100),
      format: "full",
      _persistCursor: true,
      _deltaTarget: requestedTarget
    });
    addProgress(result.deltaLink ? "delta-complete" : "delta-incomplete", {
      count: result.count,
      hasNextLink: Boolean(result.nextLink),
      hasDeltaLink: Boolean(result.deltaLink)
    });
    return {
      mode: "delta",
      result,
      cache: await syncStatus(),
      progress,
      durationMs: elapsedMs(startedAt),
      note: result.deltaLink ? "Applied delta changes and stored the latest deltaLink." : "Delta scan is incomplete; run again with nextLink before treating cache as fresh."
    };
  }

  if (mode === "delta") {
    if (!settings.deltaSyncEnabled) {
      throw new Error("Delta sync is disabled by configuration.");
    }
    if (!existing.deltaLink && !existing.deltaNextLink) {
      throw new Error("No cached deltaLink or deltaNextLink exists yet. Run onedrive_cache_refresh with mode: scan or auto first.");
    }
    throw new Error(`Cached deltaLink is for ${cachedTarget}, not requested target ${requestedTarget}. Run onedrive_cache_refresh with mode: scan for this target.`);
  }

  if (args.replaceCache === true) await clearMetadataCache();
  addProgress("scan-start", { target: requestedTarget, replaceCache: args.replaceCache === true });
  const result = await scan({
    ...args,
    format: "full",
    includeFiles: true,
    includeFolders: args.includeFolders !== false,
    maxItems: clampInteger(args.maxItems, 10000, 1, 50000),
    maxFolders: clampInteger(args.maxFolders, 2000, 1, 10000),
    maxDepth: clampInteger(args.maxDepth, settings.maxScanDepth, 0, 50)
  });
  addProgress("scan-complete", {
    itemsScanned: result.summary?.itemsScanned,
    foldersVisited: result.summary?.foldersVisited,
    traversalTruncated: result.summary?.traversalTruncated
  });
  await cacheItems([], {
    scanRoot: result.root,
    deltaLink: null,
    deltaNextLink: null,
    deltaTarget: requestedTarget
  });

  try {
    if (settings.deltaSyncEnabled) {
      addProgress("delta-prime-start");
      const deltaResult = await delta({
        ...args,
        pageSize: clampInteger(args.pageSize, 200, 1, 200),
        maxItems: clampInteger(args.maxItems, 10000, 1, 50000),
        maxPages: clampInteger(args.maxPages, 10, 1, 100),
        format: "full",
        _persistCursor: true,
        _deltaTarget: requestedTarget
      });
      const deltaComplete = Boolean(deltaResult.deltaLink);
      addProgress(deltaComplete ? "delta-prime-complete" : "delta-prime-incomplete", {
        count: deltaResult.count,
        hasNextLink: Boolean(deltaResult.nextLink),
        hasDeltaLink: Boolean(deltaResult.deltaLink)
      });
      return {
        mode,
        effectiveMode: "scan",
        scan: result,
        deltaRefresh: { attempted: true, complete: deltaComplete, count: deltaResult.count, deltaLink: deltaResult.deltaLink, nextLink: deltaResult.nextLink },
        cache: await syncStatus(),
        progress,
        durationMs: elapsedMs(startedAt),
        note: deltaComplete
          ? `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan and stored a fresh deltaLink.`
          : `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan. Delta refresh returned nextLink but no final deltaLink yet.`
      };
    }
    return {
      mode: "scan",
      effectiveMode: "scan",
      scan: result,
      cache: await syncStatus(),
      progress,
      durationMs: elapsedMs(startedAt),
      note: `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan. Delta sync is disabled by configuration.`
    };
  } catch (error) {
    addProgress("delta-prime-error", { error: safeToolErrorMessage(error) });
    return {
      mode: "scan",
      effectiveMode: "scan",
      scan: result,
      cache: await syncStatus(),
      deltaRefresh: { attempted: true, complete: false },
      deltaError: safeToolErrorMessage(error),
      progress,
      durationMs: elapsedMs(startedAt),
      note: `${args.replaceCache === true ? "Rebuilt" : "Merged"} the metadata cache from a bounded scan, but could not get a fresh deltaLink.`
    };
  }
}

function batchResponseHeader(headers = {}, name) {
  if (typeof headers?.get === "function") return headers.get(name);
  const target = String(name).toLowerCase();
  const entry = Object.entries(headers || {}).find(([key]) => key.toLowerCase() === target);
  return entry?.[1] ?? null;
}

function batchRetryDelayMs(responses, attempt) {
  const delays = responses.map((response) => {
    const retryAfter = batchResponseHeader(response.headers, "retry-after");
    if (retryAfter === null || retryAfter === undefined || retryAfter === "") return null;
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const dateMs = Date.parse(retryAfter);
    return Number.isFinite(dateMs) ? Math.max(0, dateMs - Date.now()) : null;
  }).filter((value) => value !== null);
  if (delays.length) return Math.max(...delays);
  return Math.min(1000 * 2 ** attempt, 8000);
}

function isReadOnlyBatchRequest(request = {}) {
  return ["GET", "HEAD"].includes(String(request.method || "GET").toUpperCase());
}

function isTransientBatchResponse(response = {}) {
  return response.status === 0 || [429, 500, 502, 503, 504].includes(response.status);
}

async function batchGraph(requests = []) {
  if (!requests.length) return [];
  if (requests.length > 20) throw new Error("Microsoft Graph batch requests support at most 20 subrequests.");
  const entries = requests.map((request, index) => ({ request, index, id: String(index + 1) }));
  const finalResponses = new Array(requests.length);
  let pending = entries;
  const maxRetries = 3;

  for (let attempt = 0; pending.length; attempt += 1) {
    const allReadOnly = pending.every((entry) => isReadOnlyBatchRequest(entry.request));
    const result = await graph("/$batch", {
      method: "POST",
      isMutation: !allReadOnly,
      body: JSON.stringify({
        requests: pending.map(({ request, id }) => ({
          id,
          method: request.method || "GET",
          url: String(request.url || "").replace(/^\/+/, ""),
          ...(request.headers ? { headers: request.headers } : {}),
          ...(request.body !== undefined ? { body: request.body } : {})
        }))
      }),
      ...(allReadOnly ? { maxRetries: 3 } : {})
    });
    const responses = new Map((result.responses || []).map((response) => [String(response.id), response]));
    const retryEntries = [];
    const retryResponses = [];

    for (const entry of pending) {
      const response = responses.get(entry.id);
      const normalized = response ? {
        ok: response.status >= 200 && response.status < 300,
        status: response.status,
        body: response.body,
        headers: response.headers,
        request: entry.request
      } : {
        ok: false,
        status: 0,
        error: "Missing batch response.",
        request: entry.request
      };
      if (attempt < maxRetries && isReadOnlyBatchRequest(entry.request) && isTransientBatchResponse(normalized)) {
        retryEntries.push(entry);
        retryResponses.push(normalized);
      } else {
        finalResponses[entry.index] = normalized;
      }
    }

    if (!retryEntries.length) break;
    await sleep(batchRetryDelayMs(retryResponses, attempt));
    pending = retryEntries;
  }

  return finalResponses;
}

async function batchGetInfo(args = {}) {
  const items = args.items || [];
  const responses = await batchGraph(items.map((target) => ({ url: itemBase(target), target })));
  const rawItems = responses.filter((response) => response.ok && response.body).map((response) => response.body);
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(rawItems));
  return {
    count: responses.length,
    items: responses.map((response) => response.ok
      ? formatDriveItem(response.body, args.format || "compact")
      : { error: redactAuditText(response.body?.error?.message || response.error || `HTTP ${response.status}`), status: response.status, target: response.request.target })
  };
}

async function batchPermissions(args = {}) {
  const items = args.items || [];
  const responses = await batchGraph(items.map((target) => ({ url: `${itemBase(target)}/permissions`, target })));
  return {
    count: responses.length,
    items: responses.map((response) => response.ok
      ? {
          target: response.request.target,
          permissions: (response.body?.value || []).map((permission) => simplifyPermission(permission, args.format)),
          count: (response.body?.value || []).length
        }
      : { error: redactAuditText(response.body?.error?.message || response.error || `HTTP ${response.status}`), status: response.status, target: response.request.target })
  };
}

function uniqueBatchLocalPath(destinationFolder, name, index, usedTargets) {
  const parsed = parse(name || `onedrive-download-${index + 1}`);
  const base = parsed.name || "onedrive-download";
  const ext = parsed.ext || "";
  let suffix = 1;
  let candidate;
  do {
    const fileName = suffix === 1 ? `${base}${ext}` : `${base} (${suffix})${ext}`;
    candidate = resolve(join(destinationFolder, fileName));
    suffix += 1;
  } while (usedTargets.has(candidate));
  usedTargets.add(candidate);
  return candidate;
}

async function batchDownload(args = {}) {
  const destinationFolder = args.destinationFolder ? resolve(args.destinationFolder) : null;
  if (destinationFolder) await assertNotLocalOneDriveSyncPathForWrite(destinationFolder, "Batch download", args);
  const generatedDestinationFolder = destinationFolder || downloadRoot;
  const plannedTargets = new Set();
  const results = [];
  for (const [index, item] of (args.items || []).entries()) {
    try {
      const explicitLocalPath = item.localPath ? resolve(item.localPath) : null;
      if (explicitLocalPath) await assertNotLocalOneDriveSyncPathForWrite(explicitLocalPath, "Batch download", args);
      const info = await getInfo(item);
      const localPath = explicitLocalPath
        ? explicitLocalPath
        : uniqueBatchLocalPath(generatedDestinationFolder, info.name, index, plannedTargets);
      if (localPath) {
        await assertNotLocalOneDriveSyncPathForWrite(localPath, "Batch download", args);
        if (explicitLocalPath) {
          if (plannedTargets.has(localPath)) throw new Error(`Batch download target is used more than once: ${localPath}`);
          plannedTargets.add(localPath);
        }
      }
      results.push(await downloadResolvedItem(info, {
        itemId: info.id,
        localPath,
        overwrite: item.overwrite ?? args.overwrite,
        allowLocalOneDriveSyncPath: args.allowLocalOneDriveSyncPath
      }));
    } catch (error) {
      results.push({ target: item, error: safeToolErrorMessage(error) });
    }
  }
  return { count: results.length, results };
}

async function batchDelete(args = {}) {
  const items = args.items || [];
  const warnings = batchMutationWarnings();
  if (args.dryRun === false) {
    if (args.confirmed !== true) {
      return {
        dryRun: false,
        confirmed: false,
        count: items.length,
        warnings,
        requiredToDelete: "Set dryRun: false and confirmed: true after explicit user confirmation."
      };
    }
    const missingExpected = items.filter((item) => !item.expectedName && !item.expectedId);
    if (missingExpected.length) {
      return {
        dryRun: false,
        confirmed: true,
        count: items.length,
        warnings,
        requiredToDelete: "Provide expectedName or expectedId for every item in a live batch delete.",
        missingExpectedCount: missingExpected.length
      };
    }
  }

  const preflight = [];
  const preflightErrors = [];
  const resolvedIds = new Set();
  for (const [index, item] of items.entries()) {
    try {
      requireNonRootTarget(item, "Delete");
      const rawItem = await getRawInfo(item);
      if (rawItem.root) throw new Error("Delete refuses to operate on the OneDrive root.");
      assertExpectedItem(rawItem, item, "Delete");
      if (resolvedIds.has(rawItem.id)) throw new Error(`Batch delete target resolves to duplicate item ID ${rawItem.id}.`);
      resolvedIds.add(rawItem.id);
      preflight.push({ targetArgs: item, rawItem, item: simplifyItem(rawItem) });
    } catch (error) {
      preflightErrors.push({ index, target: item, error: safeToolErrorMessage(error) });
    }
  }
  if (preflightErrors.length) {
    return {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: preflightErrors,
      requiredToDelete: "Fix every preflight error before running a batch delete."
    };
  }
  const previewProof = {
    items: preflight.map((entry) => itemVersionProof(entry.rawItem)),
    operation: "batch-delete"
  };

  if (args.dryRun !== false) {
    return {
      dryRun: true,
      confirmed: args.confirmed === true,
      count: preflight.length,
      warnings,
      ...issuePreviewToken("onedrive_batch_delete", previewProof),
      results: preflight.map((entry) => ({ wouldDelete: entry.item }))
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(
    { dryRun: false, confirmed: true, count: preflight.length, warnings, results: preflight.map((entry) => ({ wouldDelete: entry.item })) },
    "onedrive_batch_delete",
    previewProof,
    args.previewToken,
    "requiredToDelete"
  );
  if (previewTokenRequired) return previewTokenRequired;

  const results = [];
  for (const [index, entry] of preflight.entries()) {
    try {
      await graph(itemMutationBase(entry.rawItem), { method: "DELETE", headers: mutationMatchHeaders(entry.rawItem) });
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([{ ...entry.rawItem, deleted: {} }]));
      results.push({ deleted: entry.item });
    } catch (error) {
      const failure = {
        dryRun: false,
        confirmed: true,
        count: preflight.length,
        warnings,
        failed: true,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        partialResults: results
      };
      await writeMutationAudit("onedrive_batch_delete", {
        status: "failed",
        targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
        partialResults: results.map((result) => ({ deleted: itemAuditSummary(result.deleted) })),
        failedIndex: index,
        error: safeErrorInfo(error)
      });
      return failure;
    }
  }
  await writeMutationAudit("onedrive_batch_delete", {
    status: "success",
    targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
    results: results.map((result) => ({ deleted: itemAuditSummary(result.deleted) }))
  });
  return { dryRun: false, confirmed: true, count: results.length, warnings, results };
}

async function getRawInfo(args = {}) {
  const resolved = itemArgsWithResolvedPath(args);
  if (!resolved.path && !resolved.itemId) throw new Error("Provide path, preset, or itemId.");
  if (!args.includeDeletedItems && args.useCache === true) {
    const cached = resolved.itemId ? await cachedItemById(resolved.itemId) : await cachedItemByPath(resolved.path);
    if (cached) return cached;
  }
  const suffix = args.includeDeletedItems && resolved.itemId ? "?includeDeletedItems=true" : "";
  const item = await graph(`${itemBase(args)}${suffix}`);
  if (args.cacheResults !== false) {
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([item]));
  }
  return item;
}

async function getInfo(args = {}) {
  return formatDriveItem(await getRawInfo(args), args.format || "full");
}

function compactIdentity(identity) {
  if (!identity) return undefined;
  const typed = identity.user ? { type: "user", ...identity.user }
    : identity.group ? { type: "group", ...identity.group }
      : identity.siteUser ? { type: "siteUser", ...identity.siteUser }
        : identity.siteGroup ? { type: "siteGroup", ...identity.siteGroup }
          : identity.application ? { type: "application", ...identity.application }
            : identity.device ? { type: "device", ...identity.device }
              : null;
  return typed ? {
    type: typed.type,
    id: typed.id,
    displayName: typed.displayName,
    email: typed.email,
    loginName: typed.loginName
  } : identity;
}

function simplifyPermission(permission, format = "compact") {
  if (format === "full") return permission;
  const roles = new Set(permission.roles || []);
  const permissionKind = roles.has("owner")
    ? "owner"
    : permission.inheritedFrom
      ? "inherited"
      : permission.link?.scope === "anonymous"
        ? "anonymous_link"
        : permission.link
          ? "sharing_link"
          : permission.invitation
            ? "invitation"
            : "direct";
  const revocable = permissionKind !== "owner" && permissionKind !== "inherited";
  const link = permission.link && (permission.link.type || permission.link.scope) ? {
    type: permission.link.type,
    scope: permission.link.scope,
    webUrl: permission.link.webUrl,
    preventsDownload: permission.link.preventsDownload
  } : undefined;
  return {
    id: permission.id,
    roles: permission.roles,
    permissionKind,
    revocable,
    link,
    grantedTo: compactIdentity(permission.grantedTo),
    grantedToV2: compactIdentity(permission.grantedToV2),
    grantedToIdentities: permission.grantedToIdentities?.map(compactIdentity),
    grantedToIdentitiesV2: permission.grantedToIdentitiesV2?.map(compactIdentity),
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

async function permissionList(args = {}, format = "compact") {
  const result = await graph(`${itemBase(args)}/permissions`);
  return (result.value || []).map((permission) => simplifyPermission(permission, format));
}

function permissionKey(permission = {}) {
  return permission.id
    || permission.link?.webUrl
    || `${(permission.roles || []).join(",")}:${permission.link?.type || ""}:${permission.link?.scope || ""}:${permission.grantedTo?.email || permission.invitation?.email || ""}`;
}

function isExplicitSharingPermission(permission = {}) {
  if (permission.inheritedFrom) return false;
  if (permission.permissionKind === "owner") return false;
  const roles = new Set(permission.roles || []);
  if (roles.has("owner")) return false;
  if (permission.link?.type || permission.link?.scope || permission.invitation) return true;
  return Boolean(
    permission.grantedTo
    || permission.grantedToV2
    || permission.grantedToIdentities?.length
    || permission.grantedToIdentitiesV2?.length
  );
}

function isOwnerPermission(permission = {}) {
  return permission.permissionKind === "owner" || new Set(permission.roles || []).has("owner");
}

function isRevocablePermission(permission = {}) {
  return permission.revocable !== false && !isOwnerPermission(permission) && !permission.inheritedFrom;
}

function diffPermissions(before = [], after = []) {
  const beforeMap = new Map(before.map((permission) => [permissionKey(permission), permission]));
  const afterMap = new Map(after.map((permission) => [permissionKey(permission), permission]));
  const added = [];
  const removed = [];
  const unchanged = [];
  for (const [key, permission] of afterMap) {
    if (beforeMap.has(key)) unchanged.push(permission);
    else added.push(permission);
  }
  for (const [key, permission] of beforeMap) {
    if (!afterMap.has(key)) removed.push(permission);
  }
  return {
    added,
    removed,
    unchangedCount: unchanged.length,
    beforeCount: before.length,
    afterCount: after.length
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
  if (args.destinationParentRelativePath !== undefined && !args.destinationParentPreset) {
    throw new Error("destinationParentRelativePath requires destinationParentPreset.");
  }
  if (args.parentRelativePath !== undefined && !args.parentPreset) {
    throw new Error("parentRelativePath requires parentPreset.");
  }
  assertAtMostOneSelector(args, "destination parent", [
    { label: "destinationParentItemId", keys: ["destinationParentItemId"] },
    { label: "destinationParentPath", keys: ["destinationParentPath"] },
    { label: "destinationParentPreset", keys: ["destinationParentPreset"] },
    { label: "parentItemId", keys: ["parentItemId"] },
    { label: "parentPath", keys: ["parentPath"] },
    { label: "parentPreset", keys: ["parentPreset"] }
  ]);
  const explicitParentId = args.destinationParentItemId || args.parentItemId;
  if (explicitParentId) {
    const raw = await getRawInfo({ itemId: explicitParentId });
    if (!raw.folder && !raw.root) throw new Error(`Destination parent is not a folder: ${raw.name}`);
    return { id: raw.id, path: raw.parentReference?.path, name: raw.name };
  }
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
  const maxBytes = clampInteger(args.maxBytes, textFileLimit, 1, maxTextFileReadLimit);
  assertTextReadable(info, args);
  if (info.size && info.size > maxBytes) {
    throw new Error(`File is ${info.size} bytes, above maxBytes ${maxBytes}. Use onedrive_download instead.`);
  }
  const limited = await graphLimitedBuffer(contentPath(args), maxBytes);
  if (limited.truncated) {
    throw new Error(`Downloaded content is above maxBytes ${maxBytes}. Use onedrive_download instead.`);
  }
  const buffer = limited.buffer;
  if (args.force !== true) assertNoBinaryNulls(buffer);
  return { item: info, content: buffer.toString("utf8") };
}

async function download(args = {}) {
  if (args.localPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.localPath), "Download", args);
  const info = await getInfo(args);
  return await downloadResolvedItem(info, args);
}

async function downloadResolvedItem(info, args = {}) {
  const preferredTarget = args.localPath ? resolve(args.localPath) : join(downloadRoot, info.name || basename(cleanPath(args.path || args.itemId || "download")));
  await assertNotLocalOneDriveSyncPathForWrite(preferredTarget, "Download", args);
  const reservation = await reserveLocalDestination(preferredTarget, {
    overwrite: args.overwrite === true,
    allowAlternate: !args.localPath || args._allowAlternateLocalPath === true
  });
  const target = reservation.path;
  const contentArgs = info.id ? { itemId: info.id } : args;
  try {
    const downloaded = await graphDownloadToFile(contentPath(contentArgs), target, { reserved: reservation.reserved });
    return { item: info, localPath: target, bytesWritten: downloaded.bytesWritten };
  } catch (error) {
    if (reservation.reserved) await rm(target, { force: true });
    throw error;
  }
}

function officePackageKindFromName(name = "") {
  const extension = extname(name).toLowerCase();
  if ([".docx", ".docm"].includes(extension)) return "word";
  if ([".xlsx", ".xlsm"].includes(extension)) return "excel";
  if ([".pptx", ".pptm", ".ppsx"].includes(extension)) return "powerpoint";
  return null;
}

function commonExtractionKind(info = {}) {
  const extension = extname(info.name || "").toLowerCase();
  const mimeType = String(info.file?.mimeType || "").toLowerCase();
  if (extension === ".pdf" || mimeType === "application/pdf") return "pdf";
  if (extension === ".rtf" || ["application/rtf", "text/rtf"].includes(mimeType)) return "rtf";
  if ([".odt", ".ods", ".odp"].includes(extension) || mimeType.startsWith("application/vnd.oasis.opendocument.")) return "opendocument";
  if (extension === ".epub" || mimeType === "application/epub+zip") return "epub";
  if (extension === ".doc" || mimeType === "application/msword") return "legacy-word";
  if (extension === ".xls" || mimeType === "application/vnd.ms-excel") return "legacy-excel";
  if (extension === ".ppt" || mimeType === "application/vnd.ms-powerpoint") return "legacy-powerpoint";
  if ([".bmp", ".gif", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".webp"].includes(extension) || (mimeType.startsWith("image/") && mimeType !== "image/svg+xml")) return "image-ocr";
  return null;
}

async function extractCommonDocumentText(info, args = {}) {
  if (info.folder) throw new Error(`Item is a folder, not an extractable document: ${info.name}`);
  const kind = commonExtractionKind(info);
  if (!kind) throw new Error(`No local text extractor is available for ${info.name || "this file"}.`);
  if (Number(info.size || 0) > maxCommonExtractionBytes) {
    throw new Error(`File is ${info.size} bytes, above the ${maxCommonExtractionBytes}-byte common-document extraction limit.`);
  }
  const transactionRoot = join(officeEditingRoot, `extract-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const localPath = join(transactionRoot, basename(info.name || `document${extname(info.name || "")}`));
  try {
    await downloadResolvedItem(info, { localPath, overwrite: false });
    const extracted = await runCommonTextHelper({
      action: "extract",
      inputPath: localPath,
      kind,
      maxBytes: clampInteger(args.maxBytes, chatgptFetchTextByteLimit, 1, chatgptFetchTextByteLimit)
    }, { timeoutMs: 60_000, maxOutputBytes: 2 * 1024 * 1024 });
    const limited = truncateUtf8(extracted.text || "", chatgptFetchTextByteLimit);
    return {
      item: info,
      preview: limited.text,
      bytes: Buffer.byteLength(limited.text, "utf8"),
      bytesRead: Number(info.size || 0),
      truncated: limited.truncated || Boolean(extracted.truncated),
      source: `local-${kind}`,
      extractor: extracted.extractor || null
    };
  } finally {
    await rm(transactionRoot, { recursive: true, force: true });
  }
}

function officeRuntimeStatus() {
  let pythonAvailable = false;
  let pythonVersion = null;
  let helperAvailable = false;
  let error = null;
  try {
    pythonVersion = execFileSync(officePythonPath, ["--version"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
      timeout: 5000
    }).trim();
    pythonAvailable = true;
  } catch (runtimeError) {
    error = safeToolErrorMessage(runtimeError);
  }
  try {
    helperAvailable = readFileSync(officeHelperPath).length > 0;
  } catch (helperError) {
    error ||= safeToolErrorMessage(helperError);
  }
  return { pythonAvailable, pythonVersion, pythonPath: officePythonPath, helperAvailable, helperPath: officeHelperPath, error };
}

async function officeCapabilities() {
  const runtime = officeRuntimeStatus();
  let drive = null;
  let driveError = null;
  try {
    drive = await graph("/me/drive");
  } catch (error) {
    driveError = safeToolErrorMessage(error);
  }
  const driveType = drive?.driveType || null;
  return {
    runtime,
    account: {
      driveType,
      driveName: drive?.name || null,
      checked: Boolean(drive),
      error: driveError
    },
    backends: {
      openXml: {
        available: runtime.pythonAvailable && runtime.helperAvailable,
        formats: [".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".ppsx"],
        readOnlyToolsReady: true,
        mutationToolsReady: true,
        operations: {
          word: ["replaceText", "setParagraphText", "setParagraphStyle", "insertParagraph", "insertTable", "setTableCell", "setContentControlText", "addHyperlink", "addComment", "insertImage", "replaceImage", "createContentControl", "deleteContentControl", "createBookmark", "deleteBookmark", "insertTableRow", "deleteTableRow", "insertTableColumn", "deleteTableColumn", "setHeaderFooterText", "setSectionProperties"],
          excel: ["setCell", "setFormula", "setRange", "clearRange", "setStyle", "setNumberFormat", "addConditionalFormat", "setDataValidation", "freezePanes", "setColumnWidth", "renameSheet", "setDefinedName", "recalculate", "addTableRow", "deleteTableRow", "setTableTotals", "createChart", "updateChart", "addWorksheet", "deleteWorksheet", "addTable", "deleteTable", "mergeRange", "unmergeRange", "sortRange", "setAutoFilter", "setHyperlink", "addNote", "deleteNote", "insertImage", "formatChart", "setSheetProtection", "refreshPivot"],
          powerpoint: ["replaceText", "setShapeText", "setShapeGeometry", "setTextStyle", "addTextBox", "deleteShape", "replaceImage", "setTableCell", "setNotes", "duplicateSlide", "deleteSlide", "moveSlide", "addSlide", "addImage", "cropImage", "addTable", "insertTableRow", "deleteTableRow", "insertTableColumn", "deleteTableColumn", "setShapeAltText", "setZOrder", "groupShapes", "ungroupShape", "applySlideLayout"]
        },
        notes: [
          "Encrypted and legacy binary Office files are refused.",
          "Macro-enabled packages are inspectable but live mutation will require explicit macro-preservation safeguards.",
          "Digitally signed packages are detected and mutation is always refused."
        ]
      },
      graphExcel: {
        availableForAccount: Boolean(drive && driveType !== "personal"),
        driveType,
        formats: [".xlsx"],
        operations: ["setCell", "setFormula", "setRange", "clearRange", "renameSheet", "addTableRow", "createChart", "updateChart"],
        limitation: driveType === "personal"
          ? "Microsoft Graph workbook APIs do not support OneDrive Consumer workbooks; use the Open XML backend."
          : "Graph workbook sessions will be used for supported business, SharePoint, or group-drive .xlsx files."
      }
    }
  };
}

async function runPythonJsonHelper(helperPath, request, options = {}) {
  const runtime = officeRuntimeStatus();
  let helperAvailable = false;
  try {
    helperAvailable = readFileSync(helperPath).length > 0;
  } catch {
    helperAvailable = false;
  }
  if (!runtime.pythonAvailable || !helperAvailable) {
    throw new Error(`Document extraction runtime is unavailable: ${runtime.error || "Python/helper not found"}`);
  }
  const timeoutMs = clampInteger(options.timeoutMs, 60_000, 1000, 5 * 60_000);
  const maxOutputBytes = clampInteger(options.maxOutputBytes, 20 * 1024 * 1024, 1024, 50 * 1024 * 1024);
  // Python bytecode is disposable runtime state. Keep it off storageRoot,
  // which can be a pre-existing read-only NAS/container mount such as /data.
  const pythonCacheRoot = resolve(
    process.env.ONEDRIVE_OFFICE_PYCACHE_ROOT || join(tmpdir(), "onedrive-python-cache")
  );
  await ensurePrivateDirectory(pythonCacheRoot);
  return await new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(officePythonPath, [helperPath], {
      cwd: pluginRoot,
      stdio: ["pipe", "pipe", "pipe"],
      env: { ...process.env, PYTHONPYCACHEPREFIX: pythonCacheRoot }
    });
    let stdout = Buffer.alloc(0);
    let stderr = Buffer.alloc(0);
    let settled = false;
    let timer = null;
    const finish = (error, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (error) rejectPromise(error);
      else resolvePromise(value);
    };
    const appendBounded = (current, chunk, label) => {
      const next = Buffer.concat([current, chunk]);
      if (next.length > maxOutputBytes) {
        child.kill("SIGKILL");
        finish(new Error(`Office helper ${label} exceeded ${maxOutputBytes} bytes.`));
      }
      return next;
    };
    child.stdout.on("data", (chunk) => { stdout = appendBounded(stdout, chunk, "stdout"); });
    child.stderr.on("data", (chunk) => { stderr = appendBounded(stderr, chunk, "stderr"); });
    child.on("error", (error) => finish(new Error(`Could not start Office helper: ${error.message}`)));
    child.on("close", (code) => {
      let parsed = null;
      try {
        parsed = JSON.parse(stdout.toString("utf8") || "{}");
      } catch (error) {
        finish(new Error(`Office helper returned invalid JSON: ${error.message}. ${stderr.toString("utf8").trim()}`));
        return;
      }
      if (code !== 0 || parsed.ok !== true) {
        finish(new Error(parsed.error || stderr.toString("utf8").trim() || `Office helper exited with code ${code}.`));
        return;
      }
      finish(null, parsed.value);
    });
    timer = setTimeout(() => {
      child.kill("SIGKILL");
      finish(new Error(`Office helper timed out after ${timeoutMs}ms.`));
    }, timeoutMs);
    child.stdin.on("error", (error) => finish(new Error(`Could not send request to Office helper: ${error.message}`)));
    child.stdin.end(JSON.stringify(request));
  });
}

export async function runOfficeHelper(request, options = {}) {
  return await runPythonJsonHelper(officeHelperPath, request, options);
}

async function runCommonTextHelper(request, options = {}) {
  return await runPythonJsonHelper(commonTextHelperPath, request, options);
}

async function inspectRemoteOfficePackage(args = {}, expectedKind = null, action = "inspect") {
  const info = args._resolvedInfo || await getInfo(args);
  if (info.folder) throw new Error(`Item is a folder, not an Office document: ${info.name}`);
  const detectedKind = officePackageKindFromName(info.name);
  if (!detectedKind) {
    throw new Error(`Office content tools require an Open XML package (.docx, .xlsx, or .pptx family). Got ${info.name || "unnamed item"}.`);
  }
  if (expectedKind && detectedKind !== expectedKind) {
    throw new Error(`Expected a ${expectedKind} Open XML package. Got ${info.name}.`);
  }
  if (Number(info.size || 0) > maxOfficePackageBytes) {
    throw new Error(`Office package is ${info.size} bytes, above the ${maxOfficePackageBytes}-byte inspection limit.`);
  }
  const transactionRoot = join(officeEditingRoot, `inspect-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const localPath = join(transactionRoot, basename(info.name));
  try {
    await downloadResolvedItem(info, { localPath, overwrite: false });
    const value = await runOfficeHelper({
      ...args,
      action,
      inputPath: localPath,
      kind: expectedKind || args.expectedKind || detectedKind
    });
    return addSemanticAnchors(expectedKind || args.expectedKind || detectedKind, {
      item: info,
      backend: "openxml",
      ...value,
      package: value.package ? { ...value.package, path: undefined } : undefined
    });
  } finally {
    await rm(transactionRoot, { recursive: true, force: true });
  }
}

function officeEditPreviewProof(rawItem, kind, operations, editResult, args = {}, backend = "openxml") {
  return {
    item: { id: rawItem.id, name: rawItem.name, eTag: rawItem.eTag, cTag: rawItem.cTag },
    kind,
    backend,
    operations,
    changeCount: editResult.changeCount,
    semanticEvidence: officeSemanticDiff(kind, editResult.changes || []),
    allowMacros: args.allowMacros === true,
    allowSignedPackage: args.allowSignedPackage === true,
    createBackup: args.createBackup !== false,
    verify: args.verify !== false
  };
}

const excelGraphOperationTypes = new Set(["setCell", "setFormula", "setRange", "clearRange", "renameSheet", "addTableRow", "createChart", "updateChart"]);
const excelGraphOnlyOperationTypes = new Set(["createChart", "updateChart"]);

async function resolveExcelOfficeBackend(rawItem, args = {}) {
  if (args.backend === "openxml") return { backend: "openxml", driveType: null, reason: "explicit" };
  for (const operation of args.operations || []) {
    if (operation.type === "clearRange" && operation.contents === false && operation.format === false) {
      throw new Error("clearRange must clear contents, formats, or both; contents:false and format:false is a no-op.");
    }
    if (operation.type === "updateChart" && operation.chartType !== undefined) {
      if (args.backend === "graph") {
        throw new Error("The Graph Excel backend cannot change an existing chart's chartType; use backend: openxml or backend: auto.");
      }
      return { backend: "openxml", driveType: null, reason: "chart-type-update" };
    }
  }
  if (extname(rawItem.name || "").toLowerCase() !== ".xlsx") {
    if (args.backend === "graph") throw new Error("The Graph Excel backend supports .xlsx only; use openxml for this workbook.");
    return { backend: "openxml", driveType: null, reason: "format" };
  }
  const unsupported = (args.operations || []).map((entry) => entry?.type).filter((type) => !excelGraphOperationTypes.has(type));
  if (unsupported.length) {
    if (args.backend === "graph") throw new Error(`The Graph Excel backend does not support these operations: ${[...new Set(unsupported)].join(", ")}.`);
    return { backend: "openxml", driveType: null, reason: "operation" };
  }
  const target = excelWorkbookTarget(rawItem);
  const driveId = target.driveId;
  let drive;
  try {
    drive = driveId ? await graph(`/drives/${encodeURIComponent(driveId)}?$select=id,driveType,name`) : await graph("/me/drive?$select=id,driveType,name");
  } catch (error) {
    if (args.backend === "graph") throw new Error(`Could not verify the workbook drive for Graph Excel: ${safeToolErrorMessage(error)}`);
    return { backend: "openxml", driveType: null, reason: "drive-check-failed" };
  }
  if (drive.driveType === "personal") {
    if (args.backend === "graph") throw new Error("Microsoft Graph workbook APIs do not support OneDrive Consumer workbooks; use backend: openxml.");
    return { backend: "openxml", driveType: drive.driveType, driveId: drive.id, reason: "consumer" };
  }
  return { backend: "graph", driveType: drive.driveType, driveId: drive.id || driveId, reason: args.backend === "graph" ? "explicit" : "supported-drive" };
}

function excelWorkbookTarget(rawItem) {
  const remoteItem = rawItem.remoteItem;
  if (remoteItem?.id && remoteItem.parentReference?.driveId) {
    return { driveId: remoteItem.parentReference.driveId, itemId: remoteItem.id };
  }
  return {
    driveId: rawItem.parentReference?.driveId || remoteItem?.parentReference?.driveId || null,
    itemId: rawItem.id
  };
}

function excelWorkbookBase(rawItem, resolvedBackend) {
  const target = excelWorkbookTarget(rawItem);
  const driveId = resolvedBackend.driveId || target.driveId;
  return driveId
    ? `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(target.itemId)}/workbook`
    : `/me/drive/items/${encodeURIComponent(target.itemId)}/workbook`;
}

function excelGraphRangePath(base, operation) {
  const sheet = String(operation.sheet || "");
  const address = String(operation.address || "").toUpperCase();
  if (!sheet || !/^[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?$/.test(address)) {
    throw new Error("Graph Excel range operations require a worksheet name and a bounded A1 address.");
  }
  const escapedAddress = address.replaceAll("'", "''");
  return `${base}/worksheets/${encodeURIComponent(sheet)}/range(address='${escapedAddress}')`;
}

function excelGraphClearApplyTo(operation) {
  const contents = operation.contents !== false;
  const formats = operation.format === true;
  if (!contents && !formats) throw new Error("clearRange must clear contents, formats, or both.");
  if (contents && formats) return "All";
  return contents ? "Contents" : "Formats";
}

function assertExcelGraphRangeVerification(operation, observed) {
  if (!observed || typeof observed !== "object") throw new Error(`Graph Excel semantic verification returned no range for ${operation.sheet}!${operation.address}.`);
  if (operation.type === "setCell" && JSON.stringify(observed.values) !== JSON.stringify([[operation.value]])) {
    throw new Error(`Graph Excel semantic verification failed for setCell ${operation.sheet}!${operation.address}.`);
  }
  if (operation.type === "setFormula" && JSON.stringify(observed.formulas) !== JSON.stringify([[operation.formula]])) {
    throw new Error(`Graph Excel semantic verification failed for setFormula ${operation.sheet}!${operation.address}.`);
  }
  if (operation.type === "setRange") {
    if (operation.values !== undefined && JSON.stringify(observed.values) !== JSON.stringify(operation.values)) throw new Error(`Graph Excel semantic verification failed for setRange values at ${operation.sheet}!${operation.address}.`);
    if (operation.formulas !== undefined && JSON.stringify(observed.formulas) !== JSON.stringify(operation.formulas)) throw new Error(`Graph Excel semantic verification failed for setRange formulas at ${operation.sheet}!${operation.address}.`);
  }
  if (operation.type === "clearRange" && excelGraphClearApplyTo(operation) !== "Formats") {
    const values = Array.isArray(observed.values) ? observed.values.flat(Infinity) : [];
    if (values.some((value) => value !== null && value !== "")) throw new Error(`Graph Excel semantic verification found uncleared contents at ${operation.sheet}!${operation.address}.`);
  }
}

function assertExcelGraphObservedFields(observed, expected, label) {
  if (!observed || typeof observed !== "object") throw new Error(`Graph Excel semantic verification returned no ${label}.`);
  for (const [key, value] of Object.entries(expected)) {
    if (!Object.hasOwn(observed, key)) {
      throw new Error(`Graph Excel semantic verification could not observe expected ${label} field ${key}.`);
    }
    if (observed[key] !== value) {
      throw new Error(`Graph Excel semantic verification failed for ${label} field ${key}.`);
    }
  }
}

function observableRangeFormat(format) {
  if (!format || typeof format !== "object") return null;
  const keys = ["columnWidth", "rowHeight", "horizontalAlignment", "verticalAlignment", "wrapText", "numberFormat"];
  const observed = Object.fromEntries(keys.filter((key) => Object.hasOwn(format, key)).map((key) => [key, format[key]]));
  return Object.keys(observed).length ? observed : null;
}

function excelGraphChartDataVerification(operation, observed = {}) {
  const limitations = [];
  const series = Array.isArray(observed?.value) ? observed.value : [];
  const exposedSourceData = observed.sourceData ?? series.find((entry) => entry?.sourceData !== undefined)?.sourceData;
  const exposedSeriesBy = observed.seriesBy ?? series.find((entry) => entry?.seriesBy !== undefined)?.seriesBy;
  const expectedSeriesBy = operation.seriesBy || "Auto";
  let sourceDataVerified = false;
  let seriesByVerified = false;
  if (exposedSourceData !== undefined) {
    if (exposedSourceData !== operation.sourceData) throw new Error("Graph Excel semantic verification failed for updated chart sourceData.");
    sourceDataVerified = true;
  } else {
    const formulas = series.flatMap((entry) => [entry?.formula, entry?.categoryFormula, entry?.xValuesFormula, entry?.bubbleSizeFormula]).filter((entry) => typeof entry === "string");
    const sourceSheet = String(operation.sourceData || "").split("!")[0].replaceAll("'", "").replace(/^=/, "").toLowerCase();
    if (formulas.length && sourceSheet && formulas.some((formula) => formula.replaceAll("'", "").toLowerCase().includes(`${sourceSheet}!`))) {
      sourceDataVerified = true;
    } else {
      limitations.push("Graph did not expose enough chart-series source metadata to verify sourceData.");
    }
  }
  if (exposedSeriesBy !== undefined) {
    if (exposedSeriesBy !== expectedSeriesBy) throw new Error("Graph Excel semantic verification failed for chart seriesBy.");
    seriesByVerified = true;
  } else {
    limitations.push("Graph did not expose seriesBy in observable chart-series metadata.");
  }
  return {
    succeeded: sourceDataVerified && seriesByVerified,
    sourceDataVerified,
    seriesByVerified,
    ...(limitations.length ? { verificationLimited: true, limitations } : {})
  };
}

async function applyExcelGraphOperations(rawItem, resolvedBackend, operations) {
  const managedSession = await acquireExcelManagedSession(rawItem, resolvedBackend, true);
  const base = managedSession.base;
  const sessionId = managedSession.id;
  const headers = { "workbook-session-id": sessionId };
  let writeStarted = false;
  const semanticVerification = [];
  try {
    for (const operation of operations) {
      if (operation.type === "addTableRow") {
        writeStarted = true;
        const addedRow = await graph(`${base}/tables/${encodeURIComponent(String(operation.table || ""))}/rows/add`, {
          method: "POST", headers, body: JSON.stringify({ values: operation.values, index: operation.index ?? null }), maxRetries: 0
        });
        if (Array.isArray(addedRow?.values) && JSON.stringify(addedRow.values) !== JSON.stringify(operation.values)) {
          throw new Error(`Graph Excel semantic verification failed for addTableRow ${operation.table}.`);
        }
        semanticVerification.push({ operation: operation.type, table: operation.table, succeeded: true });
        continue;
      }
      if (operation.type === "createChart") {
        writeStarted = true;
        const createdChart = await graph(`${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}/charts/add`, {
          method: "POST", headers, body: JSON.stringify({ type: operation.chartType, sourceData: operation.sourceData, seriesBy: operation.seriesBy || "Auto" }), maxRetries: 0
        });
        const chartId = createdChart?.id || createdChart?.name;
        if (!chartId) throw new Error("Microsoft Graph created a chart but did not return its stable ID or name.");
        const chartBase = `${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}/charts/${encodeURIComponent(String(chartId))}`;
        const chartPatch = Object.fromEntries(["name", "height", "width", "left", "top"].filter((key) => operation[key] !== undefined).map((key) => [key, operation[key]]));
        if (Object.keys(chartPatch).length) await graph(chartBase, { method: "PATCH", headers, body: JSON.stringify(chartPatch), maxRetries: 0 });
        if (operation.titleText !== undefined) await graph(`${chartBase}/title`, { method: "PATCH", headers, body: JSON.stringify({ text: operation.titleText }), maxRetries: 0 });
        const observed = await graph(chartBase, { headers, maxRetries: 0 });
        assertExcelGraphObservedFields(observed, chartPatch, "created chart");
        if (operation.titleText !== undefined) {
          const title = await graph(`${chartBase}/title`, { headers, maxRetries: 0 });
          assertExcelGraphObservedFields(title, { text: operation.titleText }, "created chart title");
        }
        const limitations = [];
        const observedChartType = observed?.chartType ?? observed?.type;
        let chartTypeVerified = false;
        if (observedChartType === undefined) {
          limitations.push("Graph did not expose the created chart type for readback.");
        } else if (observedChartType !== operation.chartType) {
          throw new Error("Graph Excel semantic verification failed for created chart type.");
        } else {
          chartTypeVerified = true;
        }
        let dataVerification;
        try {
          dataVerification = excelGraphChartDataVerification(operation, await graph(`${chartBase}/series`, { headers, maxRetries: 0 }));
        } catch (error) {
          if (String(error?.message || "").includes("semantic verification failed")) throw error;
          dataVerification = { succeeded: false, sourceDataVerified: false, seriesByVerified: false, verificationLimited: true, limitations: [`Graph chart-series readback was unavailable: ${safeToolErrorMessage(error)}`] };
        }
        limitations.push(...(dataVerification.limitations || []));
        semanticVerification.push({
          operation: operation.type,
          sheet: operation.sheet,
          chartId,
          succeeded: chartTypeVerified && dataVerification.succeeded,
          chartTypeVerified,
          sourceDataVerified: dataVerification.sourceDataVerified,
          seriesByVerified: dataVerification.seriesByVerified,
          ...(limitations.length ? { verificationLimited: true, limitations } : {})
        });
        continue;
      }
      if (operation.type === "updateChart") {
        if (operation.chartType !== undefined) throw new Error("Graph Excel cannot update chartType; use the Open XML backend.");
        const chartBase = `${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}/charts/${encodeURIComponent(String(operation.chart || ""))}`;
        const chartPatch = Object.fromEntries(["name", "height", "width", "left", "top"].filter((key) => operation[key] !== undefined).map((key) => [key, operation[key]]));
        if (Object.keys(chartPatch).length) {
          writeStarted = true;
          await graph(chartBase, { method: "PATCH", headers, body: JSON.stringify(chartPatch), maxRetries: 0 });
        }
        if (operation.titleText !== undefined) {
          writeStarted = true;
          await graph(`${chartBase}/title`, { method: "PATCH", headers, body: JSON.stringify({ text: operation.titleText }), maxRetries: 0 });
        }
        if (operation.sourceData !== undefined) {
          writeStarted = true;
          await graph(`${chartBase}/setData`, { method: "POST", headers, body: JSON.stringify({ sourceData: operation.sourceData, seriesBy: operation.seriesBy || "Auto" }), maxRetries: 0 });
        }
        const observed = await graph(chartBase, { headers, maxRetries: 0 });
        assertExcelGraphObservedFields(observed, chartPatch, "updated chart");
        if (operation.titleText !== undefined) {
          const title = await graph(`${chartBase}/title`, { headers, maxRetries: 0 });
          assertExcelGraphObservedFields(title, { text: operation.titleText }, "updated chart title");
        }
        let dataVerification = { succeeded: true };
        if (operation.sourceData !== undefined) {
          try {
            dataVerification = excelGraphChartDataVerification(operation, await graph(`${chartBase}/series`, { headers, maxRetries: 0 }));
          } catch (error) {
            if (String(error?.message || "").includes("semantic verification failed")) throw error;
            dataVerification = { succeeded: false, sourceDataVerified: false, seriesByVerified: false, verificationLimited: true, limitations: [`Graph chart-series readback was unavailable: ${safeToolErrorMessage(error)}`] };
          }
        }
        semanticVerification.push({ operation: operation.type, sheet: operation.sheet, chart: operation.chart, ...dataVerification });
        continue;
      }
      if (operation.type === "renameSheet") {
        writeStarted = true;
        await graph(`${base}/worksheets/${encodeURIComponent(String(operation.sheet || ""))}`, {
          method: "PATCH", headers, body: JSON.stringify({ name: operation.newName }), maxRetries: 0
        });
        const observed = await graph(`${base}/worksheets/${encodeURIComponent(String(operation.newName || ""))}`, { headers, maxRetries: 0 });
        if (observed?.name !== operation.newName) throw new Error(`Graph Excel semantic verification failed for renamed worksheet ${operation.newName}.`);
        semanticVerification.push({ operation: operation.type, sheet: operation.newName, succeeded: true });
        continue;
      }
      const rangePath = excelGraphRangePath(base, operation);
      if (operation.type === "clearRange") {
        const applyTo = excelGraphClearApplyTo(operation);
        let formatBefore = null;
        let formatReadError = null;
        if (applyTo !== "Contents") {
          try {
            formatBefore = observableRangeFormat(await graph(`${rangePath}/format`, { headers, maxRetries: 0 }));
          } catch (error) {
            formatReadError = safeToolErrorMessage(error);
          }
        }
        writeStarted = true;
        await graph(`${rangePath}/clear`, {
          method: "POST", headers, body: JSON.stringify({ applyTo }), maxRetries: 0
        });
        const observed = await graph(rangePath, { headers, maxRetries: 0 });
        assertExcelGraphRangeVerification(operation, observed);
        let formatVerified = applyTo === "Contents";
        const limitations = [];
        if (applyTo !== "Contents") {
          try {
            const formatAfter = observableRangeFormat(await graph(`${rangePath}/format`, { headers, maxRetries: 0 }));
            formatVerified = Boolean(formatBefore && formatAfter && JSON.stringify(formatBefore) !== JSON.stringify(formatAfter));
            if (!formatVerified) limitations.push("Graph did not expose a distinguishable post-clear format state.");
          } catch (error) {
            limitations.push(`Graph format readback was unavailable: ${safeToolErrorMessage(error)}`);
          }
          if (formatReadError) limitations.push(`Graph pre-clear format readback was unavailable: ${formatReadError}`);
        }
        semanticVerification.push({
          operation: operation.type,
          sheet: operation.sheet,
          address: operation.address,
          applyTo,
          succeeded: formatVerified,
          contentsVerified: applyTo !== "Formats",
          formatVerified,
          ...(limitations.length ? { verificationLimited: true, limitations } : {})
        });
        continue;
      }
      const body = operation.type === "setCell"
        ? { values: [[operation.value]] }
        : operation.type === "setFormula"
          ? { formulas: [[operation.formula]], ...(operation.value !== undefined ? { values: [[operation.value]] } : {}) }
          : { ...(operation.values !== undefined ? { values: operation.values } : {}), ...(operation.formulas !== undefined ? { formulas: operation.formulas } : {}) };
      writeStarted = true;
      await graph(rangePath, { method: "PATCH", headers, body: JSON.stringify(body), maxRetries: 0 });
      const observed = await graph(rangePath, { headers, maxRetries: 0 });
      assertExcelGraphRangeVerification(operation, observed);
      semanticVerification.push({ operation: operation.type, sheet: operation.sheet, address: operation.address, succeeded: true });
    }
  } catch (error) {
    const suffix = writeStarted ? " A persistent workbook session had begun; the request was not replayed." : "";
    throw new Error(`Graph Excel update failed: ${safeToolErrorMessage(error)}${suffix}`);
  } finally {
    managedSession.lastUsedAt = Date.now();
    scheduleExcelSessionClose(managedSession);
  }
  return { sessionClosed: false, sessionManaged: true, sessionPersistent: true, sessionReused: managedSession.reused, operationCount: operations.length, semanticVerification };
}

function officeSemanticDiff(kind, changes = []) {
  const operationCounts = {};
  const affectedParts = new Set();
  const affectedObjects = [];
  for (const change of changes) {
    const operation = change.operation || "unknown";
    operationCounts[operation] = (operationCounts[operation] || 0) + 1;
    if (change.part) affectedParts.add(change.part);
    const selector = Object.fromEntries(["part", "paragraphIndex", "tableIndex", "rowIndex", "columnIndex", "contentControlIndex", "sheet", "address", "name", "slideIndex", "shapeId", "commentId"]
      .filter((key) => change[key] !== undefined).map((key) => [key, change[key]]));
    if (Object.keys(selector).length) affectedObjects.push({ operation, ...selector });
  }
  return { kind, changeCount: changes.length, operationCounts, affectedParts: [...affectedParts].sort(), affectedObjects: affectedObjects.slice(0, 200), affectedObjectsTruncated: affectedObjects.length > 200 };
}

function officeInspectionSummary(kind, value = {}) {
  if (kind === "word") return { paragraphCount: value.paragraphCount || 0, tableCount: value.tableCount || 0, contentControlCount: value.contentControlCount || 0, commentCount: value.commentCount || 0 };
  if (kind === "excel") return { sheetCount: value.sheetCount || 0, sheets: (value.sheets || []).map((sheet) => ({ name: sheet.name, cellCount: sheet.cellCount ?? sheet.cells?.length ?? 0 })) };
  return { slideCount: value.slideCount || 0, slides: (value.slides || []).map((slide) => ({ index: slide.index, shapeCount: slide.shapeCount ?? slide.shapes?.length ?? 0 })) };
}

function compareOfficeInspections(kind, before = {}, after = {}, maxChanges = 100) {
  const changes = [];
  let totalChanges = 0;
  const add = (change) => { totalChanges += 1; if (changes.length < maxChanges) changes.push(change); };
  if (kind === "word") {
    const paragraphs = (value) => new Map((value.paragraphs || []).map((entry) => [`${entry.part}:${entry.index}`, entry]));
    const leftMap = paragraphs(before), rightMap = paragraphs(after);
    for (const key of new Set([...leftMap.keys(), ...rightMap.keys()])) {
      const left = leftMap.get(key), right = rightMap.get(key);
      if (left?.text !== right?.text || left?.style !== right?.style) add({ objectType: "paragraph", key, before: left ? { text: left.text, style: left.style } : null, after: right ? { text: right.text, style: right.style } : null });
    }
  } else if (kind === "excel") {
    const cells = (value) => new Map((value.sheets || []).flatMap((sheet) => (sheet.cells || []).map((cell) => [`${sheet.name}!${cell.address}`, cell])));
    const leftMap = cells(before), rightMap = cells(after);
    for (const key of new Set([...leftMap.keys(), ...rightMap.keys()])) {
      const left = leftMap.get(key), right = rightMap.get(key);
      if (JSON.stringify([left?.value, left?.formula, left?.styleIndex]) !== JSON.stringify([right?.value, right?.formula, right?.styleIndex])) add({ objectType: "cell", key, before: left ? { value: left.value, formula: left.formula, styleIndex: left.styleIndex } : null, after: right ? { value: right.value, formula: right.formula, styleIndex: right.styleIndex } : null });
    }
  } else {
    const shapes = (value) => new Map((value.slides || []).flatMap((slide) => (slide.shapes || []).map((shape) => [`${slide.index}:${shape.id}`, shape])));
    const leftMap = shapes(before), rightMap = shapes(after);
    for (const key of new Set([...leftMap.keys(), ...rightMap.keys()])) {
      const left = leftMap.get(key), right = rightMap.get(key);
      if (JSON.stringify([left?.text, left?.geometry]) !== JSON.stringify([right?.text, right?.geometry])) add({ objectType: "shape", key, before: left ? { text: left.text, geometry: left.geometry } : null, after: right ? { text: right.text, geometry: right.geometry } : null });
    }
  }
  const beforeSummary = officeInspectionSummary(kind, before), afterSummary = officeInspectionSummary(kind, after);
  return { kind, sameSemantics: totalChanges === 0 && JSON.stringify(beforeSummary) === JSON.stringify(afterSummary), before: beforeSummary, after: afterSummary, changeCount: totalChanges, changes, truncated: totalChanges > changes.length };
}

function assertOfficeBackupId(backupId) {
  const value = String(backupId || "");
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)) throw new Error("backupId is invalid.");
  return value.toLowerCase();
}

function officeBackupManifestPath(backupId) {
  return join(backupRoot, `office-${assertOfficeBackupId(backupId)}.json`);
}

async function persistOfficeBackup(localPath, rawItem, kind, fingerprint, reason = "edit") {
  await ensurePrivateDirectory(backupRoot);
  const scope = await activeStorageScope();
  const backupId = randomUUID();
  const fileName = `office-${backupId}-${basename(rawItem.name || `${kind}.bin`)}`;
  const backupPath = join(backupRoot, fileName);
  await copyFile(localPath, backupPath);
  await chmod(backupPath, 0o600);
  const manifest = { version: 2, scope, backupId, createdAt: new Date().toISOString(), reason, kind, fileName, bytes: (await stat(backupPath)).size, fingerprint: fingerprint || null, item: { id: rawItem.id, name: rawItem.name, remotePath: itemRemotePath(rawItem) || null, eTag: rawItem.eTag || null, cTag: rawItem.cTag || null, size: rawItem.size ?? null, lastModifiedDateTime: rawItem.lastModifiedDateTime || null } };
  await writePrivateFileAtomic(officeBackupManifestPath(backupId), JSON.stringify(manifest));
  return { backupId, createdAt: manifest.createdAt, kind, item: manifest.item, bytes: manifest.bytes, localPath: backupPath, reason };
}

async function loadOfficeBackup(backupId) {
  const id = assertOfficeBackupId(backupId);
  const manifestPath = officeBackupManifestPath(id);
  let manifest;
  try {
    const manifestInfo = await lstat(manifestPath);
    if (!manifestInfo.isFile() || manifestInfo.isSymbolicLink()) throw new Error("unsafe manifest");
    manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  }
  catch { throw new Error(`Office backup ${id} was not found or its manifest is invalid.`); }
  if (manifest.version !== 2 || !manifest.scope) {
    throw new Error(`Office backup ${id} uses a legacy unscoped manifest and cannot be used. Create a new backup under the current account.`);
  }
  const scope = await activeStorageScope();
  if (!storageScopesEqual(manifest.scope, scope)) throw new Error(`Office backup ${id} belongs to a different OneDrive authentication context or drive.`);
  if (manifest.backupId !== id || !manifest.item?.id || !["word", "excel", "powerpoint"].includes(manifest.kind)) throw new Error(`Office backup ${id} has an invalid manifest.`);
  if (!manifest.fileName || basename(manifest.fileName) !== manifest.fileName || !manifest.fileName.startsWith(`office-${id}-`)) throw new Error(`Office backup ${id} has an unsafe file reference.`);
  const localPath = join(backupRoot, manifest.fileName);
  const info = await lstat(localPath);
  if (!info.isFile() || info.isSymbolicLink() || info.size > maxOfficePackageBytes) throw new Error(`Office backup ${id} is missing, unsafe, or exceeds the Office size limit.`);
  const [resolvedRoot, resolvedFile] = await Promise.all([realpath(backupRoot), realpath(localPath)]);
  if (dirname(resolvedFile) !== resolvedRoot) throw new Error(`Office backup ${id} resolves outside managed backup storage.`);
  return { manifest, localPath };
}

async function officeBackups(args = {}) {
  await ensurePrivateDirectory(backupRoot);
  const entries = await readdir(backupRoot, { withFileTypes: true });
  const items = [];
  for (const entry of entries) {
    const match = entry.isFile() && entry.name.match(/^office-([0-9a-f-]{36})\.json$/i);
    if (!match) continue;
    try {
      const { manifest } = await loadOfficeBackup(match[1]);
      if (args.itemId && manifest.item.id !== args.itemId) continue;
      if (args.kind && manifest.kind !== args.kind) continue;
      items.push({ backupId: manifest.backupId, createdAt: manifest.createdAt, reason: manifest.reason, kind: manifest.kind, bytes: manifest.bytes, fingerprint: manifest.fingerprint, item: manifest.item });
    } catch { /* Ignore malformed legacy/unmanaged entries. */ }
  }
  items.sort((left, right) => String(right.createdAt).localeCompare(String(left.createdAt)));
  const limit = clampInteger(args.limit, 100, 1, 500);
  return { count: Math.min(items.length, limit), total: items.length, items: items.slice(0, limit), truncated: items.length > limit };
}

async function inspectLocalOffice(localPath, kind) {
  return await runOfficeHelper({ action: "inspect", inputPath: localPath, kind, maxParagraphs: 10000, maxCells: 50000, maxSlides: 5000 });
}

async function officeCompareBackup(args = {}) {
  const { manifest, localPath } = await loadOfficeBackup(args.backupId);
  const current = await getRawInfo({ itemId: manifest.item.id });
  if (officePackageKindFromName(current.name) !== manifest.kind) throw new Error("Current remote item no longer has the backup's Office document kind.");
  const transactionRoot = join(officeEditingRoot, `compare-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const currentPath = join(transactionRoot, basename(current.name));
  try {
    await downloadResolvedItem(simplifyItem(current), { localPath: currentPath, overwrite: false });
    const [backupInspection, currentInspection] = await Promise.all([inspectLocalOffice(localPath, manifest.kind), inspectLocalOffice(currentPath, manifest.kind)]);
    return { backup: { backupId: manifest.backupId, createdAt: manifest.createdAt, kind: manifest.kind, bytes: manifest.bytes, fingerprint: backupInspection.package?.fingerprint, item: manifest.item }, current: { item: simplifyItem(current), fingerprint: currentInspection.package?.fingerprint }, sameContent: backupInspection.package?.fingerprint === currentInspection.package?.fingerprint, semanticDiff: compareOfficeInspections(manifest.kind, backupInspection, currentInspection, clampInteger(args.maxChanges, 100, 1, 500)) };
  } finally { await rm(transactionRoot, { recursive: true, force: true }); }
}

async function officeRestoreBackup(args = {}) {
  const { manifest, localPath } = await loadOfficeBackup(args.backupId);
  const current = await getRawInfo({ itemId: manifest.item.id });
  const backupValidation = await runOfficeHelper({ action: "validate", inputPath: localPath, kind: manifest.kind });
  const comparison = await officeCompareBackup({ backupId: manifest.backupId, maxChanges: 100 });
  const preview = { dryRun: args.dryRun !== false, confirmed: args.confirmed === true, wouldRestore: { backupId: manifest.backupId, createdAt: manifest.createdAt, kind: manifest.kind, originalItem: manifest.item, currentItem: simplifyItem(current), backupFingerprint: backupValidation.package?.fingerprint }, sameContent: comparison.sameContent, semanticDiff: comparison.semanticDiff };
  const proof = { backupId: manifest.backupId, itemId: current.id, currentETag: current.eTag, backupFingerprint: backupValidation.package?.fingerprint };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_office_restore_backup", proof);
  if (args.confirmed !== true) return { ...preview, requiredToRestore: "Set dryRun: false and confirmed: true after reviewing the Office backup preview." };
  if (!args.expectedId || args.expectedId !== current.id || args.expectedId !== manifest.item.id) return { ...preview, dryRun: false, requiredToRestore: "Provide expectedId matching the backup's original stable item ID." };
  if (!args.expectedETag || args.expectedETag !== current.eTag) return { ...preview, dryRun: false, requiredToRestore: "Provide expectedETag matching the current remote item eTag." };
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_office_restore_backup", proof, args.previewToken, "requiredToRestore");
  if (tokenRequired) return tokenRequired;
  const transactionRoot = join(officeEditingRoot, `restore-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  let rollbackBackup = null, verificationIncomplete = false, afterRaw = current;
  try {
    const currentPath = join(transactionRoot, `current-${basename(current.name)}`);
    await downloadResolvedItem(simplifyItem(current), { localPath: currentPath, overwrite: false });
    const currentValidation = await runOfficeHelper({ action: "validate", inputPath: currentPath, kind: manifest.kind });
    rollbackBackup = await persistOfficeBackup(currentPath, current, manifest.kind, currentValidation.package?.fingerprint, "pre-restore");
    const uploaded = await upload({ localPath, itemId: current.id, remotePath: itemRemotePath(current), conflictBehavior: "replace", ifMatch: current.eTag, auditTool: "onedrive_office_restore_backup", skipAudit: true });
    afterRaw = uploaded.item ? { ...current, ...uploaded.item } : current;
    try { afterRaw = await getRawInfo({ itemId: current.id }); } catch (error) { verificationIncomplete = true; recordLocalWarning("Office backup restore metadata verification", error); }
    let remoteValidation = null;
    if (args.verify !== false) {
      try {
        const verifyPath = join(transactionRoot, `verify-${basename(current.name)}`);
        await downloadResolvedItem(simplifyItem(afterRaw), { localPath: verifyPath, overwrite: false });
        remoteValidation = await runOfficeHelper({ action: "validate", inputPath: verifyPath, kind: manifest.kind });
        if (remoteValidation.package?.fingerprint !== backupValidation.package?.fingerprint) { verificationIncomplete = true; recordLocalWarning("Office backup restore fingerprint verification", new Error("Restored package fingerprint differs from the selected backup.")); }
      } catch (error) { verificationIncomplete = true; recordLocalWarning("Office backup restore validation", error); }
    }
    await writeMutationAudit("onedrive_office_restore_backup", { status: "success", target: itemAuditSummary(current), before: itemAuditSummary(current), after: itemAuditSummary(afterRaw), backupId: manifest.backupId, rollbackBackupId: rollbackBackup.backupId, verificationIncomplete });
    return { dryRun: false, confirmed: true, restoredBackupId: manifest.backupId, item: simplifyItem(afterRaw), rollbackBackup, backupValidation, remoteValidation, verificationIncomplete, uploaded };
  } catch (error) {
    await writeMutationAudit("onedrive_office_restore_backup", { status: "failed", target: itemAuditSummary(current), backupId: manifest.backupId, rollbackBackupId: rollbackBackup?.backupId, error: safeErrorInfo(error) });
    throw error;
  } finally { await rm(transactionRoot, { recursive: true, force: true }); }
}

function trustedExcelSessionOperationUrl(value, label) {
  if (!value) throw new Error(`Microsoft Graph returned an asynchronous Excel session response without ${label}.`);
  try {
    return graphUrl(value);
  } catch (error) {
    throw new Error(`Microsoft Graph returned an untrusted Excel session ${label}: ${safeToolErrorMessage(error)}`);
  }
}

async function createExcelWorkbookSession(base, persistChanges = true) {
  // Microsoft documents 504 as safe to retry specifically for createSession.
  const requestOptions = {
    method: "POST",
    body: JSON.stringify({ persistChanges }),
    headers: { Prefer: "respond-async" },
    maxRetries: 0,
    returnResponse: true
  };
  let response = await graph(`${base}/createSession`, requestOptions);
  if (response.status === 504) response = await graph(`${base}/createSession`, requestOptions);
  if (!response.ok) {
    throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Excel session creation failed" });
  }
  if (response.status !== 202) return response.body;

  const operationUrl = trustedExcelSessionOperationUrl(response.headers.get("location"), "Location URL");
  const timeoutMs = clampInteger(process.env.ONEDRIVE_EXCEL_SESSION_TIMEOUT_MS, 60_000, 1000, 5 * 60_000);
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const operation = await graph(operationUrl, { returnResponse: true, maxRetries: 3 });
    if (!operation.ok) {
      throw microsoftGraphError(operation.body, { headers: operation.headers, status: operation.status, statusText: "Excel session operation failed" });
    }
    const status = String(operation.body?.status || "");
    if (status === "succeeded") {
      const resourceLocation = operation.body?.resourceLocation || operation.headers.get("resource-location");
      const resourceUrl = trustedExcelSessionOperationUrl(resourceLocation, "resourceLocation URL");
      return await graph(resourceUrl, { maxRetries: 3 });
    }
    if (status === "failed") {
      const detail = operation.body?.error?.message || operation.body?.error?.code || "the operation reported failure";
      throw new Error(`Microsoft Graph Excel session creation failed: ${detail}`);
    }
    if (status && !["notStarted", "running", "inProgress"].includes(status)) {
      throw new Error(`Microsoft Graph returned an unknown Excel session operation status: ${status}`);
    }
    const configuredDelay = clampInteger(process.env.ONEDRIVE_EXCEL_SESSION_POLL_MS, 2000, 0, 10_000);
    await sleep(operation.headers.get("retry-after") ? retryDelayMs(operation, 0) : configuredDelay);
  }
  throw new Error(`Microsoft Graph Excel session creation did not complete within ${timeoutMs}ms.`);
}

async function excelSessionKey(rawItem, resolvedBackend, persistent) {
  const scope = await activeStorageScope();
  const target = excelWorkbookTarget(rawItem);
  return `${storageScopeKey(scope)}:${resolvedBackend.driveId || target.driveId || "me"}:${target.itemId}:${persistent ? "write" : "preview"}`;
}

async function closeExcelManagedSession(entry, warningOnly = true) {
  if (!entry || entry.closed) return true;
  clearTimeout(entry.closeTimer);
  try {
    await graph(`${entry.base}/closeSession`, { method: "POST", headers: { "workbook-session-id": entry.id }, maxRetries: 0 });
    entry.closed = true;
    excelSessionPool.delete(entry.key);
    return true;
  } catch (error) {
    if (!warningOnly) throw error;
    toolCallContext.getStore()?.localWarnings?.push({ operation: "Graph Excel session close", error: safeToolErrorMessage(error) });
    return false;
  }
}

function scheduleExcelSessionClose(entry) {
  clearTimeout(entry.closeTimer);
  entry.closeTimer = setTimeout(() => closeExcelManagedSession(entry, true).catch(() => null), 2 * 60 * 1000);
  entry.closeTimer.unref?.();
}

async function acquireExcelManagedSession(rawItem, resolvedBackend, persistent = true) {
  const base = excelWorkbookBase(rawItem, resolvedBackend);
  const key = await excelSessionKey(rawItem, resolvedBackend, persistent);
  const existing = excelSessionPool.get(key);
  if (existing && !existing.closed && existing.expiresAt > Date.now()) {
    try {
      await graph(`${base}/refreshSession`, { method: "POST", headers: { "workbook-session-id": existing.id }, maxRetries: 1 });
      existing.expiresAt = Date.now() + 5 * 60 * 1000;
      existing.lastUsedAt = Date.now();
      existing.reused = true;
      scheduleExcelSessionClose(existing);
      return existing;
    } catch {
      await closeExcelManagedSession(existing, true);
    }
  }
  const created = await createExcelWorkbookSession(base, persistent);
  if (!created?.id) throw new Error("Microsoft Graph did not return an Excel workbook session ID.");
  const entry = { key, id: created.id, base, persistent, createdAt: Date.now(), lastUsedAt: Date.now(), expiresAt: Date.now() + 5 * 60 * 1000, reused: false, closed: false, closeTimer: null };
  excelSessionPool.set(key, entry);
  scheduleExcelSessionClose(entry);
  return entry;
}

async function closeAllExcelSessions() {
  await Promise.allSettled([...excelSessionPool.values()].map((entry) => closeExcelManagedSession(entry, true)));
}

async function officeBatchUpdate(args = {}, kind, toolName) {
  const rawItem = await getRawInfo(args);
  if (rawItem.folder) throw new Error(`Item is a folder, not an Office document: ${rawItem.name}`);
  assertExpectedItem(rawItem, args, `${kind} batch update`);
  const detectedKind = officePackageKindFromName(rawItem.name);
  if (detectedKind !== kind) throw new Error(`Expected a ${kind} Open XML package. Got ${rawItem.name}.`);
  if (Number(rawItem.size || 0) > maxOfficePackageBytes) {
    throw new Error(`Office package is ${rawItem.size} bytes, above the ${maxOfficePackageBytes}-byte editing limit.`);
  }
  const backendResolution = kind === "excel" ? await resolveExcelOfficeBackend(rawItem, args) : { backend: "openxml" };
  const selectedBackend = backendResolution.backend;
  if (args.dryRun === false && args.confirmed !== true) {
    return { dryRun: false, confirmed: false, requiredToUpdate: "Set dryRun: false and confirmed: true after reviewing the Office edit preview." };
  }
  if (args.dryRun === false && !hasExpectedIdentity(args)) {
    return { dryRun: false, confirmed: true, requiredToUpdate: "Provide expectedName or expectedId for live Office edits." };
  }
  if (args.dryRun === false && args.expectedETag && args.expectedETag !== rawItem.eTag) {
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(rawItem),
      revisionConflict: true,
      requiredToUpdate: "The explicit expectedETag does not match the current remote item. Re-read the item and run a new preview from the latest revision."
    };
  }

  const transactionRoot = join(officeEditingRoot, `edit-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const sourcePath = join(transactionRoot, `source-${basename(rawItem.name)}`);
  const editedPath = join(transactionRoot, `edited-${basename(rawItem.name)}`);
  let editResult;
  let backup = null;
  try {
    await downloadResolvedItem(simplifyItem(rawItem), { localPath: sourcePath, overwrite: false });
    let resolvedOperations = args.operations;
    let anchorResolutions = [];
    if (args.operations.some((operation) => operation.anchor)) {
      const currentInspection = addSemanticAnchors(kind, await runOfficeHelper({
        action: "inspect",
        inputPath: sourcePath,
        kind,
        maxParagraphs: 10000,
        maxCells: 50000,
        maxSlides: 5000
      }));
      const resolution = resolveSemanticOperations(kind, currentInspection, args.operations, "unique");
      if (resolution.conflicts.length) {
        return {
          dryRun: true,
          confirmed: false,
          item: simplifyItem(rawItem),
          backend: selectedBackend,
          anchorConflict: true,
          conflicts: resolution.conflicts,
          requiredToUpdate: "Re-read the document and choose a unique semantic anchor before previewing this edit again."
        };
      }
      resolvedOperations = resolution.operations;
      anchorResolutions = resolution.resolutions;
    }
    const helperOperations = kind === "excel" && selectedBackend === "graph"
      ? resolvedOperations.filter((operation) => !excelGraphOnlyOperationTypes.has(operation.type))
      : resolvedOperations;
    const graphOnlyChanges = kind === "excel" && selectedBackend === "graph"
      ? resolvedOperations.filter((operation) => excelGraphOnlyOperationTypes.has(operation.type)).map((operation) => ({
          operation: operation.type,
          sheet: operation.sheet,
          table: operation.table,
          chart: operation.chart,
          after: Object.fromEntries(Object.entries(operation).filter(([key]) => key !== "type"))
        }))
      : [];
    if (helperOperations.length) {
      editResult = await runOfficeHelper({
        action: "edit",
        inputPath: sourcePath,
        outputPath: editedPath,
        kind,
        operations: helperOperations,
        ...(kind === "word" ? { trackedChanges: args.trackedChanges || "refuse" } : {}),
        allowMacros: args.allowMacros === true,
        allowSignedPackage: args.allowSignedPackage === true
      }, { timeoutMs: 120_000 });
    } else {
      editResult = { kind, changes: [], changeCount: 0, validation: await runOfficeHelper({ action: "validate", inputPath: sourcePath, kind }) };
    }
    if (graphOnlyChanges.length) {
      editResult.changes = [...(editResult.changes || []), ...graphOnlyChanges];
      editResult.changeCount = (editResult.changeCount || 0) + graphOnlyChanges.length;
    }
    if (!editResult.changeCount) {
      return {
        dryRun: args.dryRun !== false,
        confirmed: args.confirmed === true,
        item: simplifyItem(rawItem),
        backend: "openxml",
        changes: [],
        changeCount: 0,
        semanticDiff: officeSemanticDiff(kind, []),
        noChanges: true,
        requiredToUpdate: "No requested edit matched the document. Re-read the document and correct the operation anchors."
      };
    }
    const proof = officeEditPreviewProof(rawItem, kind, { requested: args.operations, resolved: resolvedOperations, anchorResolutions }, editResult, args, selectedBackend);
    const preview = {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      item: simplifyItem(rawItem),
      backend: selectedBackend,
      previewBackend: selectedBackend === "graph" ? "openxml" : undefined,
      changes: editResult.changes,
      changeCount: editResult.changeCount,
      semanticDiff: officeSemanticDiff(kind, editResult.changes),
      anchorResolutions,
      validation: editResult.validation
    };
    if (args.dryRun !== false) return previewWithToken(preview, toolName, proof);
    const tokenRequired = previewTokenRequiredResult(preview, toolName, proof, args.previewToken, "requiredToUpdate");
    if (tokenRequired) return tokenRequired;

    if (args.createBackup !== false) {
      const sourceValidation = await runOfficeHelper({ action: "validate", inputPath: sourcePath, kind });
      backup = await persistOfficeBackup(sourcePath, rawItem, kind, sourceValidation.package?.fingerprint, "edit");
    }

    const uploaded = selectedBackend === "graph"
      ? { item: simplifyItem(rawItem), ...(await applyExcelGraphOperations(rawItem, backendResolution, resolvedOperations)), uploadMode: "graph-workbook-session" }
      : await upload({
        localPath: editedPath,
        itemId: rawItem.id,
        remotePath: itemRemotePath(rawItem),
        conflictBehavior: "replace",
        ifMatch: rawItem.eTag,
        auditTool: toolName,
        skipAudit: true
      });
    let afterRaw = uploaded.item ? { ...rawItem, ...uploaded.item } : rawItem;
    let remoteValidation = null;
    let verificationIncomplete = selectedBackend === "graph"
      && (uploaded.semanticVerification || []).some((entry) => entry.succeeded !== true);
    if (verificationIncomplete) {
      toolCallContext.getStore()?.localWarnings?.push({
        operation: "Graph Excel semantic verification",
        error: "At least one requested Graph Excel effect could not be fully observed after the write; inspect uploaded.semanticVerification limitations."
      });
    }
    try {
      afterRaw = await getRawInfo({ itemId: rawItem.id });
    } catch (error) {
      verificationIncomplete = true;
      toolCallContext.getStore()?.localWarnings?.push({ operation: "Office post-commit metadata verification", error: safeToolErrorMessage(error) });
    }
    if (args.verify !== false) {
      try {
        const verifyPath = join(transactionRoot, `verify-${basename(rawItem.name)}`);
        await downloadResolvedItem(simplifyItem(afterRaw), { localPath: verifyPath, overwrite: false });
        remoteValidation = await runOfficeHelper({ action: "validate", inputPath: verifyPath, kind });
        if (selectedBackend === "openxml" && remoteValidation.package?.fingerprint !== editResult.validation?.package?.fingerprint) {
          verificationIncomplete = true;
          toolCallContext.getStore()?.localWarnings?.push({ operation: "Office post-commit fingerprint verification", error: "Uploaded package fingerprint differs from the locally validated edit." });
        }
      } catch (error) {
        verificationIncomplete = true;
        toolCallContext.getStore()?.localWarnings?.push({ operation: "Office post-commit validation", error: safeToolErrorMessage(error) });
      }
    }
    await writeMutationAudit(toolName, {
      status: verificationIncomplete ? "success-verification-incomplete" : "success",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem),
      after: itemAuditSummary(afterRaw),
      officeEdit: { kind, backend: selectedBackend, changeCount: editResult.changeCount, operationTypes: args.operations.map((entry) => entry.type) },
      backupCreated: Boolean(backup),
      verificationIncomplete
    });
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(afterRaw),
      backend: selectedBackend,
      changes: editResult.changes,
      changeCount: editResult.changeCount,
      semanticDiff: officeSemanticDiff(kind, editResult.changes),
      validation: editResult.validation,
      remoteValidation,
      verificationIncomplete,
      backup,
      uploaded
    };
  } catch (error) {
    if (args.dryRun === false && args.confirmed === true) {
      await writeMutationAudit(toolName, {
        status: "failed",
        target: itemAuditSummary(rawItem),
        officeEdit: { kind, backend: selectedBackend, operationTypes: (args.operations || []).map((entry) => entry.type) },
        backupCreated: Boolean(backup),
        error: safeErrorInfo(error)
      });
    }
    throw error;
  } finally {
    await rm(transactionRoot, { recursive: true, force: true });
  }
}

function officeToolForKind(kind) {
  return `onedrive_${kind === "powerpoint" ? "powerpoint" : kind}_batch_update`;
}

async function officeBatchTransform(args = {}) {
  const toolName = "onedrive_office_batch_transform";
  const requested = args.items || [];
  if (args.dryRun === false && args.confirmed !== true) {
    return { dryRun: false, confirmed: false, requiredToUpdate: "Set dryRun: false and confirmed: true after reviewing the complete cross-file Office preview." };
  }
  if (args.dryRun === false) {
    const missingIdentity = requested.findIndex((item) => !hasExpectedIdentity(item));
    if (missingIdentity >= 0) return { dryRun: false, confirmed: true, requiredToUpdate: `Item ${missingIdentity} requires expectedName or expectedId.` };
  }

  const preflight = [];
  for (const [index, item] of requested.entries()) {
    const individualTool = officeToolForKind(item.kind);
    try {
      const preview = await officeBatchUpdate({
        ...item,
        dryRun: true,
        confirmed: false,
        createBackup: args.createBackup !== false,
        verify: args.verify !== false
      }, item.kind, individualTool);
      if (preview.noChanges) throw new Error(preview.requiredToUpdate || "No requested edits matched.");
      preflight.push({ index, request: item, individualTool, preview });
    } catch (error) {
      return {
        dryRun: args.dryRun !== false,
        confirmed: args.confirmed === true,
        preflightComplete: false,
        mutationStarted: false,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        items: preflight.map((entry) => entry.preview)
      };
    }
  }

  const proof = {
    items: preflight.map((entry) => officeEditPreviewProof(
      entry.preview.item,
      entry.request.kind,
      entry.request.operations,
      entry.preview,
      {
        ...entry.request,
        createBackup: args.createBackup !== false,
        verify: args.verify !== false
      },
      entry.preview.backend
    ))
  };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    preflightComplete: true,
    mutationStarted: false,
    itemCount: preflight.length,
    totalChangeCount: preflight.reduce((sum, entry) => sum + Number(entry.preview.changeCount || 0), 0),
    items: preflight.map((entry) => ({ index: entry.index, kind: entry.request.kind, ...entry.preview }))
  };
  if (args.dryRun !== false) return previewWithToken(preview, toolName, proof);
  const tokenRequired = previewTokenRequiredResult(preview, toolName, proof, args.previewToken, "requiredToUpdate");
  if (tokenRequired) return tokenRequired;

  const completed = [];
  for (const entry of preflight) {
    try {
      const result = await officeBatchUpdate({
        ...entry.request,
        dryRun: false,
        confirmed: true,
        previewToken: entry.preview.previewToken,
        createBackup: args.createBackup !== false,
        verify: args.verify !== false
      }, entry.request.kind, entry.individualTool);
      const requiredField = Object.entries(result || {}).find(([key, value]) => key.startsWith("requiredTo") && Boolean(value));
      const committed = Boolean(
        result
        && result.dryRun === false
        && result.confirmed === true
        && result.previewTokenRequired !== true
        && result.noChanges !== true
        && !requiredField
        && Number(result.changeCount) > 0
        && result.uploaded
        && result.item?.id
        && result.item.id === entry.preview.item?.id
      );
      if (!committed) {
        const reason = result?.previewTokenRequired === true
          ? "The item changed after the cross-file preflight, so its individual preview token was refused."
          : requiredField
            ? String(requiredField[1])
            : result?.noChanges === true
              ? String(result.requiredToUpdate || "The individual update reported no changes.")
              : "The individual Office update did not return proof of a committed mutation.";
        const remaining = preflight.slice(entry.index + 1).map((candidate) => ({ index: candidate.index, item: candidate.preview.item, kind: candidate.request.kind }));
        const partialState = completed.length > 0;
        await writeMutationAudit(toolName, {
          status: "refused",
          officeBatch: { itemCount: preflight.length, completedCount: completed.length, failedIndex: entry.index, mutationStarted: partialState },
          error: { name: "OfficeBatchCommitRefused", message: reason },
          refusal: sanitizeAuditValue(result)
        });
        return {
          dryRun: false,
          confirmed: true,
          preflightComplete: true,
          mutationStarted: partialState,
          partialState,
          completed,
          failed: {
            index: entry.index,
            item: entry.preview.item,
            kind: entry.request.kind,
            refused: true,
            reason,
            result: sanitizeAuditValue(result)
          },
          remaining,
          recovery: completed.map((candidate) => ({ index: candidate.index, item: candidate.result.item, backup: candidate.result.backup }))
        };
      }
      completed.push({ index: entry.index, kind: entry.request.kind, result });
    } catch (error) {
      const remaining = preflight.slice(entry.index + 1).map((candidate) => ({ index: candidate.index, item: candidate.preview.item, kind: candidate.request.kind }));
      await writeMutationAudit(toolName, {
        status: "failed",
        officeBatch: { itemCount: preflight.length, completedCount: completed.length, failedIndex: entry.index },
        error: safeErrorInfo(error)
      });
      return {
        dryRun: false,
        confirmed: true,
        preflightComplete: true,
        mutationStarted: true,
        partialState: completed.length > 0,
        completed,
        failed: { index: entry.index, item: entry.preview.item, kind: entry.request.kind, error: safeToolErrorMessage(error) },
        remaining,
        recovery: completed.map((candidate) => ({ index: candidate.index, item: candidate.result.item, backup: candidate.result.backup }))
      };
    }
  }
  const verificationIncomplete = completed.some((entry) => entry.result?.verificationIncomplete === true);
  await writeMutationAudit(toolName, {
    status: verificationIncomplete ? "success-verification-incomplete" : "success",
    officeBatch: { itemCount: preflight.length, completedCount: completed.length, totalChangeCount: preview.totalChangeCount },
    verificationIncomplete
  });
  return {
    dryRun: false,
    confirmed: true,
    preflightComplete: true,
    mutationStarted: true,
    partialState: false,
    verificationIncomplete,
    itemCount: completed.length,
    totalChangeCount: preview.totalChangeCount,
    completed
  };
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
  return await downloadResolvedItem(info, {
    ...args,
    localPath: args.localPath || join(downloadRoot, kindName, info.name || `${kindName}-download`),
    overwrite: args.overwrite,
    _allowAlternateLocalPath: !args.localPath
  });
}

function exportFileName(name = "document", extension = ".pdf") {
  const base = basename(name, extname(name)) || "document";
  return `${base}${extension}`;
}

async function downloadExport(args = {}, formatName) {
  const exportFormats = {
    pdf: { graphFormat: "pdf", extension: ".pdf", label: "PDF" },
    text: { graphFormat: "text", extension: ".txt", label: "plain text" }
  };
  const format = exportFormats[formatName];
  if (!format) throw new Error(`Unknown export format: ${formatName}`);
  if (args.localPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.localPath), "Export", args);
  const info = await getInfo(args);
  if (info.folder) throw new Error(`Item is a folder, not an exportable document: ${info.name}`);
  const preferredTarget = args.localPath
    ? resolve(args.localPath)
    : join(downloadRoot, "export", exportFileName(info.name, format.extension));
  await assertNotLocalOneDriveSyncPathForWrite(preferredTarget, "Export", args);
  const reservation = await reserveLocalDestination(preferredTarget, {
    overwrite: args.overwrite === true,
    allowAlternate: !args.localPath
  });
  const target = reservation.path;
  const params = new URLSearchParams();
  params.set("format", format.graphFormat);
  try {
    const downloaded = await graphDownloadToFile(`${contentPath(args)}?${params.toString()}`, target, { reserved: reservation.reserved });
    return {
      item: info,
      localPath: target,
      bytesWritten: downloaded.bytesWritten,
      exportFormat: formatName,
      note: `Exported using Microsoft Graph format=${format.graphFormat}. Some file types may not support ${format.label} conversion.`
    };
  } catch (error) {
    if (reservation.reserved) await rm(target, { force: true });
    throw error;
  }
}

function truncateUtf8(text, maxBytes) {
  const buffer = Buffer.from(String(text), "utf8");
  if (buffer.length <= maxBytes) return { text: String(text), truncated: false, bytes: buffer.length };
  return {
    text: buffer.subarray(0, maxBytes).toString("utf8").replace(/\uFFFD$/u, ""),
    truncated: true,
    bytes: buffer.length
  };
}

async function preview(args = {}) {
  const info = args._resolvedInfo || await getInfo(args);
  const maxBytes = clampInteger(args.maxBytes, 65536, 1, 1048576);
  if (info.folder) return { item: info, preview: null, note: "Item is a folder; preview is only available for files." };

  if (args.preferExportText !== false && !isLikelyTextItem(info, args)) {
    try {
      const params = new URLSearchParams();
      params.set("format", "text");
      const exported = await graphLimitedBuffer(`${contentPath(args)}?${params.toString()}`, maxBytes);
      const previewText = exported.buffer.toString("utf8").replace(/\uFFFD$/u, "");
      return { item: info, preview: previewText, bytes: Buffer.byteLength(previewText, "utf8"), bytesRead: exported.bytesRead, truncated: exported.truncated, source: "graph-text-export" };
    } catch (error) {
      return { item: info, preview: null, source: "metadata", exportError: safeToolErrorMessage(error), note: "Graph text export was not available for this file." };
    }
  }

  assertTextReadable(info, args);
  const limited = await graphLimitedBuffer(contentPath(args), maxBytes);
  if (args.force !== true) assertNoBinaryNulls(limited.buffer);
  const previewText = limited.buffer.toString("utf8").replace(/\uFFFD$/u, "");
  return { item: info, preview: previewText, bytes: Buffer.byteLength(previewText, "utf8"), bytesRead: limited.bytesRead, truncated: limited.truncated, source: "text-read" };
}

function absoluteWebUrl(value) {
  const url = String(value || "").trim();
  return /^https?:\/\//iu.test(url) ? url : "";
}

function scheduleChatgptCacheRevalidation(query, cache) {
  const scopeKey = storageScopeKey(cache?.scope);
  const normalizedQuery = String(query || "").trim();
  if (!scopeKey || !normalizedQuery) return false;
  const queryKey = createHash("sha256").update(normalizedQuery.toLowerCase()).digest("hex").slice(0, 16);
  const key = `${scopeKey}:${queryKey}`;
  const now = Date.now();
  if (chatgptRevalidations.has(key)) return false;
  if (now - (chatgptRevalidationLastStartedAt.get(key) || 0) < chatgptRevalidationCooldownMs) return false;
  for (const [candidateKey, startedAt] of chatgptRevalidationLastStartedAt) {
    if (now - startedAt >= chatgptRevalidationCooldownMs) chatgptRevalidationLastStartedAt.delete(candidateKey);
  }
  while (chatgptRevalidationLastStartedAt.size >= 256) {
    chatgptRevalidationLastStartedAt.delete(chatgptRevalidationLastStartedAt.keys().next().value);
  }
  chatgptRevalidationLastStartedAt.set(key, now);
  const generation = storageScopeGeneration;
  const parentContext = toolCallContext.getStore();
  const backgroundContext = parentContext ? {
    ...parentContext,
    toolName: "chatgpt_cache_revalidate",
    localWarnings: [],
    lastGraphRequestId: null,
    lastMutationGraphRequestId: null,
    metadataCacheWrites: 0
  } : null;
  const promise = new Promise((resolvePromise) => setImmediate(resolvePromise))
    .then(async () => {
      if (generation !== storageScopeGeneration) throw accountContextChangedError("ChatGPT cache revalidation");
      const run = async () => {
        const current = await loadMetadataCache();
        if (!storageScopesEqual(current.scope, cache.scope)) throw accountContextChangedError("ChatGPT cache revalidation");
        const settings = pluginSettings();
        if (settings.deltaSyncEnabled && (current.deltaNextLink || current.deltaLink)) {
          const refreshed = await cacheRefresh({ mode: "delta", maxItems: 2000, maxPages: 5, pageSize: 200 });
          return { strategy: "delta", count: refreshed.result?.count || 0 };
        }
        const refreshed = await find({
          query: normalizedQuery,
          maxResults: 10,
          maxResultsLimit: 10,
          maxSearchTerms: 2,
          searchConcurrency: 2,
          initialSearchTermCount: 2,
          searchPageSize: 10,
          searchMaxItemsPerTerm: 10,
          minConfidenceForSearchOnly: 60,
          preferFreshLocalResults: false,
          preferStaleLocalResults: false,
          scanFallback: false,
          confirmCacheCandidates: false,
          useCache: true,
          useContentIndex: false,
          format: "compact"
        });
        return { strategy: "search", count: refreshed.items?.length || 0 };
      };
      return backgroundContext ? await toolCallContext.run(backgroundContext, run) : await run();
    })
    .then((result) => {
      console.error(JSON.stringify({
        event: "onedrive-chatgpt-cache-revalidated",
        strategy: result.strategy,
        count: result.count
      }));
      return result;
    })
    .catch((error) => {
      if (!isAccountContextChangedError(error)) {
        console.error(JSON.stringify({
          event: "onedrive-chatgpt-cache-revalidation-error",
          error: safeToolErrorMessage(error)
        }));
      }
      return null;
    })
    .finally(() => chatgptRevalidations.delete(key));
  chatgptRevalidations.set(key, promise);
  return true;
}

async function chatgptSearch(args = {}) {
  const startedAt = Date.now();
  const cache = await loadMetadataCache();
  const found = await find({
    query: String(args.query || "").trim(),
    maxResults: 10,
    maxResultsLimit: 10,
    maxSearchTerms: 6,
    searchConcurrency: 3,
    initialSearchTermCount: 3,
    searchPageSize: 10,
    searchMaxItemsPerTerm: 10,
    minConfidenceForSearchOnly: 74,
    preferFreshLocalResults: true,
    freshLocalMinConfidence: 60,
    preferStaleLocalResults: true,
    staleLocalMinConfidence: 75,
    staleLocalMaxAgeSeconds: chatgptStaleCacheMaxAgeSeconds,
    scanFallback: false,
    confirmCacheCandidates: false,
    useCache: true,
    useContentIndex: true,
    contentMaxQueries: 6,
    contentMaxResults: 10,
    format: "compact"
  });
  const revalidationScheduled = found.summary?.usedStaleLocalFastPath === true
    ? scheduleChatgptCacheRevalidation(args.query, cache)
    : false;
  const results = (found.items || []).slice(0, 10).map((item) => ({
    id: String(item.id || ""),
    title: String(item.name || item.remotePath || item.id || "OneDrive item"),
    url: absoluteWebUrl(item.webUrl)
  })).filter((item) => item.id);
  if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
    console.error(JSON.stringify({
      event: "onedrive-chatgpt-search",
      durationMs: elapsedMs(startedAt),
      results: results.length,
      graphSearchCalls: found.summary?.graphSearchCalls || 0,
      searchTermsExecuted: found.summary?.searchTermsExecuted || 0,
      bestScore: found.summary?.bestScore || 0,
      cacheFresh: Boolean(found.summary?.cacheFresh),
      cacheAgeSeconds: found.summary?.cacheAgeSeconds ?? null,
      cacheCandidates: found.summary?.cacheCandidates || 0,
      metadataCacheWrites: found.summary?.metadataCacheWrites || 0,
      usedFreshLocalFastPath: Boolean(found.summary?.usedFreshLocalFastPath),
      usedStaleLocalFastPath: Boolean(found.summary?.usedStaleLocalFastPath),
      revalidationScheduled,
      usedScanFallback: Boolean(found.summary?.usedScanFallback)
    }));
  }
  return { results };
}

function chatgptOfficeText(document, fileName = "Office document") {
  const lines = [`Document: ${fileName}`, `Format: ${document.kind || "office"}`];
  const addTableRows = (rows = [], prefix = "") => {
    for (const [rowIndex, row] of rows.entries()) {
      lines.push(`${prefix}row ${rowIndex + 1}\t${row.map((value) => String(value ?? "").replace(/[\r\n\t]+/gu, " ")).join("\t")}`);
    }
  };
  if (document.kind === "excel") {
    lines.push(`Worksheets: ${(document.sheets || []).map((sheet) => sheet.name).join(", ") || "none"}`);
    lines.push(`Extracted cells: ${document.cellCount || 0}`);
    for (const sheet of document.sheets || []) {
      lines.push("", `## Worksheet: ${sheet.name}${sheet.state && sheet.state !== "visible" ? ` (${sheet.state})` : ""}`);
      for (const cell of sheet.cells || []) {
        const value = cell.value === null || cell.value === undefined ? "" : String(cell.value).replace(/[\r\n\t]+/gu, " ");
        const formula = cell.formula ? `\tformula=${String(cell.formula).replace(/[\r\n\t]+/gu, " ")}` : "";
        lines.push(`${cell.address || "?"}\t${value}${formula}`);
      }
      for (const table of sheet.tables || []) {
        lines.push(`Table: ${table.name || table.displayName || "unnamed"}${table.ref ? ` (${table.ref})` : ""}`);
      }
      for (const chart of sheet.charts || []) {
        lines.push(`Chart: ${chart.title || chart.name || chart.type || "unnamed"}${chart.type ? ` [${chart.type}]` : ""}`);
      }
      for (const pivot of sheet.pivots || []) {
        lines.push(`Pivot table: ${pivot.name || "unnamed"}${pivot.location?.ref ? ` (${pivot.location.ref})` : ""}`);
      }
    }
    if ((document.definedNames || []).length) {
      lines.push("", "## Defined names");
      for (const entry of document.definedNames) lines.push(`${entry.name || "unnamed"}\t${entry.value || ""}`);
    }
  } else if (document.kind === "word") {
    lines.push(`Paragraphs: ${document.paragraphCount || 0}; tables: ${document.tableCount || 0}`);
    for (const paragraph of document.paragraphs || []) {
      const value = String(paragraph.text || "").trim();
      if (value) lines.push(`${paragraph.style ? `[${paragraph.style}] ` : ""}${value}`);
    }
    for (const table of document.tables || []) {
      lines.push("", `## Table ${Number(table.index || 0) + 1}`);
      addTableRows(table.rows || []);
    }
    for (const comment of document.comments || []) {
      lines.push(`Comment${comment.author ? ` by ${comment.author}` : ""}: ${comment.text || ""}`);
    }
  } else if (document.kind === "powerpoint") {
    lines.push(`Slides: ${document.slideCount || 0}`);
    const addShape = (shape, indent = "") => {
      const value = String(shape.text || "").trim();
      if (value) lines.push(`${indent}${shape.name ? `${shape.name}: ` : ""}${value}`);
      if (shape.table?.rows?.length) addTableRows(shape.table.rows, indent);
      for (const child of shape.children || []) addShape(child, `${indent}  `);
    };
    for (const slide of document.slides || []) {
      lines.push("", `## Slide ${Number(slide.index || 0) + 1}`);
      for (const shape of slide.shapes || []) addShape(shape);
      if (slide.notes) lines.push(`Speaker notes: ${slide.notes}`);
    }
  }
  const limited = truncateUtf8(lines.join("\n"), chatgptFetchTextByteLimit);
  return { ...limited, truncated: limited.truncated || Boolean(document.truncated) };
}

function chatgptIndexedText(indexed, fileName = "OneDrive item") {
  if (!indexed?.structuredKind || !Array.isArray(indexed.segments) || !indexed.segments.length) {
    return String(indexed?.text || "");
  }
  const lines = [`Document: ${fileName}`, `Format: ${indexed.structuredKind}`];
  let section = null;
  for (const segment of indexed.segments) {
    const anchor = segment.anchor || {};
    const value = String(segment.text ?? "").replace(/[\r\n\t]+/gu, " ").trim();
    if (!value) continue;
    if (indexed.structuredKind === "excel") {
      const sheet = anchor.sheet || "Unknown";
      if (sheet !== section) {
        section = sheet;
        lines.push("", `## Worksheet: ${sheet}`);
      }
      if (anchor.type === "formula") continue;
      lines.push(`${anchor.address || "?"}\t${value}${anchor.formula ? `\tformula=${String(anchor.formula).replace(/[\r\n\t]+/gu, " ")}` : ""}`);
    } else if (indexed.structuredKind === "powerpoint") {
      const slide = Number(anchor.slideIndex || 0) + 1;
      if (slide !== section) {
        section = slide;
        lines.push("", `## Slide ${slide}`);
      }
      lines.push(value);
    } else if (indexed.structuredKind === "word") {
      if (anchor.type === "tableCell") {
        lines.push(`Table ${Number(anchor.tableIndex || 0) + 1}, row ${Number(anchor.rowIndex || 0) + 1}, column ${Number(anchor.columnIndex || 0) + 1}: ${value}`);
      } else {
        lines.push(value);
      }
    }
  }
  return lines.join("\n");
}

function chatgptFetchFingerprint(item = {}) {
  return createHash("sha256").update([
    String(item.id || ""),
    String(item.cTag || ""),
    String(item.eTag || ""),
    String(item.lastModifiedDateTime || ""),
    String(item.size ?? "")
  ].join("\0")).digest("hex").slice(0, 20);
}

function encodeChatgptFetchContinuation({ itemId, fingerprint, offset, part }) {
  const payload = Buffer.from(JSON.stringify({ v: 1, i: itemId, f: fingerprint, o: offset, p: part }), "utf8").toString("base64url");
  return `onedrive-fetch-chunk:${payload}`;
}

function decodeChatgptFetchContinuation(value) {
  const id = String(value || "");
  if (!id.startsWith("onedrive-fetch-chunk:")) return null;
  if (id.length > 4096) throw new Error("The OneDrive fetch continuation ID is too long.");
  try {
    const parsed = JSON.parse(Buffer.from(id.slice("onedrive-fetch-chunk:".length), "base64url").toString("utf8"));
    if (parsed?.v !== 1
      || typeof parsed.i !== "string" || !parsed.i || parsed.i.length > 1024
      || typeof parsed.f !== "string" || !/^[a-f0-9]{20}$/u.test(parsed.f)
      || !Number.isInteger(parsed.o) || parsed.o < 0 || parsed.o > chatgptFetchTextByteLimit
      || !Number.isInteger(parsed.p) || parsed.p < 1 || parsed.p > 100) {
      throw new Error("invalid fields");
    }
    return { itemId: parsed.i, fingerprint: parsed.f, offset: parsed.o, part: parsed.p };
  } catch {
    throw new Error("The OneDrive fetch continuation ID is invalid or expired. Fetch the original item ID again.");
  }
}

function utf8ByteSlice(text, offset = 0, maxBytes = chatgptFetchChunkByteLimit) {
  const buffer = Buffer.from(String(text || ""), "utf8");
  let start = Math.min(Math.max(0, offset), buffer.length);
  while (start < buffer.length && (buffer[start] & 0xc0) === 0x80) start += 1;
  let end = Math.min(buffer.length, start + maxBytes);
  while (end < buffer.length && (buffer[end] & 0xc0) === 0x80) end -= 1;
  if (end <= start && start < buffer.length) end = Math.min(buffer.length, start + maxBytes);
  return {
    text: buffer.subarray(start, end).toString("utf8").replace(/^\uFFFD|\uFFFD$/gu, ""),
    start,
    end,
    totalBytes: buffer.length,
    hasMore: end < buffer.length
  };
}

function evenlySampleChatgptSections(sections, maxSections = 24) {
  if (sections.length <= maxSections) return sections;
  const indexes = new Set();
  for (let index = 0; index < maxSections; index += 1) {
    indexes.add(Math.round(index * (sections.length - 1) / (maxSections - 1)));
  }
  return [...indexes].sort((left, right) => left - right).map((index) => sections[index]);
}

function compactChatgptProgressiveText(text, maxBytes = chatgptInitialFetchTextByteLimit) {
  const fullText = String(text || "");
  const fullBytes = Buffer.byteLength(fullText, "utf8");
  if (fullBytes <= maxBytes) return { text: fullText, progressive: false, fullBytes, sampledSections: 0, totalSections: 0 };
  const notice = "Progressive preview: this is a compact cross-section of a larger file. Fetch metadata.nextChunkId only if complete detail is needed.";
  const lines = fullText.split("\n");
  const header = [];
  const sections = [];
  let current = null;
  for (const line of lines) {
    if (/^##\s+/u.test(line)) {
      current = { heading: line, lines: [] };
      sections.push(current);
    } else if (current) {
      current.lines.push(line);
    } else {
      header.push(line);
    }
  }
  if (!sections.length) {
    const noticeBytes = Buffer.byteLength(`${notice}\n\n`, "utf8");
    const available = Math.max(1024, maxBytes - noticeBytes - 80);
    const head = utf8ByteSlice(fullText, 0, Math.floor(available * 0.7));
    const tailStart = Math.max(head.end, fullBytes - Math.ceil(available * 0.3));
    const tail = utf8ByteSlice(fullText, tailStart, fullBytes - tailStart);
    const previewText = `${notice}\n\n${head.text}\n\n[… middle omitted from compact preview …]\n\n${tail.text}`;
    return { text: truncateUtf8(previewText, maxBytes).text, progressive: true, fullBytes, sampledSections: 0, totalSections: 0 };
  }
  const selected = evenlySampleChatgptSections(sections);
  const headerText = truncateUtf8(header.join("\n"), 4096).text;
  const fixedBytes = Buffer.byteLength(`${notice}\n\n${headerText}\n\n`, "utf8")
    + selected.reduce((sum, section) => sum + Buffer.byteLength(`${section.heading}\n`, "utf8"), 0)
    + 256;
  const perSectionBytes = Math.max(384, Math.floor(Math.max(1024, maxBytes - fixedBytes) / selected.length));
  const parts = [notice];
  if (headerText.trim()) parts.push(headerText);
  if (selected.length < sections.length) parts.push(`Section sampling: ${selected.length} of ${sections.length} sections are represented.`);
  for (const section of selected) {
    const body = truncateUtf8(section.lines.join("\n"), perSectionBytes).text;
    parts.push(`${section.heading}\n${body}`.trim());
  }
  return {
    text: truncateUtf8(parts.join("\n\n"), maxBytes).text,
    progressive: true,
    fullBytes,
    sampledSections: selected.length,
    totalSections: sections.length
  };
}

function pruneChatgptFetchSnapshots(now = Date.now()) {
  for (const [key, entry] of chatgptFetchSnapshots) {
    if (!entry || entry.expiresAt <= now) chatgptFetchSnapshots.delete(key);
  }
  while (chatgptFetchSnapshots.size > chatgptFetchSnapshotMaxEntries) {
    chatgptFetchSnapshots.delete(chatgptFetchSnapshots.keys().next().value);
  }
}

function chatgptFetchSnapshotKey(scopeKey, itemId, fingerprint) {
  return `${scopeKey}:${fingerprint}:${itemId}`;
}

function rememberChatgptFetchSnapshot(scopeKey, snapshot) {
  if (!scopeKey || !snapshot?.item?.id) return;
  pruneChatgptFetchSnapshots();
  const fingerprint = chatgptFetchFingerprint(snapshot.item);
  const key = chatgptFetchSnapshotKey(scopeKey, snapshot.item.id, fingerprint);
  chatgptFetchSnapshots.delete(key);
  chatgptFetchSnapshots.set(key, { ...snapshot, fingerprint, expiresAt: Date.now() + chatgptFetchSnapshotTtlMs });
  pruneChatgptFetchSnapshots();
}

function rememberedChatgptFetchSnapshot(scopeKey, continuation) {
  pruneChatgptFetchSnapshots();
  const key = chatgptFetchSnapshotKey(scopeKey, continuation.itemId, continuation.fingerprint);
  const snapshot = chatgptFetchSnapshots.get(key) || null;
  if (snapshot) {
    chatgptFetchSnapshots.delete(key);
    chatgptFetchSnapshots.set(key, snapshot);
  }
  return snapshot;
}

function chatgptIndexedSnapshot(indexed, item, source) {
  const indexedText = chatgptIndexedText(indexed, item.name || item.id);
  const limited = truncateUtf8(indexedText, chatgptFetchTextByteLimit);
  return {
    item,
    preview: limited.text,
    bytes: Buffer.byteLength(limited.text, "utf8"),
    bytesRead: indexed.bytes ?? Buffer.byteLength(indexed.text || "", "utf8"),
    truncated: limited.truncated || Boolean(indexed.truncated),
    source
  };
}

async function resolveChatgptFetchSnapshot(id, cache, contentIndex) {
  const cached = cache.itemsById?.[id] || null;
  const cacheFresh = metadataCacheFresh(cache);
  const indexed = contentIndex.entriesById?.[id] || null;
  if (cacheFresh && cached?.folder) {
    return { cacheFresh, usedCachedMetadata: true, usedContentIndex: false, snapshot: { item: cached, preview: null, note: "This OneDrive item is a folder.", source: "metadata-cache", truncated: false } };
  }
  if (cacheFresh && cached && indexed && contentIndexEntryFresh(indexed, cached)) {
    return { cacheFresh, usedCachedMetadata: true, usedContentIndex: true, snapshot: chatgptIndexedSnapshot(indexed, cached, "content-index") };
  }
  const info = cacheFresh && cached ? cached : await getInfo({ itemId: id });
  if (info.folder) {
    return { cacheFresh, usedCachedMetadata: Boolean(cacheFresh && cached), usedContentIndex: false, snapshot: { item: info, preview: null, note: "This OneDrive item is a folder.", source: "metadata", truncated: false } };
  }
  if (indexed && contentIndexEntryFresh(indexed, info)) {
    return { cacheFresh, usedCachedMetadata: Boolean(cacheFresh && cached), usedContentIndex: true, snapshot: chatgptIndexedSnapshot(indexed, info, cacheFresh ? "content-index" : "content-index-validated") };
  }
  const officeKind = officePackageKindFromName(info.name);
  const commonKind = commonExtractionKind(info);
  let snapshot;
  if (officeKind) {
    try {
      const document = await inspectRemoteOfficePackage({
        itemId: id,
        _resolvedInfo: info,
        maxParagraphs: 4000,
        maxCells: 15_000,
        maxSlides: 500,
        includeCells: true,
        includeTables: true,
        includeCharts: true,
        includePivots: true,
        strictRelationships: true
      }, officeKind, "inspect");
      const rendered = chatgptOfficeText(document, info.name);
      snapshot = {
        item: info,
        preview: rendered.text,
        bytes: Buffer.byteLength(rendered.text, "utf8"),
        bytesRead: Number(info.size || 0),
        truncated: rendered.truncated,
        source: "office-openxml",
        officeKind
      };
    } catch (error) {
      snapshot = await preview({ itemId: id, _resolvedInfo: info, maxBytes: chatgptFetchTextByteLimit, preferExportText: true });
      if (!snapshot.preview) snapshot.exportError = `Open XML extraction failed: ${safeToolErrorMessage(error)}; ${snapshot.exportError || "Graph text export was unavailable."}`;
    }
  } else if (commonKind) {
    try {
      snapshot = await extractCommonDocumentText(info, { maxBytes: chatgptFetchTextByteLimit });
    } catch (error) {
      snapshot = await preview({ itemId: id, _resolvedInfo: info, maxBytes: chatgptFetchTextByteLimit, preferExportText: true });
      if (!snapshot.preview) snapshot.exportError = `Local ${commonKind} extraction failed: ${safeToolErrorMessage(error)}; ${snapshot.exportError || "Graph text export was unavailable."}`;
    }
  } else {
    snapshot = await preview({ itemId: id, _resolvedInfo: info, maxBytes: chatgptFetchTextByteLimit, preferExportText: true });
  }
  return { cacheFresh, usedCachedMetadata: Boolean(cacheFresh && cached), usedContentIndex: false, snapshot };
}

function chatgptFetchMetadata(snapshot, overrides = {}) {
  const item = snapshot.item || {};
  const metadata = {
    type: item.folder ? "folder" : "file",
    size: String(item.size ?? ""),
    modified: String(item.lastModifiedDateTime || ""),
    mimeType: String(item.file?.mimeType || ""),
    previewSource: String(snapshot.source || "metadata"),
    truncated: String(Boolean(snapshot.truncated)),
    ...Object.fromEntries(Object.entries(overrides).map(([key, value]) => [key, String(value)]))
  };
  if (snapshot.exportError) metadata.previewError = String(snapshot.exportError);
  return metadata;
}

async function chatgptFetch(args = {}) {
  const startedAt = Date.now();
  const requestedId = String(args.id || "").trim();
  const continuation = decodeChatgptFetchContinuation(requestedId);
  const [cache, contentIndex] = await Promise.all([loadMetadataCache(), loadContentIndex()]);
  const scopeKey = storageScopeKey(cache.scope) || storageScopeKey(await activeStorageScope());
  let resolved;
  let snapshot;
  let result;
  if (continuation) {
    snapshot = rememberedChatgptFetchSnapshot(scopeKey, continuation);
    if (!snapshot) {
      resolved = await resolveChatgptFetchSnapshot(continuation.itemId, cache, contentIndex);
      snapshot = resolved.snapshot;
      const currentFingerprint = chatgptFetchFingerprint(snapshot.item);
      if (currentFingerprint !== continuation.fingerprint) {
        throw new Error("The OneDrive item changed after the progressive preview. Fetch the original item ID again to read the current version.");
      }
      rememberChatgptFetchSnapshot(scopeKey, snapshot);
    }
    const fullText = String(snapshot.preview ?? snapshot.note ?? "");
    const chunk = utf8ByteSlice(fullText, continuation.offset, chatgptFetchChunkByteLimit);
    if (!chunk.text && continuation.offset > 0) throw new Error("The OneDrive fetch continuation is past the available extracted text. Fetch the original item ID again.");
    const nextChunkId = chunk.hasMore
      ? encodeChatgptFetchContinuation({ itemId: continuation.itemId, fingerprint: continuation.fingerprint, offset: chunk.end, part: continuation.part + 1 })
      : "";
    const chunkCount = Math.max(1, Math.ceil(chunk.totalBytes / chatgptFetchChunkByteLimit));
    result = {
      id: requestedId,
      title: String(snapshot.item?.name || continuation.itemId || "OneDrive item"),
      text: chunk.text,
      url: absoluteWebUrl(snapshot.item?.webUrl),
      metadata: chatgptFetchMetadata(snapshot, {
        progressive: true,
        sourceItemId: continuation.itemId,
        chunkIndex: continuation.part,
        chunkCount,
        fullTextBytes: chunk.totalBytes,
        returnedTextBytes: Buffer.byteLength(chunk.text, "utf8"),
        sourceTruncated: Boolean(snapshot.truncated),
        truncated: Boolean(snapshot.truncated || chunk.hasMore),
        ...(nextChunkId ? { nextChunkId } : {})
      })
    };
  } else {
    resolved = await resolveChatgptFetchSnapshot(requestedId, cache, contentIndex);
    snapshot = resolved.snapshot;
    const item = snapshot.item || {};
    const fullText = item.folder
      ? String(snapshot.note || "This OneDrive item is a folder.")
      : String(snapshot.preview ?? snapshot.note ?? "");
    const compact = item.folder ? { text: fullText, progressive: false, fullBytes: Buffer.byteLength(fullText, "utf8"), sampledSections: 0, totalSections: 0 }
      : compactChatgptProgressiveText(fullText);
    const fingerprint = chatgptFetchFingerprint(item);
    const nextChunkId = compact.progressive
      ? encodeChatgptFetchContinuation({ itemId: String(item.id || requestedId), fingerprint, offset: 0, part: 1 })
      : "";
    if (compact.progressive) rememberChatgptFetchSnapshot(scopeKey, snapshot);
    result = {
      id: String(item.id || requestedId || ""),
      title: String(item.name || requestedId || "OneDrive item"),
      text: compact.text,
      url: absoluteWebUrl(item.webUrl),
      metadata: chatgptFetchMetadata(snapshot, {
        progressive: compact.progressive,
        fullTextBytes: compact.fullBytes,
        returnedTextBytes: Buffer.byteLength(compact.text, "utf8"),
        sourceTruncated: Boolean(snapshot.truncated),
        truncated: Boolean(snapshot.truncated || compact.progressive),
        ...(compact.progressive ? {
          sampledSections: compact.sampledSections,
          totalSections: compact.totalSections,
          fullTextChunkCount: Math.max(1, Math.ceil(compact.fullBytes / chatgptFetchChunkByteLimit)),
          nextChunkId
        } : {})
      })
    };
  }
  if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
    console.error(JSON.stringify({
      event: "onedrive-chatgpt-fetch",
      durationMs: elapsedMs(startedAt),
      cacheFresh: resolved?.cacheFresh ?? metadataCacheFresh(cache),
      usedCachedMetadata: Boolean(resolved?.usedCachedMetadata),
      usedContentIndex: Boolean(resolved?.usedContentIndex || String(snapshot.source || "").startsWith("content-index")),
      previewSource: snapshot.source || "metadata",
      continuation: Boolean(continuation),
      progressive: result.metadata?.progressive === "true",
      bytes: Buffer.byteLength(result.text || "", "utf8"),
      truncated: result.metadata?.truncated === "true"
    }));
  }
  return result;
}

function exactFilenameKey(value) {
  return String(value || "").trim().normalize("NFKC").toLowerCase();
}

async function chatgptOpenFiles(args = {}) {
  const startedAt = Date.now();
  const names = args.names || [];
  const files = await mapWithConcurrency(names, 2, async (name) => {
    const fileStartedAt = Date.now();
    const requestedName = String(name || "").trim();
    try {
      const searched = await chatgptSearch({ query: requestedName });
      const exact = (searched.results || []).filter((candidate) => exactFilenameKey(candidate.title) === exactFilenameKey(requestedName));
      if (exact.length === 0) {
        return {
          name: requestedName,
          status: "not_found",
          candidates: (searched.results || []).slice(0, 3),
          durationMs: elapsedMs(fileStartedAt)
        };
      }
      if (exact.length > 1) {
        return {
          name: requestedName,
          status: "ambiguous",
          candidates: exact.slice(0, 3),
          durationMs: elapsedMs(fileStartedAt)
        };
      }
      const fetched = await chatgptFetch({ id: exact[0].id });
      return {
        name: requestedName,
        status: "found",
        id: fetched.id,
        title: fetched.title,
        text: fetched.text,
        url: fetched.url,
        metadata: fetched.metadata,
        durationMs: elapsedMs(fileStartedAt)
      };
    } catch (error) {
      return {
        name: requestedName,
        status: "error",
        error: safeToolErrorMessage(error),
        durationMs: elapsedMs(fileStartedAt)
      };
    }
  });
  const result = { files, durationMs: elapsedMs(startedAt) };
  if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
    console.error(JSON.stringify({
      event: "onedrive-chatgpt-open-files",
      durationMs: result.durationMs,
      requested: names.length,
      found: files.filter((file) => file.status === "found").length,
      ambiguous: files.filter((file) => file.status === "ambiguous").length,
      notFound: files.filter((file) => file.status === "not_found").length,
      errors: files.filter((file) => file.status === "error").length
    }));
  }
  return result;
}

async function chatgptPermissionSummary(itemId) {
  const current = await permissionList({ itemId }, "compact");
  const linkPermissions = current.filter((permission) => Boolean(permission.link));
  return {
    permissionCount: current.length,
    sharingLinkCount: linkPermissions.length,
    anonymousLinkCount: linkPermissions.filter((permission) => permission.link?.scope === "anonymous").length,
    roles: [...new Set(current.flatMap((permission) => permission.roles || []).map(String))].sort()
  };
}

function chatgptPreviewActionSummary(action) {
  switch (action.operation) {
    case "rename":
      return `Rename item ${action.itemId} to ${action.newName}.`;
    case "move":
      return `Move item ${action.itemId} to ${action.destinationParentItemId || action.destinationParentPath || "/"}${action.newName ? ` as ${action.newName}` : ""}.`;
    case "copy":
      return `Copy item ${action.itemId} to ${action.destinationParentItemId || action.destinationParentPath || "/"}${action.newName ? ` as ${action.newName}` : ""}.`;
    case "createSharingLink":
      return `Create a ${action.scope || "anonymous"} ${action.linkType || "view"} sharing link for item ${action.itemId}.`;
    case "revokePermission":
      return `Revoke permission ${action.permissionId} from item ${action.itemId}.`;
    default:
      return `Preview ${action.operation} for item ${action.itemId}.`;
  }
}

function validateChatgptPreviewAction(action, index) {
  if (action.operation === "rename" && !action.newName) {
    throw new Error(`Preview action ${index} rename requires newName.`);
  }
  if (["move", "copy"].includes(action.operation)
    && action.destinationParentPath !== undefined
    && action.destinationParentItemId !== undefined) {
    throw new Error(`Preview action ${index} must use only one destination parent selector.`);
  }
  if (action.operation === "revokePermission" && !action.permissionId) {
    throw new Error(`Preview action ${index} revokePermission requires permissionId.`);
  }
}

async function chatgptPreviewAction(action, index) {
  const startedAt = Date.now();
  try {
    validateChatgptPreviewAction(action, index);
    let preview;
    let accessSummary;
    let accessSummaryError;
    if (action.operation === "rename") {
      preview = await rename({ itemId: action.itemId, newName: action.newName, dryRun: true, confirmed: false });
    } else if (action.operation === "move") {
      preview = await moveItem({
        itemId: action.itemId,
        destinationParentPath: action.destinationParentPath,
        destinationParentItemId: action.destinationParentItemId,
        newName: action.newName,
        dryRun: true,
        confirmed: false
      });
    } else if (action.operation === "copy") {
      preview = await copyItem({
        itemId: action.itemId,
        destinationParentPath: action.destinationParentPath,
        destinationParentItemId: action.destinationParentItemId,
        newName: action.newName,
        waitForCompletion: false,
        dryRun: true,
        confirmed: false
      });
    } else if (action.operation === "createSharingLink") {
      const [linkPreview, summaryResult] = await Promise.all([
        createSharingLink({
          itemId: action.itemId,
          type: action.linkType || "view",
          scope: action.scope || "anonymous",
          includePermissionDiff: false,
          dryRun: true,
          confirmed: false
        }),
        chatgptPermissionSummary(action.itemId)
          .then((value) => ({ value }))
          .catch((error) => ({ error: safeToolErrorMessage(error) }))
      ]);
      preview = linkPreview;
      accessSummary = summaryResult.value;
      accessSummaryError = summaryResult.error;
    } else if (action.operation === "revokePermission") {
      preview = await revokePermission({
        itemId: action.itemId,
        permissionId: action.permissionId,
        includePermissions: false,
        dryRun: true,
        confirmed: false
      });
    } else {
      throw new Error(`Unsupported preview operation: ${action.operation}.`);
    }
    const previewToken = String(preview?.previewToken || "");
    return {
      index,
      operation: action.operation,
      itemId: action.itemId,
      expectedId: action.itemId,
      isError: false,
      dryRun: true,
      previewTokenPresent: Boolean(previewToken),
      ...(previewToken ? { previewToken } : {}),
      ...(preview?.previewTokenExpiresAt ? { previewTokenExpiresAt: String(preview.previewTokenExpiresAt) } : {}),
      summary: chatgptPreviewActionSummary(action),
      ...(accessSummary ? { accessSummary } : {}),
      ...(accessSummaryError ? { accessSummaryError } : {}),
      durationMs: elapsedMs(startedAt)
    };
  } catch (error) {
    return {
      index,
      operation: action.operation,
      itemId: action.itemId,
      isError: true,
      dryRun: true,
      previewTokenPresent: false,
      error: safeToolErrorMessage(error),
      durationMs: elapsedMs(startedAt)
    };
  }
}

async function chatgptPreviewActions(args = {}) {
  const startedAt = Date.now();
  const actions = args.actions || [];
  const results = await mapWithConcurrency(actions.map((action, index) => ({ action, index })), 3, async ({ action, index }) => (
    await chatgptPreviewAction(action, index)
  ));
  const result = { dryRun: true, results, durationMs: elapsedMs(startedAt) };
  if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
    console.error(JSON.stringify({
      event: "onedrive-chatgpt-preview-actions",
      durationMs: result.durationMs,
      actions: actions.length,
      errors: results.filter((entry) => entry.isError).length,
      sharingPreviews: actions.filter((entry) => entry.operation === "createSharingLink").length
    }));
  }
  return result;
}

function updateManifestPath(localPath, manifestPath) {
  return manifestPath ? resolve(manifestPath) : `${resolve(localPath)}.onedrive-update.json`;
}

async function updateFile(args = {}) {
  const remote = assertSafeRemotePath(args.remotePath, "remotePath");
  if (!remote) throw new Error("remotePath is required.");
  const scope = await activeStorageScope();

  if (args.mode === "checkout") {
    if (args.localPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.localPath), "Checkout", args);
    if (args.manifestPath) await assertNotLocalOneDriveSyncPathForWrite(resolve(args.manifestPath), "Checkout manifest", args);
    const info = await getInfo(args.itemId ? { itemId: args.itemId } : { path: remote });
    if (info.folder) throw new Error(`Cannot checkout a folder: ${info.name}`);
    const localPath = args.localPath ? resolve(args.localPath) : join(updateRoot, info.name || basename(remote));
    await assertNotLocalOneDriveSyncPathForWrite(localPath, "Checkout", args);
    const manifestPath = updateManifestPath(localPath, args.manifestPath);
    await assertNotLocalOneDriveSyncPathForWrite(manifestPath, "Checkout manifest", args);
    const manifestReservation = await reserveLocalDestination(manifestPath, {
      overwrite: args.overwriteManifest === true,
      allowAlternate: false
    });
    try {
      const downloaded = await download({
        ...(args.itemId ? { itemId: args.itemId } : { path: remote }),
        localPath,
        overwrite: args.overwriteLocal === true,
        allowLocalOneDriveSyncPath: args.allowLocalOneDriveSyncPath
      });
      const manifest = {
        version: 2,
        scope,
        checkedOutAt: new Date().toISOString(),
        remotePath: remote,
        item: downloaded.item,
        localPath: downloaded.localPath
      };
      await writePrivateFileAtomic(manifestPath, JSON.stringify(manifest));
      return { mode: "checkout", ...downloaded, manifestPath };
    } catch (error) {
      if (manifestReservation.reserved) await rm(manifestPath, { force: true });
      throw error;
    }
  }

  if (args.mode !== "commit") throw new Error("mode must be checkout or commit.");
  const localPath = resolve(args.localPath || join(updateRoot, basename(remote)));
  await assertNotLocalOneDriveSyncPathForRead(localPath, "Commit", args);
  const manifestPath = updateManifestPath(localPath, args.manifestPath);
  let manifest = null;
  try {
    manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  } catch (error) {
    if (args.force !== true) throw new Error(`Could not read checkout manifest ${manifestPath}. Pass force: true only when intentionally overwriting without checkout metadata.`);
  }
  if (manifest && (manifest.version !== 2 || !manifest.scope)) {
    throw new Error(`Checkout manifest ${manifestPath} is legacy and unscoped. Re-checkout the file before committing.`);
  }
  if (manifest && !storageScopesEqual(manifest.scope, scope)) {
    throw new Error(`Checkout manifest ${manifestPath} belongs to a different OneDrive authentication context or drive. Re-checkout the file.`);
  }
  if (manifest && args.force !== true) {
    const manifestProblems = [];
    if (manifest.version !== 2) manifestProblems.push("version");
    if (manifest.remotePath && cleanPath(manifest.remotePath) !== cleanPath(remote)) manifestProblems.push("remotePath");
    if (manifest.localPath && resolve(manifest.localPath) !== localPath) manifestProblems.push("localPath");
    if (manifestPath && updateManifestPath(localPath, manifestPath) !== manifestPath) manifestProblems.push("manifestPath");
    if (manifestProblems.length) {
      throw new Error(`Checkout manifest does not match this commit request (${manifestProblems.join(", ")}). Re-checkout or pass force: true if you intend to override.`);
    }
  }

  const current = await getRawInfo({ path: remote });
  if (args.conflictCheck !== false && manifest?.item && args.force !== true) {
    const checked = manifest.item;
    const changed = [
      checked.id && current.id !== checked.id ? "id" : null,
      checked.eTag && current.eTag !== checked.eTag ? "eTag" : null,
      checked.cTag && current.cTag !== checked.cTag ? "cTag" : null,
      Number.isFinite(checked.size) && current.size !== checked.size ? "size" : null,
      checked.lastModifiedDateTime && current.lastModifiedDateTime !== checked.lastModifiedDateTime ? "lastModifiedDateTime" : null
    ].filter(Boolean);
    if (changed.length) {
      throw new Error(`Remote file changed since checkout (${changed.join(", ")}). Re-checkout or pass force: true if you intend to overwrite.`);
    }
  }

  let backup = null;
  if (args.createBackup !== false) {
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const backupPath = join(backupRoot, `${stamp}-${randomUUID()}-${basename(remote)}`);
    backup = await download({ path: remote, localPath: backupPath, overwrite: false });
  }

  const uploaded = await upload({
    localPath,
    remotePath: remote,
    conflictBehavior: "replace",
    allowLocalOneDriveSyncPath: args.allowLocalOneDriveSyncPath,
    ifMatch: args.conflictCheck !== false && args.force !== true ? manifest?.item?.eTag : undefined,
    auditTool: "onedrive_update_file",
    guardedInternalReplace: true
  });
  const verified = args.verify !== false
    ? await bestEffortLocalWrite("post-commit remote verification", async () => await getInfo({ path: remote }))
    : null;
  if (manifest && args.force !== true) {
    const updatedManifest = {
      ...manifest,
      committedAt: new Date().toISOString(),
      remotePath: remote,
      localPath,
      item: verified || uploaded.item
    };
    await bestEffortLocalWrite("update manifest write", async () => await writePrivateFile(manifestPath, JSON.stringify(updatedManifest, null, 2)));
  }
  return {
    mode: "commit",
    remotePath: remote,
    localPath,
    manifestPath,
    backup,
    uploaded,
    verified,
    verificationIncomplete: args.verify !== false && !verified,
    note: "Committed local edits after checkout-manifest conflict checks."
  };
}

async function recent(args = {}) {
  const params = new URLSearchParams();
  params.set("$top", String(clampInteger(args.limit, 50, 1, 200)));
  const result = await graph(`/me/drive/recent?${params.toString()}`);
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems(result.value || []));
  return { items: (result.value || []).map((item) => formatDriveItem(item, args.format)), count: (result.value || []).length };
}

async function largeFiles(args = {}) {
  const minBytes = args.minBytes ?? 104857600;
  const matches = [];
  const result = await scan({
    ...args,
    includeFiles: true,
    includeFolders: false,
    maxResults: 1,
    format: "full",
    onItem: (item) => {
      if (item.file && (item.size || 0) >= minBytes) matches.push(simplifyItem(item));
    }
  });
  const files = matches
    .sort((left, right) => (right.size || 0) - (left.size || 0))
    .slice(0, clampInteger(args.limit, 50, 1, 200))
    .map((item) => formatSimplifiedItem(item, args.format));
  return {
    filters: { minBytes },
    scanSummary: result.summary,
    count: files.length,
    itemsReturned: files.length,
    scanItemsReturned: result.summary.returned,
    items: files,
    note: "count/itemsReturned report files matching the large-file filter; scanSummary.returned reports the bounded traversal sample."
  };
}

function duplicateKey(item) {
  const hashes = item.file?.hashes || {};
  const hash = hashes.quickXorHash || hashes.sha1Hash || hashes.sha256Hash;
  if (hash) return `hash:${hash}`;
  return `name-size:${String(item.name || "").toLowerCase()}:${item.size || 0}`;
}

async function duplicates(args = {}) {
  const groups = new Map();
  const result = await scan({
    ...args,
    includeFiles: true,
    includeFolders: false,
    maxResults: 1,
    format: "full",
    onItem: (item) => {
      if (!item.file) return;
      const simplified = simplifyItem(item);
      const key = duplicateKey(simplified);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(simplified);
    }
  });
  const duplicates = [...groups.entries()]
    .filter(([, items]) => items.length > 1)
    .map(([key, items]) => ({
      key,
      count: items.length,
      totalBytes: items.reduce((sum, item) => sum + (item.size || 0), 0),
      items: items.map((item) => formatSimplifiedItem(item, args.format))
    }))
    .sort((left, right) => right.totalBytes - left.totalBytes)
    .slice(0, clampInteger(args.limit, 50, 1, 200));
  return {
    scanSummary: result.summary,
    duplicateGroups: duplicates.length,
    groupsReturned: duplicates.length,
    scanItemsReturned: result.summary.returned,
    groups: duplicates,
    note: "duplicateGroups/groupsReturned report duplicate groups after scanning; scanSummary.returned reports the bounded traversal sample."
  };
}

async function sharingAudit(args = {}, publicOnly = false) {
  const scanResult = await scan({
    ...args,
    includeFiles: true,
    includeFolders: args.includeFolders !== false,
    maxResults: Math.min(clampInteger(args.maxItems, 1000, 1, 50000), 5000),
    format: "full"
  });
  const matches = [];
  const candidates = scanResult.items || [];
  const limit = clampInteger(args.limit, 50, 1, 200);
  let auditedCount = 0;
  let errorCount = 0;
  let matchingItemCount = 0;
  const errors = [];
  for (let offset = 0; offset < candidates.length && matches.length < limit; offset += 20) {
    const chunk = candidates.slice(offset, offset + 20);
    let batch;
    try {
      batch = await batchPermissions({
        items: chunk.map((item) => ({ itemId: item.id })),
        format: "compact"
      });
    } catch (error) {
      errorCount += chunk.length;
      if (errors.length < 10) errors.push({ error: safeToolErrorMessage(error), itemIds: chunk.map((item) => item.id) });
      continue;
    }
    for (const [index, result] of (batch.items || []).entries()) {
      const item = chunk[index];
      if (!item) continue;
      if (result.error) {
        errorCount += 1;
        if (errors.length < 10) errors.push({ itemId: item.id, status: result.status, error: result.error });
        continue;
      }
      auditedCount += 1;
      const permissions = publicOnly
        ? result.permissions.filter((permission) => permission.permissionKind === "anonymous_link")
        : result.permissions.filter((permission) => args.includeOwnerPermissions === true ? !permission.inheritedFrom : isExplicitSharingPermission(permission));
      if (permissions.length) {
        matchingItemCount += 1;
        if (matches.length < limit) {
          matches.push({ item: formatSimplifiedItem(item, "compact"), permissions, count: permissions.length });
        }
      }
    }
  }
  const unauditedCount = Math.max(0, candidates.length - auditedCount - errorCount);
  const unseenCandidateCount = Math.max(0, (scanResult.summary?.matched || candidates.length) - candidates.length);
  const scanResultTruncated = Boolean(scanResult.summary?.resultTruncated);
  const resultTruncated = matchingItemCount > matches.length || unauditedCount > 0 || scanResultTruncated;
  const incomplete = Boolean(scanResult.summary?.traversalTruncated || scanResultTruncated || unseenCandidateCount || errorCount || resultTruncated);
  return {
    scanSummary: scanResult.summary,
    count: matches.length,
    itemsReturned: matches.length,
    scanItemsReturned: scanResult.summary.returned,
    candidateCount: candidates.length,
    auditedCount,
    errorCount,
    unauditedCount,
    unseenCandidateCount,
    matchingItemCount,
    resultTruncated,
    incomplete,
    errors,
    items: matches,
    note: incomplete
      ? "Sharing audit is incomplete because traversal was bounded, permission reads failed, or the result limit was reached. Review the audit counters and errors."
      : publicOnly
      ? "Returned items with anonymous sharing links."
      : (args.includeOwnerPermissions === true
        ? "Returned items with explicit sharing permissions plus owner grants."
        : "Returned items with explicit non-owner sharing permissions.")
  };
}

async function existingReplacementTarget(remotePathValue) {
  try {
    return await getRawInfo({ path: remotePathValue });
  } catch (error) {
    if (error?.graphStatus === 404) return null;
    throw error;
  }
}

function trustedChatgptFileUrl(value) {
  let target;
  try {
    target = new URL(String(value || ""));
  } catch {
    throw new Error("sourceFile.download_url must be a valid HTTPS URL.");
  }
  const host = target.hostname.toLowerCase();
  const trustedHost = host === "files.openai.com"
    || host.endsWith(".files.openai.com")
    || host === "files.oaiusercontent.com"
    || host.endsWith(".oaiusercontent.com")
    || host.endsWith(".openaiusercontent.com");
  if (target.protocol !== "https:" || target.username || target.password || !trustedHost) {
    throw new Error("Refusing to download a ChatGPT file from an untrusted URL.");
  }
  return target;
}

async function downloadChatgptFile(sourceFile = {}) {
  await ensurePrivateDirectory(chatgptUploadRoot);
  const targetPath = join(chatgptUploadRoot, `${randomUUID()}.upload`);
  let current = trustedChatgptFileUrl(sourceFile.download_url);
  let response;
  for (let redirects = 0; redirects <= 3; redirects += 1) {
    response = await fetchWithRetry(current.toString(), { method: "GET", redirect: "manual" }, { maxRetries: 2 });
    if (![301, 302, 303, 307, 308].includes(response.status)) break;
    const location = response.headers.get("location");
    if (!location) throw new Error("ChatGPT file download redirected without a Location header.");
    if (redirects === 3) throw new Error("ChatGPT file download exceeded the redirect limit.");
    current = trustedChatgptFileUrl(new URL(location, current).toString());
  }
  if (!response?.ok) {
    await parseResponseBody(response).catch(() => null);
    throw new Error(`ChatGPT file download failed with HTTP ${response?.status || "unknown"}.`);
  }
  const declaredBytes = contentLength(response);
  if (declaredBytes !== null && declaredBytes > simpleUploadLimit) {
    throw new Error(`ChatGPT file is larger than the ${simpleUploadLimit}-byte safe upload limit.`);
  }
  let bytesWritten = 0;
  try {
    if (response.body) {
      const counter = new TransformStream({
        transform(chunk, controller) {
          bytesWritten += chunk.byteLength;
          if (bytesWritten > simpleUploadLimit) {
            controller.error(new Error(`ChatGPT file exceeded the ${simpleUploadLimit}-byte safe upload limit.`));
            return;
          }
          controller.enqueue(chunk);
        }
      });
      await pipeline(Readable.fromWeb(response.body.pipeThrough(counter)), createWriteStream(targetPath, { flags: "wx", mode: 0o600 }));
    } else {
      const buffer = Buffer.from(await response.arrayBuffer());
      if (buffer.length > simpleUploadLimit) throw new Error(`ChatGPT file exceeded the ${simpleUploadLimit}-byte safe upload limit.`);
      bytesWritten = buffer.length;
      await writePrivateFile(targetPath, buffer);
    }
    await hardenPrivateFile(targetPath);
    return { localPath: targetPath, bytes: bytesWritten };
  } catch (error) {
    await rm(targetPath, { force: true }).catch(() => null);
    throw error;
  }
}

async function uploadChatgptFile(args = {}) {
  const destinationPath = assertSafeRemotePath(args.remotePath, "remotePath");
  if (!destinationPath) throw new Error("remotePath must include a filename.");
  const conflictBehavior = args.conflictBehavior || "fail";
  const source = {
    fileId: String(args.sourceFile?.file_id || ""),
    fileName: String(args.sourceFile?.file_name || basename(destinationPath)),
    mimeType: String(args.sourceFile?.mime_type || "")
  };
  const current = conflictBehavior === "replace" ? await existingReplacementTarget(destinationPath) : null;
  if (current?.folder) throw new Error(`Refusing to replace a folder with file content: ${current.name}`);
  const proof = {
    destinationPath,
    conflictBehavior,
    source,
    existing: current ? itemVersionProof(current) : null
  };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldUpload: { source, destinationPath, conflictBehavior },
    wouldReplace: current ? simplifyItem(current) : null
  };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_upload_file", proof);
  if (args.confirmed !== true) {
    return { ...preview, dryRun: false, requiredToUpload: "Set dryRun: false and confirmed: true after reviewing the upload preview." };
  }
  if (current && !hasExpectedIdentity(args)) {
    return { ...preview, dryRun: false, confirmed: true, requiredToUpload: "Provide expectedName or expectedId matching the existing remote file." };
  }
  if (current) assertExpectedItem(current, args, "Upload replacement");
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_upload_file", proof, args.previewToken, "requiredToUpload");
  if (previewTokenRequired) return previewTokenRequired;

  const downloaded = await downloadChatgptFile(args.sourceFile);
  try {
    const uploaded = await upload({
      localPath: downloaded.localPath,
      remotePath: destinationPath,
      conflictBehavior: current ? "replace" : conflictBehavior === "replace" ? "fail" : conflictBehavior,
      guardedInternalReplace: true,
      ifMatch: current?.eTag,
      uploadMode: "simple",
      auditTool: "onedrive_upload_file",
      auditSource: source
    });
    const { localPath: omittedLocalPath, ...safeUploaded } = uploaded;
    return { ...safeUploaded, sourceFile: source };
  } finally {
    await rm(downloaded.localPath, { force: true }).catch(() => null);
  }
}

async function guardPublicReplacement(args, destinationPath, replacement, toolName) {
  if (args.conflictBehavior !== "replace" || args.itemId || args.guardedInternalReplace === true) return null;
  const current = await existingReplacementTarget(destinationPath);
  if (!current) return { allowed: true, createOnly: true };
  if (current.folder) throw new Error(`Refusing to replace a folder with file content: ${current.name}`);
  const proof = {
    destinationPath,
    existing: { id: current.id, name: current.name, eTag: current.eTag, cTag: current.cTag },
    replacement
  };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldReplace: simplifyItem(current),
    replacement
  };
  if (args.dryRun !== false) return previewWithToken(preview, toolName, proof);
  if (args.confirmed !== true) {
    return { ...preview, dryRun: false, requiredToReplace: "Set dryRun: false and confirmed: true after reviewing the existing-file replacement preview." };
  }
  if (!hasExpectedIdentity(args)) {
    return { ...preview, dryRun: false, confirmed: true, requiredToReplace: "Provide expectedName or expectedId matching the existing remote file." };
  }
  assertExpectedItem(current, args, "Replace");
  return previewTokenRequiredResult(preview, toolName, proof, args.previewToken, "requiredToReplace") || { allowed: true, current };
}

async function upload(args = {}) {
  const localPath = resolve(args.localPath);
  await assertNotLocalOneDriveSyncPathForRead(localPath, "Upload", args);
  const destinationPath = args.itemId ? (args.remotePath || `item:${args.itemId}`) : remotePath(args);
  const fileStat = await stat(localPath);
  if (!fileStat.isFile()) throw new Error(`Not a file: ${localPath}`);
  const replacementGuard = await guardPublicReplacement(args, destinationPath, {
    localPath,
    bytes: fileStat.size,
    modifiedAtMs: Math.trunc(fileStat.mtimeMs)
  }, "onedrive_upload");
  if (replacementGuard && replacementGuard.allowed !== true) return replacementGuard;
  if (replacementGuard?.allowed === true && replacementGuard.current?.eTag && !args.ifMatch) {
    args = { ...args, ifMatch: replacementGuard.current.eTag };
  }
  const effectiveConflictBehavior = replacementGuard?.createOnly === true ? "fail" : (args.conflictBehavior || "fail");
  const ifNoneMatch = replacementGuard?.createOnly === true ? "*" : undefined;
  const uploadMode = args.uploadMode || "auto";
  const auditTool = args.auditTool || "onedrive_upload";
  try {
    let response;
    if (fileStat.size > 0 && (uploadMode === "session" || (uploadMode === "auto" && fileStat.size > simpleUploadLimit))) {
      const sessionTarget = args.itemId
        ? { endpoints: [`/me/drive/items/${encodeURIComponent(args.itemId)}/createUploadSession`], name: basename(destinationPath) }
        : undefined;
      response = await uploadLarge({ ...args, localPath, remotePath: destinationPath, sessionTarget, effectiveConflictBehavior, ifNoneMatch }, fileStat);
    } else {
      if (uploadMode === "simple" && fileStat.size > simpleUploadLimit) {
        throw new Error(`Simple upload only supports files up to ${simpleUploadLimit} bytes. Use uploadMode: "session" or "auto".`);
      }
      const stream = createReadStream(localPath);
      const targetPath = args.itemId
        ? `/me/drive/items/${encodeURIComponent(args.itemId)}/content`
        : uploadPath(destinationPath, effectiveConflictBehavior);
      const result = await graph(targetPath, {
        method: "PUT",
        body: stream,
        duplex: "half",
        headers: {
          "Content-Type": "application/octet-stream",
          ...(args.ifMatch ? { "If-Match": args.ifMatch } : {}),
          ...(ifNoneMatch ? { "If-None-Match": ifNoneMatch } : {})
        }
      });
      await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
      response = { item: simplifyItem(result), localPath, bytesUploaded: fileStat.size, uploadMode: "simple" };
    }
    if (args.skipAudit !== true) {
      await writeMutationAudit(auditTool, {
        status: "success",
        target: { remotePath: destinationPath },
        after: itemAuditSummary(response.item),
        ...(args.auditSource ? { source: args.auditSource } : { localPath }),
        bytes: fileStat.size
      });
    }
    return response;
  } catch (error) {
    if (args.skipAudit !== true) {
      await writeMutationAudit(auditTool, {
        status: "failed",
        target: { remotePath: destinationPath },
        ...(args.auditSource ? { source: args.auditSource } : { localPath }),
        bytes: fileStat.size,
        error: safeErrorInfo(error)
      });
    }
    throw error;
  }
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
  const sessionTarget = args.sessionTarget || await uploadSessionTarget(args.remotePath);
  const session = await createUploadSession(
    sessionTarget,
    args.effectiveConflictBehavior || args.conflictBehavior || "fail",
    { ifMatch: args.ifMatch, ifNoneMatch: args.ifNoneMatch }
  );
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
        response = await graph(assertTrustedUploadSessionUrl(session.uploadUrl), {
          method: "PUT",
          skipAuth: true,
          body,
          headers: {
            "Content-Range": `bytes ${position}-${end}/${fileStat.size}`
          },
          maxRetries: 4
        });
      } catch (error) {
        throw new Error(`Upload session failed for byte range ${position}-${end}/${fileStat.size}: ${safeToolErrorMessage(error)}`);
      }
      uploaded = end + 1;
      if (response?.id) finalItem = response;
      if (!response?.id && Array.isArray(response?.nextExpectedRanges)) {
        const nextStart = Number(String(response.nextExpectedRanges[0] || "").split("-")[0]);
        if (Number.isFinite(nextStart) && nextStart !== uploaded) {
          throw new Error(`Upload session expected next byte ${nextStart}, but local upload position is ${uploaded}.`);
        }
      }
      position += bytesRead;
    }
  } finally {
    await handle.close();
  }

  if (!finalItem) throw new Error("Upload session completed local chunks but Microsoft Graph did not return a final drive item.");
  await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([finalItem]));
  return {
    item: simplifyItem(finalItem),
    localPath: args.localPath,
    bytesUploaded: uploaded,
    uploadMode: "session",
    chunkSize
  };
}

async function createUploadSession(sessionTarget, conflictBehavior, conditions = {}) {
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
    ...(conditions.ifNoneMatch ? [] : [null])
  ];
  const errors = [];
  for (const endpoint of sessionTarget.endpoints) {
    for (const body of bodies) {
      try {
        return await graph(endpoint, {
          method: "POST",
          ...((conditions.ifMatch || conditions.ifNoneMatch) ? { headers: {
            ...(conditions.ifMatch ? { "If-Match": conditions.ifMatch } : {}),
            ...(conditions.ifNoneMatch ? { "If-None-Match": conditions.ifNoneMatch } : {})
          } } : {}),
          ...(body ? { body: JSON.stringify(body) } : {})
        });
      } catch (error) {
        errors.push(`${endpoint}: ${safeToolErrorMessage(error)}`);
      }
    }
  }
  throw new Error(`Could not create upload session. Tried ${errors.length} compatible request shapes. Last error: ${errors.at(-1)}`);
}

async function writeText(args = {}) {
  const destinationPath = remotePath(args);
  const contentBuffer = Buffer.from(args.content, "utf8");
  const replacementGuard = await guardPublicReplacement(args, destinationPath, {
    bytes: contentBuffer.length,
    sha256: createHash("sha256").update(contentBuffer).digest("hex")
  }, "onedrive_write_text");
  if (replacementGuard && replacementGuard.allowed !== true) return replacementGuard;
  const effectiveConflictBehavior = replacementGuard?.createOnly === true ? "fail" : (args.conflictBehavior || "fail");
  try {
    const result = await graph(uploadPath(destinationPath, effectiveConflictBehavior), {
      method: "PUT",
      body: contentBuffer,
      headers: {
        "Content-Type": "text/plain; charset=utf-8",
        ...(replacementGuard?.allowed === true && replacementGuard.current?.eTag ? { "If-Match": replacementGuard.current.eTag } : {}),
        ...(replacementGuard?.createOnly === true ? { "If-None-Match": "*" } : {})
      }
    });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
    const response = { item: simplifyItem(result), bytesUploaded: contentBuffer.length };
    await writeMutationAudit("onedrive_write_text", {
      status: "success",
      target: { remotePath: destinationPath },
      after: itemAuditSummary(response.item),
      bytes: response.bytesUploaded
    });
    return response;
  } catch (error) {
    await writeMutationAudit("onedrive_write_text", {
      status: "failed",
      target: { remotePath: destinationPath },
      bytes: Buffer.byteLength(args.content || "", "utf8"),
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function driveItemVersionPath(itemId, versionId = null, suffix = "") {
  const base = `/me/drive/items/${encodeURIComponent(itemId)}/versions`;
  return `${base}${versionId === null ? "" : `/${encodeURIComponent(String(versionId))}`}${suffix}`;
}

function simplifyDriveItemVersion(version = {}) {
  return {
    id: version.id,
    size: version.size,
    lastModifiedDateTime: version.lastModifiedDateTime,
    lastModifiedBy: compactIdentity(version.lastModifiedBy),
    publication: version.publication,
    contentSha1: version.file?.hashes?.sha1Hash,
    contentSha256: version.file?.hashes?.sha256Hash
  };
}

async function listDriveItemVersions(rawItem, maxItems = 50) {
  const limit = clampInteger(maxItems, 50, 1, 200);
  const versions = [];
  let next = `${driveItemVersionPath(rawItem.id)}?$top=${Math.min(limit, 200)}`;
  while (next && versions.length < limit) {
    const page = await graph(next);
    versions.push(...(page.value || []).slice(0, limit - versions.length));
    next = page["@odata.nextLink"] || null;
  }
  return versions;
}

async function versions(args = {}) {
  const rawItem = await getRawInfo(args);
  if (rawItem.folder) throw new Error("Version history is available only for files.");
  const entries = await listDriveItemVersions(rawItem, args.maxItems);
  return {
    item: simplifyItem(rawItem),
    versions: entries.map(simplifyDriveItemVersion),
    count: entries.length,
    bounded: entries.length >= clampInteger(args.maxItems, 50, 1, 200)
  };
}

async function readVersionContent(rawItem, versionId = null, maxBytes = maxOfficePackageBytes) {
  const path = versionId === null
    ? contentPath({ itemId: rawItem.id })
    : driveItemVersionPath(rawItem.id, versionId, "/content");
  const result = await graphLimitedBuffer(path, maxBytes);
  if (result.truncated) throw new Error(`Version content is above the ${maxBytes}-byte comparison limit.`);
  return result.buffer;
}

async function inspectOfficeBuffer(kind, buffer, name, label) {
  const transactionRoot = join(officeEditingRoot, `version-${randomUUID()}`);
  await ensurePrivateDirectory(transactionRoot);
  const localPath = join(transactionRoot, assertSafeItemName(name || `document${kind === "word" ? ".docx" : kind === "excel" ? ".xlsx" : ".pptx"}`));
  try {
    await writePrivateFile(localPath, buffer);
    return await runOfficeHelper({ action: "inspect", inputPath: localPath, kind, maxParagraphs: 10000, maxCells: 50000, maxSlides: 5000 });
  } catch (error) {
    throw new Error(`Could not inspect ${label}: ${safeToolErrorMessage(error)}`);
  } finally {
    await rm(transactionRoot, { recursive: true, force: true });
  }
}

function binaryFingerprint(buffer) {
  return { bytes: buffer.length, sha256: createHash("sha256").update(buffer).digest("hex") };
}

async function compareItemContents(rawItem, leftBuffer, rightBuffer, labels = {}, maxChanges = 200) {
  const kind = officePackageKindFromName(rawItem.name);
  if (kind) {
    const [left, right] = await Promise.all([
      inspectOfficeBuffer(kind, leftBuffer, rawItem.name, labels.left || "left version"),
      inspectOfficeBuffer(kind, rightBuffer, rawItem.name, labels.right || "right version")
    ]);
    return { comparisonType: "office-semantic", ...compareOfficeInspections(kind, left, right, maxChanges) };
  }
  if (isLikelyTextItem(rawItem, { path: rawItem.name })) {
    const left = decodeTextBuffer(leftBuffer);
    const right = decodeTextBuffer(rightBuffer);
    return {
      comparisonType: "text",
      sameContent: left.text === right.text,
      left: { ...binaryFingerprint(leftBuffer), encoding: left.encoding, newline: left.newline, trailingNewline: left.trailingNewline },
      right: { ...binaryFingerprint(rightBuffer), encoding: right.encoding, newline: right.newline, trailingNewline: right.trailingNewline },
      diff: boundedLineDiff(left.text, right.text, maxChanges)
    };
  }
  const left = binaryFingerprint(leftBuffer);
  const right = binaryFingerprint(rightBuffer);
  return { comparisonType: "binary", sameContent: left.sha256 === right.sha256, left, right };
}

async function compareVersion(args = {}) {
  const rawItem = await getRawInfo(args);
  if (rawItem.folder) throw new Error("Version comparison is available only for files.");
  const versions = await listDriveItemVersions(rawItem, 200);
  if (!versions.some((entry) => String(entry.id) === String(args.versionId))) throw new Error(`Version ${args.versionId} was not found for ${rawItem.name}.`);
  if (args.compareToVersionId && !versions.some((entry) => String(entry.id) === String(args.compareToVersionId))) {
    throw new Error(`Version ${args.compareToVersionId} was not found for ${rawItem.name}.`);
  }
  const [left, right] = await Promise.all([
    readVersionContent(rawItem, args.versionId),
    readVersionContent(rawItem, args.compareToVersionId || null)
  ]);
  return {
    item: simplifyItem(rawItem),
    leftVersionId: args.versionId,
    rightVersionId: args.compareToVersionId || "current",
    comparison: await compareItemContents(rawItem, left, right, {
      left: `version ${args.versionId}`,
      right: args.compareToVersionId ? `version ${args.compareToVersionId}` : "current version"
    }, clampInteger(args.maxChanges, 200, 1, 1000))
  };
}

async function restoreVersion(args = {}) {
  const rawItem = await getRawInfo(args);
  if (rawItem.folder) throw new Error("Version restore is available only for files.");
  assertExpectedItem(rawItem, args, "Version restore");
  const entries = await listDriveItemVersions(rawItem, 200);
  const target = entries.find((entry) => String(entry.id) === String(args.versionId));
  if (!target) throw new Error(`Version ${args.versionId} was not found for ${rawItem.name}.`);
  const preview = {
    dryRun: true,
    confirmed: false,
    item: simplifyItem(rawItem),
    version: simplifyDriveItemVersion(target),
    nativeRestore: true,
    warning: "Restoring creates a new current version. It does not replace content through an upload fallback."
  };
  const proof = { itemId: rawItem.id, eTag: rawItem.eTag, versionId: String(target.id) };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_restore_version", proof);
  if (args.confirmed !== true || !hasExpectedIdentity(args) || !args.expectedETag) {
    return { ...preview, dryRun: false, confirmed: args.confirmed === true, requiredToRestore: "Pass confirmed:true, expectedId or expectedName, expectedETag, and the matching previewToken." };
  }
  if (args.expectedETag !== rawItem.eTag) throw new Error("Version restore expectedETag no longer matches the current item. Run a fresh preview.");
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_restore_version", proof, args.previewToken, "requiredToRestore");
  if (tokenRequired) return tokenRequired;
  try {
    await graph(driveItemVersionPath(rawItem.id, target.id, "/restoreVersion"), { method: "POST", headers: { "If-Match": rawItem.eTag }, maxRetries: 0 });
    const after = await getRawInfo({ itemId: rawItem.id, cacheResults: false });
    if (after.eTag === rawItem.eTag) throw new Error("Graph accepted the restore but the current eTag did not change; restore verification failed.");
    await writeMutationAudit("onedrive_restore_version", { status: "success", target: itemAuditSummary(rawItem), before: itemAuditSummary(rawItem), after: itemAuditSummary(after), versionId: String(target.id) });
    return { dryRun: false, confirmed: true, restoredVersionId: String(target.id), item: simplifyItem(after), verified: true };
  } catch (error) {
    await writeMutationAudit("onedrive_restore_version", { status: "failed", target: itemAuditSummary(rawItem), versionId: String(target.id), error: safeErrorInfo(error) });
    throw error;
  }
}

async function patchText(args = {}) {
  const rawItem = await getRawInfo(args);
  assertExpectedItem(rawItem, args, "Text patch");
  assertTextReadable(rawItem, args);
  if (rawItem.size > maxTextFileReadLimit) throw new Error(`File is ${rawItem.size} bytes, above the bounded-edit limit ${maxTextFileReadLimit}.`);
  const source = await readVersionContent(rawItem, null, maxTextFileReadLimit);
  const patched = applyTextPatch(source, args.patch);
  if (patched.bytes.equals(source)) throw new Error("The patch is a no-op; no remote mutation is needed.");
  const proof = {
    itemId: rawItem.id,
    eTag: rawItem.eTag,
    sourceSha256: binaryFingerprint(source).sha256,
    resultSha256: binaryFingerprint(patched.bytes).sha256,
    patch: stablePreviewValue(args.patch)
  };
  const preview = {
    dryRun: true,
    confirmed: false,
    item: simplifyItem(rawItem),
    patch: { mode: args.patch.mode, operationCount: args.patch.operations?.length || 1 },
    preservation: { encoding: patched.encoding, bom: patched.bom, newline: patched.newline, trailingNewline: patched.trailingNewline },
    result: binaryFingerprint(patched.bytes),
    diff: boundedLineDiff(patched.beforeText, patched.afterText, 300)
  };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_patch_text", proof);
  if (args.confirmed !== true || !hasExpectedIdentity(args) || !args.expectedETag) {
    return { ...preview, dryRun: false, confirmed: args.confirmed === true, requiredToPatch: "Pass confirmed:true, expectedId or expectedName, expectedETag, and the matching previewToken." };
  }
  if (args.expectedETag !== rawItem.eTag) throw new Error("Text patch expectedETag no longer matches the current item. Run a fresh preview.");
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_patch_text", proof, args.previewToken, "requiredToPatch");
  if (tokenRequired) return tokenRequired;
  try {
    const result = await graph(`${itemIdBase(rawItem.id)}/content`, {
      method: "PUT",
      body: patched.bytes,
      headers: { "Content-Type": rawItem.file?.mimeType || "application/octet-stream", "If-Match": rawItem.eTag },
      maxRetries: 0
    });
    const verification = await readVersionContent(result?.id ? result : rawItem, null, maxTextFileReadLimit);
    const expectedHash = binaryFingerprint(patched.bytes).sha256;
    const observedHash = binaryFingerprint(verification).sha256;
    if (expectedHash !== observedHash) throw new Error("Text patch post-commit verification failed.");
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
    await writeMutationAudit("onedrive_patch_text", { status: "success", target: itemAuditSummary(rawItem), before: itemAuditSummary(rawItem), after: itemAuditSummary(result), patchMode: args.patch.mode, sha256: observedHash });
    return { dryRun: false, confirmed: true, item: simplifyItem(result), patch: preview.patch, preservation: preview.preservation, verified: true, sha256: observedHash };
  } catch (error) {
    await writeMutationAudit("onedrive_patch_text", { status: "failed", target: itemAuditSummary(rawItem), patchMode: args.patch.mode, error: safeErrorInfo(error) });
    throw error;
  }
}

function emptyWorkspaceState(scope) {
  return { version: 1, scope, updatedAt: null, workspaces: {} };
}

async function loadWorkspaceState() {
  const scope = await activeStorageScope();
  try {
    const state = JSON.parse(await readFile(scopedStatePath(workspaceStateRoot, scope), "utf8"));
    if (state?.version !== 1 || !storageScopesEqual(state.scope, scope) || !state.workspaces || typeof state.workspaces !== "object") {
      return emptyWorkspaceState(scope);
    }
    return state;
  } catch (error) {
    if (error.code !== "ENOENT" && !(error instanceof SyntaxError)) throw error;
    return emptyWorkspaceState(scope);
  }
}

async function saveWorkspaceState(state) {
  const scope = await activeStorageScope();
  if (!storageScopesEqual(state.scope, scope)) throw new Error("Workspace state scope changed; refusing to persist it.");
  state.updatedAt = new Date().toISOString();
  await ensurePrivateDirectory(workspaceStateRoot);
  await writePrivateFileAtomic(scopedStatePath(workspaceStateRoot, scope), `${JSON.stringify(state, null, 2)}\n`);
}

function assertOpaqueId(candidate, label) {
  const value = String(candidate || "");
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)) throw new Error(`${label} is invalid.`);
  return value.toLowerCase();
}

const assertWorkspaceId = (workspaceId) => assertOpaqueId(workspaceId, "workspaceId");
const assertWatchId = (watchId) => assertOpaqueId(watchId, "watchId");

async function workspaceManifest(workspaceId) {
  const state = await loadWorkspaceState();
  const id = assertWorkspaceId(workspaceId);
  const manifest = state.workspaces[id];
  if (!manifest) throw new Error(`Workspace ${id} was not found in the current OneDrive account and drive scope.`);
  return { state, id, manifest };
}

async function ensureOwnerOnlyWorkspaceRoot() {
  let root;
  try {
    root = await getRawInfo({ path: managedWorkspaceRootName, cacheResults: false });
  } catch (error) {
    if (error.graphStatus !== 404) throw error;
    root = await graph("/me/drive/root/children", {
      method: "POST",
      body: JSON.stringify({ name: managedWorkspaceRootName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
    });
  }
  if (!root.folder) throw new Error(`${managedWorkspaceRootName} exists but is not a folder.`);
  const permissions = await permissionList({ itemId: root.id }, "compact");
  const unsafe = permissions.filter((permission) => !isOwnerPermission(permission));
  if (unsafe.length) {
    throw new Error(`${managedWorkspaceRootName} has ${unsafe.length} non-owner permission(s). Remove sharing before using managed edit workspaces.`);
  }
  return root;
}

async function workspaceList() {
  const state = await loadWorkspaceState();
  return {
    rootName: managedWorkspaceRootName,
    scope: state.scope,
    workspaces: Object.values(state.workspaces).map((entry) => ({ ...entry, source: { ...entry.source }, draft: { ...entry.draft } })),
    count: Object.keys(state.workspaces).length
  };
}

async function workspaceCreate(args = {}) {
  const source = await getRawInfo(args);
  if (source.folder) throw new Error("Managed editing workspaces require a source file.");
  assertExpectedItem(source, args, "Workspace creation");
  const preview = {
    dryRun: true,
    confirmed: false,
    source: simplifyItem(source),
    managedRoot: managedWorkspaceRootName,
    ownerOnlyRequired: true,
    cleanup: "A successful promotion removes the draft; failed or conflicted workspaces remain recoverable."
  };
  const proof = { sourceId: source.id, sourceETag: source.eTag, managedRoot: managedWorkspaceRootName };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_workspace_create", proof);
  if (args.confirmed !== true || !hasExpectedIdentity(args) || !args.expectedETag) {
    return { ...preview, dryRun: false, confirmed: args.confirmed === true, requiredToCreate: "Pass confirmed:true, expectedId or expectedName, expectedETag, and the matching previewToken." };
  }
  if (args.expectedETag !== source.eTag) throw new Error("Workspace creation expectedETag no longer matches the source. Run a fresh preview.");
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_workspace_create", proof, args.previewToken, "requiredToCreate");
  if (tokenRequired) return tokenRequired;
  let workspaceFolder = null;
  try {
    const managedRoot = await ensureOwnerOnlyWorkspaceRoot();
    const workspaceId = randomUUID();
    workspaceFolder = await graph(`${itemIdBase(managedRoot.id)}/children`, {
      method: "POST",
      body: JSON.stringify({ name: workspaceId, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
    });
    const sourceBytes = await readVersionContent(source, null, maxOfficePackageBytes);
    const draft = await graph(`${itemIdBase(workspaceFolder.id)}:/${encodeURIComponent(source.name)}:/content?@microsoft.graph.conflictBehavior=fail`, {
      method: "PUT",
      body: sourceBytes,
      headers: { "Content-Type": source.file?.mimeType || "application/octet-stream", "If-None-Match": "*" },
      maxRetries: 0
    });
    const state = await loadWorkspaceState();
    const versionsAtCreation = await listDriveItemVersions(source, 1).catch(() => []);
    const manifest = {
      workspaceId,
      status: "ready",
      stale: false,
      createdAt: new Date().toISOString(),
      rootId: managedRoot.id,
      folderId: workspaceFolder.id,
      source: { id: source.id, name: source.name, eTag: source.eTag, cTag: source.cTag, versionId: versionsAtCreation[0]?.id || null },
      draft: { id: draft.id, name: draft.name, eTag: draft.eTag, cTag: draft.cTag }
    };
    state.workspaces[workspaceId] = manifest;
    await saveWorkspaceState(state);
    await writeMutationAudit("onedrive_workspace_create", { status: "success", target: itemAuditSummary(source), workspaceId, draft: itemAuditSummary(draft) });
    return { dryRun: false, confirmed: true, workspace: manifest, source: simplifyItem(source), draft: simplifyItem(draft) };
  } catch (error) {
    if (workspaceFolder?.id) await graph(itemIdBase(workspaceFolder.id), { method: "DELETE", maxRetries: 0 }).catch(() => null);
    await writeMutationAudit("onedrive_workspace_create", { status: "failed", target: itemAuditSummary(source), error: safeErrorInfo(error) });
    throw error;
  }
}

async function workspaceStatus(args = {}) {
  const { manifest } = await workspaceManifest(args.workspaceId);
  const [source, draft] = await Promise.all([
    getRawInfo({ itemId: manifest.source.id, cacheResults: false }),
    getRawInfo({ itemId: manifest.draft.id, cacheResults: false })
  ]);
  const sourceDrift = source.eTag !== manifest.source.eTag;
  const draftDrift = draft.eTag !== manifest.draft.eTag;
  const [sourceBytes, draftBytes] = await Promise.all([
    readVersionContent(source, null, maxOfficePackageBytes),
    readVersionContent(draft, null, maxOfficePackageBytes)
  ]);
  const comparison = await compareItemContents(source, sourceBytes, draftBytes, { left: "source", right: "draft" }, clampInteger(args.maxChanges, 200, 1, 1000));
  return {
    workspaceId: manifest.workspaceId,
    status: sourceDrift || manifest.stale ? "conflicted" : draftDrift ? "edited" : "ready",
    source: simplifyItem(source),
    draft: simplifyItem(draft),
    sourceDrift,
    draftDrift,
    markedStale: manifest.stale === true,
    promotionReady: !sourceDrift && manifest.stale !== true,
    comparison
  };
}

async function workspacePromote(args = {}) {
  const { state, id, manifest } = await workspaceManifest(args.workspaceId);
  const status = await workspaceStatus({ workspaceId: id, maxChanges: 300 });
  const proof = { workspaceId: id, sourceId: status.source.id, sourceETag: status.source.eTag, draftId: status.draft.id, draftETag: status.draft.eTag };
  const preview = { dryRun: true, confirmed: false, ...status, action: "promote draft to the original stable item ID" };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_workspace_promote", proof);
  if (args.confirmed !== true || args.expectedId !== status.source.id || !args.expectedETag) {
    return { ...preview, dryRun: false, confirmed: args.confirmed === true, requiredToPromote: "Pass confirmed:true, the exact source expectedId, expectedETag, and the matching previewToken." };
  }
  if (args.expectedETag !== status.source.eTag || !status.promotionReady) throw new Error("Workspace source drifted or was marked stale. Run a fresh status/preview; promotion is blocked.");
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_workspace_promote", proof, args.previewToken, "requiredToPromote");
  if (tokenRequired) return tokenRequired;
  try {
    const draftRaw = await getRawInfo({ itemId: manifest.draft.id, cacheResults: false });
    const draftBytes = await readVersionContent(draftRaw, null, maxOfficePackageBytes);
    const result = await graph(`${itemIdBase(manifest.source.id)}/content`, {
      method: "PUT", body: draftBytes, headers: { "Content-Type": draftRaw.file?.mimeType || "application/octet-stream", "If-Match": status.source.eTag }, maxRetries: 0
    });
    const verify = await readVersionContent(result, null, maxOfficePackageBytes);
    if (binaryFingerprint(verify).sha256 !== binaryFingerprint(draftBytes).sha256) throw new Error("Workspace promotion post-commit verification failed.");
    await graph(itemIdBase(manifest.folderId), { method: "DELETE", maxRetries: 0 });
    delete state.workspaces[id];
    await saveWorkspaceState(state);
    await writeMutationAudit("onedrive_workspace_promote", { status: "success", target: itemAuditSummary(status.source), after: itemAuditSummary(result), workspaceId: id, draftId: manifest.draft.id });
    return { dryRun: false, confirmed: true, workspaceId: id, promoted: true, cleanedUp: true, item: simplifyItem(result), verified: true };
  } catch (error) {
    await writeMutationAudit("onedrive_workspace_promote", { status: "failed", target: itemAuditSummary(status.source), workspaceId: id, draftRetained: true, error: safeErrorInfo(error) });
    throw error;
  }
}

async function workspaceAbandon(args = {}) {
  const { state, id, manifest } = await workspaceManifest(args.workspaceId);
  const draft = await getRawInfo({ itemId: manifest.draft.id, cacheResults: false });
  const preview = { dryRun: true, confirmed: false, workspace: manifest, currentDraft: simplifyItem(draft), wouldDeleteFolderId: manifest.folderId, sourceUnaffected: true };
  const proof = { workspaceId: id, folderId: manifest.folderId, draftId: draft.id, draftETag: draft.eTag };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_workspace_abandon", proof);
  if (args.confirmed !== true || args.expectedId !== draft.id || !args.expectedETag) return { ...preview, dryRun: false, requiredToAbandon: "Pass confirmed:true, the exact draft expectedId, expectedETag, and the matching previewToken." };
  if (args.expectedETag !== draft.eTag) throw new Error("Workspace draft changed. Run a fresh abandonment preview.");
  const tokenRequired = previewTokenRequiredResult(preview, "onedrive_workspace_abandon", proof, args.previewToken, "requiredToAbandon");
  if (tokenRequired) return tokenRequired;
  try {
    await graph(itemIdBase(manifest.folderId), { method: "DELETE", maxRetries: 0 });
    delete state.workspaces[id];
    await saveWorkspaceState(state);
    await writeMutationAudit("onedrive_workspace_abandon", { status: "success", workspaceId: id, draftId: manifest.draft.id });
    return { dryRun: false, confirmed: true, workspaceId: id, abandoned: true, sourceUnaffected: true };
  } catch (error) {
    await writeMutationAudit("onedrive_workspace_abandon", { status: "failed", workspaceId: id, draftId: manifest.draft.id, error: safeErrorInfo(error) });
    throw error;
  }
}

function emptyWatchState(scope) {
  return { version: 1, scope, updatedAt: null, watches: {} };
}

async function loadWatchState() {
  const scope = await activeStorageScope();
  try {
    const state = JSON.parse(await readFile(scopedStatePath(watchStateRoot, scope), "utf8"));
    if (state?.version !== 1 || !storageScopesEqual(state.scope, scope) || !state.watches || typeof state.watches !== "object") return emptyWatchState(scope);
    return state;
  } catch (error) {
    if (error.code !== "ENOENT" && !(error instanceof SyntaxError)) throw error;
    return emptyWatchState(scope);
  }
}

async function saveWatchState(state) {
  const scope = await activeStorageScope();
  if (!storageScopesEqual(state.scope, scope)) throw new Error("Watch state scope changed; refusing to persist it.");
  state.updatedAt = new Date().toISOString();
  await ensurePrivateDirectory(watchStateRoot);
  await writePrivateFileAtomic(scopedStatePath(watchStateRoot, scope), `${JSON.stringify(state, null, 2)}\n`);
}

function scheduleWatch(watch) {
  const existing = watchTimers.get(watch.watchId);
  if (existing) clearTimeout(existing);
  if (watch.status !== "active" || Date.parse(watch.expiresAt) <= Date.now()) return;
  const delayMs = Math.max(1000, Number(watch.nextPollAt ? Date.parse(watch.nextPollAt) - Date.now() : watch.intervalSeconds * 1000));
  const timer = setTimeout(() => {
    watchTimers.delete(watch.watchId);
    pollWatch(watch.watchId).catch((error) => recordLocalWarning(`watch ${watch.watchId} poll`, error));
  }, delayMs);
  timer.unref?.();
  watchTimers.set(watch.watchId, timer);
}

async function markWorkspacesStale(changedIds) {
  const changed = new Set(changedIds);
  const state = await loadWorkspaceState();
  let dirty = false;
  for (const manifest of Object.values(state.workspaces)) {
    if (changed.has(manifest.source.id)) {
      manifest.stale = true;
      manifest.staleAt = new Date().toISOString();
      dirty = true;
    } else if (changed.has(manifest.draft.id)) {
      manifest.draftChangedAt = new Date().toISOString();
      dirty = true;
    }
  }
  if (dirty) await saveWorkspaceState(state);
}

function invalidatePreviewTokensForItems(changedIds) {
  const changed = new Set(changedIds);
  let invalidated = 0;
  for (const [token, entry] of previewTokens) {
    if ((entry.itemIds || []).some((id) => changed.has(id))) {
      previewTokens.delete(token);
      invalidated += 1;
    }
  }
  return invalidated;
}

function watchEvent(item) {
  return {
    sequence: null,
    observedAt: new Date().toISOString(),
    itemId: item.id,
    name: item.name,
    remotePath: item.remotePath,
    deleted: Boolean(item.deleted),
    eTag: item.eTag,
    cTag: item.cTag,
    lastModifiedDateTime: item.lastModifiedDateTime
  };
}

async function pollWatch(watchId) {
  const state = await loadWatchState();
  const watch = state.watches[watchId];
  if (!watch || watch.status !== "active") return;
  if (Date.parse(watch.expiresAt) <= Date.now()) {
    watch.status = "expired";
    watch.stoppedAt = new Date().toISOString();
    await saveWatchState(state);
    return;
  }
  try {
    let result = await delta({ deltaLink: watch.deltaLink, maxItems: 5000, pageSize: 200, format: "compact" });
    while (result.nextLink && !result.deltaLink) result = await delta({ nextLink: result.nextLink, maxItems: 5000, pageSize: 200, format: "compact" });
    const events = (result.items || []).map(watchEvent);
    const startSequence = Number(watch.lastSequence || 0);
    events.forEach((event, index) => { event.sequence = startSequence + index + 1; });
    watch.lastSequence = startSequence + events.length;
    watch.events = [...(watch.events || []), ...events].slice(-500);
    watch.eventCount = Number(watch.eventCount || 0) + events.length;
    watch.deltaLink = result.deltaLink || watch.deltaLink;
    watch.lastPollAt = new Date().toISOString();
    watch.consecutiveErrors = 0;
    watch.lastError = null;
    watch.nextPollAt = new Date(Date.now() + watch.intervalSeconds * 1000).toISOString();
    const changedIds = events.map((event) => event.itemId).filter(Boolean);
    watch.previewTokensInvalidated = Number(watch.previewTokensInvalidated || 0) + invalidatePreviewTokensForItems(changedIds);
    if (changedIds.length) {
      await markWorkspacesStale(changedIds);
    }
    await saveWatchState(state);
    scheduleWatch(watch);
  } catch (error) {
    watch.consecutiveErrors = Number(watch.consecutiveErrors || 0) + 1;
    watch.lastError = { at: new Date().toISOString(), message: safeToolErrorMessage(error), graphStatus: error.graphStatus || null };
    const backoffSeconds = Math.min(300, watch.intervalSeconds * 2 ** Math.min(watch.consecutiveErrors, 5));
    watch.nextPollAt = new Date(Date.now() + backoffSeconds * 1000).toISOString();
    if (error.graphStatus === 401 || error.graphStatus === 403 || error.graphStatus === 410) {
      watch.status = "error";
      watch.terminalReason = error.graphStatus === 410 ? "delta_cursor_expired" : "authentication_or_permission_error";
    }
    await saveWatchState(state);
    if (watch.status === "active") scheduleWatch(watch);
  }
}

async function ensureWatchesLoaded() {
  if (watchesLoaded) return;
  const state = await loadWatchState();
  watchesLoaded = true;
  for (const watch of Object.values(state.watches)) {
    if (watch.status === "active" && Date.parse(watch.expiresAt) > Date.now()) scheduleWatch(watch);
  }
}

async function watchStart(args = {}) {
  await ensureWatchesLoaded();
  const resolvedPath = args.itemId ? null : resolvePresetPath(args);
  const target = args.itemId
    ? await getRawInfo({ itemId: args.itemId, cacheResults: false })
    : resolvedPath
      ? await getRawInfo({ path: resolvedPath, cacheResults: false })
      : await graph("/me/drive/root");
  if (!target.folder && !target.root) throw new Error("Change watches require a folder or drive root target.");
  let baseline = await delta({ itemId: target.id, maxItems: 5000, pageSize: 200, format: "compact" });
  while (baseline.nextLink && !baseline.deltaLink) baseline = await delta({ nextLink: baseline.nextLink, maxItems: 5000, pageSize: 200, format: "compact" });
  if (!baseline.deltaLink) throw new Error("Could not establish a complete delta baseline for this watch.");
  const intervalSeconds = clampInteger(args.intervalSeconds, 30, 15, 300);
  const expiresInSeconds = clampInteger(args.expiresInSeconds, 3600, 60, 28800);
  const watchId = randomUUID();
  const now = Date.now();
  const watch = {
    watchId,
    status: "active",
    target: simplifyItem(target),
    intervalSeconds,
    createdAt: new Date(now).toISOString(),
    expiresAt: new Date(now + expiresInSeconds * 1000).toISOString(),
    nextPollAt: new Date(now + intervalSeconds * 1000).toISOString(),
    deltaLink: baseline.deltaLink,
    eventCount: 0,
    lastSequence: 0,
    events: [],
    previewTokensInvalidated: 0,
    consecutiveErrors: 0
  };
  const state = await loadWatchState();
  state.watches[watchId] = watch;
  await saveWatchState(state);
  scheduleWatch(watch);
  return { watch: { ...watch, deltaLink: undefined }, baselineItemCount: baseline.count, defaultIntervalSeconds: 30, loopbackOnly: true };
}

async function watchStatus(args = {}) {
  await ensureWatchesLoaded();
  const state = await loadWatchState();
  const selected = args.watchId ? [state.watches[assertWatchId(args.watchId)]].filter(Boolean) : Object.values(state.watches);
  if (args.watchId && !selected.length) throw new Error(`Watch ${args.watchId} was not found in the current OneDrive account and drive scope.`);
  const maxEvents = clampInteger(args.maxEvents, 100, 0, 500);
  const watches = selected.map((watch) => ({ ...watch, deltaLink: undefined, events: (watch.events || []).slice(-maxEvents) }));
  if (args.consume === true) {
    for (const watch of selected) watch.events = [];
    await saveWatchState(state);
  }
  return { watches, count: watches.length, eventsConsumed: args.consume === true };
}

async function watchStop(args = {}) {
  await ensureWatchesLoaded();
  const watchId = assertWatchId(args.watchId);
  const state = await loadWatchState();
  const watch = state.watches[watchId];
  if (!watch) throw new Error(`Watch ${watchId} was not found in the current OneDrive account and drive scope.`);
  const timer = watchTimers.get(watchId);
  if (timer) clearTimeout(timer);
  watchTimers.delete(watchId);
  watch.status = "stopped";
  watch.stoppedAt = new Date().toISOString();
  watch.nextPollAt = null;
  await saveWatchState(state);
  return { watch: { ...watch, deltaLink: undefined }, stopped: true };
}

async function createFolder(args = {}) {
  const name = assertSafeItemName(args.name, "name");
  assertAtMostOneSelector(args, "folder parent", [
    { label: "parentItemId", keys: ["parentItemId"] },
    { label: "parentPath", keys: ["parentPath"] },
    { label: "parentPreset", keys: ["parentPreset"] }
  ]);
  const endpoint = args.parentItemId
    ? `/me/drive/items/${encodeURIComponent(args.parentItemId)}/children`
    : childrenPath({
        path: args.parentPath || "",
        preset: args.parentPreset,
        relativePath: args.parentRelativePath
      });
  try {
    const result = await graph(endpoint, {
      method: "POST",
      body: JSON.stringify({
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": args.conflictBehavior || "fail"
      })
    });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([result]));
    const item = simplifyItem(result);
    await writeMutationAudit("onedrive_create_folder", {
      status: "success",
      target: { parentPath: args.parentPath, parentItemId: args.parentItemId, name },
      after: itemAuditSummary(item)
    });
    return item;
  } catch (error) {
    await writeMutationAudit("onedrive_create_folder", {
      status: "failed",
      target: { parentPath: args.parentPath, parentItemId: args.parentItemId, name },
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function rename(args = {}) {
  const newName = assertSafeItemName(args.newName, "newName");
  requireNonRootTarget(args, "Rename");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Rename refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Rename");
  const item = simplifyItem(current);
  const preview = { dryRun: args.dryRun !== false, confirmed: args.confirmed === true, wouldRename: item, newName };
  const previewProof = { items: [itemVersionProof(current)], operation: "rename", newName };
  if (args.dryRun !== false) {
    return previewWithToken(preview, "onedrive_rename", previewProof);
  }
  if (args.confirmed !== true) {
    return {
      ...preview,
      requiredToRename: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      dryRun: false,
      confirmed: true,
      wouldRename: item,
      newName,
      requiredToRename: "Provide expectedName or expectedId for live renames."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_rename", previewProof, args.previewToken, "requiredToRename");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const result = await graph(itemMutationBase(current), {
      method: "PATCH",
      headers: mutationMatchHeaders(current),
      body: JSON.stringify({ name: newName })
    });
    await cacheMovedOrRenamedItem(current, result);
    const renamed = simplifyItem(result);
    await writeMutationAudit("onedrive_rename", {
      status: "success",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      after: itemAuditSummary(renamed)
    });
    return { dryRun: false, confirmed: true, renamed };
  } catch (error) {
    await writeMutationAudit("onedrive_rename", {
      status: "failed",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      requestedName: newName,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function moveItem(args = {}) {
  if (args.newName) assertSafeItemName(args.newName, "newName");
  requireNonRootTarget(args, "Move");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Move refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Move");
  const parentReference = await resolveDestinationParent(args);
  const body = { parentReference: { id: parentReference.id } };
  if (args.newName) body.name = args.newName;
  const item = simplifyItem(current);
  const preview = { dryRun: args.dryRun !== false, confirmed: args.confirmed === true, wouldMove: item, destination: parentReference, newName: args.newName || null };
  const previewProof = { items: [itemVersionProof(current)], operation: "move", destinationParentId: parentReference.id, newName: args.newName || null };
  if (args.dryRun !== false) {
    return previewWithToken(preview, "onedrive_move", previewProof);
  }
  if (args.confirmed !== true) {
    return {
      ...preview,
      requiredToMove: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      dryRun: false,
      confirmed: true,
      wouldMove: item,
      destination: parentReference,
      newName: args.newName || null,
      requiredToMove: "Provide expectedName or expectedId for live moves."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_move", previewProof, args.previewToken, "requiredToMove");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const result = await graph(itemMutationBase(current), {
      method: "PATCH",
      headers: mutationMatchHeaders(current),
      body: JSON.stringify(body)
    });
    await cacheMovedOrRenamedItem(current, result);
    const moved = simplifyItem(result);
    await writeMutationAudit("onedrive_move", {
      status: "success",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      after: itemAuditSummary(moved),
      destination: parentReference
    });
    return { dryRun: false, confirmed: true, moved };
  } catch (error) {
    await writeMutationAudit("onedrive_move", {
      status: "failed",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      destination: parentReference,
      newName: args.newName || null,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function pollCopyMonitor(monitorUrl, timeoutSeconds = 60) {
  const trustedMonitorUrl = assertTrustedCopyMonitorUrl(monitorUrl);
  const safeMonitorUrl = safeDisplayPath(trustedMonitorUrl);
  const deadline = Date.now() + timeoutSeconds * 1000;
  let last = null;
  while (Date.now() < deadline) {
    const response = await graph(trustedMonitorUrl, { skipAuth: true, returnResponse: true, maxRetries: 3, redirect: "manual" });
    last = response.body && !(response.body instanceof ArrayBuffer) ? response.body : null;
    if (!response.ok && response.status !== 303) {
      return {
        complete: true,
        terminal: true,
        succeeded: false,
        terminalState: "failed",
        status: response.status,
        monitorUrl: safeMonitorUrl,
        monitor: sanitizeAuditValue(last)
      };
    }
    if (response.status === 303) {
      return {
        complete: true,
        terminal: true,
        succeeded: true,
        terminalState: "succeeded",
        status: response.status,
        resourceLocation: response.headers.get("location") ? safeDisplayPath(response.headers.get("location")) : null,
        monitorUrl: safeMonitorUrl
      };
    }
    if (response.ok && last?.status && !["notStarted", "running", "inProgress"].includes(String(last.status))) {
      const normalizedStatus = String(last.status).toLowerCase();
      const succeeded = ["completed", "complete", "succeeded", "success"].includes(normalizedStatus);
      const failed = ["failed", "error", "cancelled", "canceled"].includes(normalizedStatus);
      return {
        complete: true,
        terminal: true,
        succeeded,
        terminalState: succeeded ? "succeeded" : failed ? "failed" : "unknown",
        status: response.status,
        monitorUrl: safeMonitorUrl,
        monitor: sanitizeAuditValue(last)
      };
    }
    await sleep(2000);
  }
  return { complete: false, terminal: false, succeeded: false, terminalState: "timeout", timeoutSeconds, monitorUrl: safeMonitorUrl, monitor: sanitizeAuditValue(last) };
}

async function copyItem(args = {}) {
  if (args.newName) assertSafeItemName(args.newName, "newName");
  requireNonRootTarget(args, "Copy");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Copy refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Copy");
  const parentReference = await resolveDestinationParent(args);
  const item = simplifyItem(current);
  const preview = { dryRun: args.dryRun !== false, confirmed: args.confirmed === true, wouldCopy: item, destination: parentReference, newName: args.newName || null };
  const previewProof = { items: [itemVersionProof(current)], operation: "copy", destinationParentId: parentReference.id, newName: args.newName || null };
  if (args.dryRun !== false) {
    return previewWithToken(preview, "onedrive_copy", previewProof);
  }
  if (args.confirmed !== true) {
    return {
      ...preview,
      requiredToCopy: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      dryRun: false,
      confirmed: true,
      wouldCopy: item,
      destination: parentReference,
      newName: args.newName || null,
      requiredToCopy: "Provide expectedName or expectedId for live copies."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_copy", previewProof, args.previewToken, "requiredToCopy");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const response = await graph(`${itemMutationBase(current)}/copy`, {
      method: "POST",
      returnResponse: true,
      body: JSON.stringify({
        parentReference: { id: parentReference.id },
        ...(args.newName ? { name: args.newName } : {})
      })
    });
    if (!response.ok) throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Copy failed" });
    const monitorUrl = response.headers.get("location");
    const result = {
      dryRun: false,
      confirmed: true,
      accepted: response.status === 202 || response.ok,
      status: response.status,
      source: item,
      monitorUrl: monitorUrl ? safeDisplayPath(monitorUrl) : null
    };
    if (args.waitForCompletion) {
      if (!monitorUrl) {
        result.monitor = {
          complete: false,
          terminal: false,
          succeeded: false,
          terminalState: "missing_monitor_url"
        };
      } else {
        try {
          result.monitor = await pollCopyMonitor(monitorUrl, args.timeoutSeconds ?? 60);
        } catch (error) {
          result.monitorError = safeToolErrorMessage(error);
          result.monitor = {
            complete: false,
            terminal: false,
            succeeded: false,
            terminalState: "monitor_error"
          };
        }
      }
    }
    await writeMutationAudit("onedrive_copy", {
      status: result.monitorError
        ? "accepted-monitor-failed"
        : result.monitor?.terminalState === "missing_monitor_url"
          ? "accepted-monitor-unavailable"
        : result.monitor?.terminal && !result.monitor?.succeeded
          ? "accepted-copy-failed"
          : result.monitor?.succeeded
            ? "success"
            : "accepted",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      destination: parentReference,
      newName: args.newName || null,
      monitorError: result.monitorError,
      graphRequestId: response.graphRequestId
    });
    return result;
  } catch (error) {
    await writeMutationAudit("onedrive_copy", {
      status: "failed",
      target: itemAuditSummary(current),
      before: itemAuditSummary(current),
      destination: parentReference,
      newName: args.newName || null,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function createSharingLink(args = {}) {
  requireNonRootTarget(args, "Create sharing link");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Create sharing link refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Create sharing link");
  const includePermissionDiff = args.includePermissionDiff !== false;
  const beforePermissions = includePermissionDiff ? await permissionList({ itemId: current.id }, "compact") : null;
  const warnings = [];
  const requestedScope = args.scope || "anonymous";
  if (requestedScope === "organization") {
    try {
      const driveInfo = await graph("/me/drive");
      if (driveInfo.driveType === "personal") {
        warnings.push("Organization-scoped links are not meaningful on personal OneDrive drives; Microsoft Graph may reject or reinterpret this scope.");
      }
    } catch {
      warnings.push("Could not verify drive type before previewing organization-scoped link creation.");
    }
  }
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldCreate: {
      item: simplifyItem(current),
      type: args.type || "view",
      scope: requestedScope,
      passwordProvided: Boolean(args.password),
      expirationDateTime: args.expirationDateTime,
      retainInheritedPermissions: args.retainInheritedPermissions
    },
    warnings,
    ...(includePermissionDiff ? {
      beforePermissions,
      beforePermissionCount: beforePermissions.length
    } : {})
  };
  const previewProof = {
    item: { id: current.id, name: current.name },
    type: args.type || "view",
    scope: requestedScope,
    password: args.password,
    expirationDateTime: args.expirationDateTime,
    retainInheritedPermissions: args.retainInheritedPermissions
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_create_sharing_link", previewProof),
      requiredToCreate: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToCreate: "Provide expectedName or expectedId for live sharing-link creation."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_create_sharing_link", previewProof, args.previewToken, "requiredToCreate");
  if (previewTokenRequired) return previewTokenRequired;
  const body = {
    type: args.type || "view",
    scope: args.scope || "anonymous"
  };
  if (args.password) body.password = args.password;
  if (args.expirationDateTime) body.expirationDateTime = args.expirationDateTime;
  if (typeof args.retainInheritedPermissions === "boolean") {
    body.retainInheritedPermissions = args.retainInheritedPermissions;
  }
  try {
    const result = await graph(`${itemMutationBase(current)}/createLink`, {
      method: "POST",
      body: JSON.stringify(body)
    });
    const afterPermissions = includePermissionDiff
      ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: current.id }, "compact"))
      : null;
    const permissionDiff = includePermissionDiff && afterPermissions ? diffPermissions(beforePermissions, afterPermissions) : null;
    await writeMutationAudit("onedrive_create_sharing_link", {
      status: "success",
      target: itemAuditSummary(current),
      link: {
        type: body.type,
        scope: body.scope,
        passwordProvided: Boolean(args.password),
        expirationDateTime: args.expirationDateTime,
        retainInheritedPermissions: body.retainInheritedPermissions
      },
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      afterPermissions: includePermissionDiff && afterPermissions ? afterPermissions.map(permissionAuditSummary) : undefined,
      permissionDiff: includePermissionDiff && permissionDiff ? permissionDiffAuditSummary(permissionDiff) : undefined
    });
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(current),
      permission: result,
      verificationIncomplete: includePermissionDiff && !afterPermissions,
      ...(includePermissionDiff ? {
        beforePermissions,
        afterPermissions,
        permissionDiff
      } : {})
    };
  } catch (error) {
    await writeMutationAudit("onedrive_create_sharing_link", {
      status: "failed",
      target: itemAuditSummary(current),
      link: {
        type: body.type,
        scope: body.scope,
        passwordProvided: Boolean(args.password),
        expirationDateTime: args.expirationDateTime,
        retainInheritedPermissions: body.retainInheritedPermissions
      },
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function normalizeInviteRecipients(recipients = []) {
  return recipients.map((recipient, index) => {
    const keys = ["email", "alias", "objectId"].filter((key) => recipient[key]);
    if (keys.length !== 1) {
      throw new Error(`Invite recipient at index ${index} must include exactly one of email, alias, or objectId.`);
    }
    return { [keys[0]]: recipient[keys[0]] };
  });
}

function recipientKinds(recipients = []) {
  return recipients.map((recipient) => ["email", "alias", "objectId"].find((key) => recipient[key]) || "unknown");
}

function inviteBody(args = {}) {
  const body = {
    recipients: normalizeInviteRecipients(args.recipients || []),
    roles: [args.role || "read"],
    sendInvitation: args.sendInvitation === true,
    requireSignIn: args.requireSignIn !== false
  };
  if (args.message) body.message = args.message;
  if (args.password) body.password = args.password;
  if (args.expirationDateTime) body.expirationDateTime = args.expirationDateTime;
  if (typeof args.retainInheritedPermissions === "boolean") {
    body.retainInheritedPermissions = args.retainInheritedPermissions;
  }
  return body;
}

function inviteAuditSummary(invite = {}) {
  return {
    recipientCount: invite.recipientCount,
    recipientKinds: invite.recipientKinds,
    role: invite.role,
    sendInvitation: invite.sendInvitation,
    requireSignIn: invite.requireSignIn,
    messageProvided: invite.messageProvided,
    passwordProvided: invite.passwordProvided,
    expirationDateTime: invite.expirationDateTime,
    retainInheritedPermissions: invite.retainInheritedPermissions
  };
}

async function invitePermission(args = {}) {
  requireNonRootTarget(args, "Invite permission");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Invite permission refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Invite permission");
  const body = inviteBody(args);
  const includePermissionDiff = args.includePermissionDiff !== false;
  const beforePermissions = includePermissionDiff ? await permissionList({ itemId: current.id }, "compact") : null;
  const safeInvite = {
    item: simplifyItem(current),
    recipientCount: body.recipients.length,
    recipientKinds: recipientKinds(body.recipients),
    role: body.roles[0],
    sendInvitation: body.sendInvitation,
    requireSignIn: body.requireSignIn,
    messageProvided: Boolean(args.message),
    passwordProvided: Boolean(args.password),
    expirationDateTime: args.expirationDateTime,
    retainInheritedPermissions: args.retainInheritedPermissions
  };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldInvite: safeInvite,
    ...(includePermissionDiff ? {
      beforePermissions,
      beforePermissionCount: beforePermissions.length
    } : {})
  };
  const previewProof = {
    item: { id: current.id, name: current.name },
    recipients: body.recipients,
    role: body.roles[0],
    sendInvitation: body.sendInvitation,
    requireSignIn: body.requireSignIn,
    message: args.message,
    password: args.password,
    expirationDateTime: args.expirationDateTime,
    retainInheritedPermissions: args.retainInheritedPermissions
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_invite_permission", previewProof),
      requiredToInvite: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToInvite: "Provide expectedName or expectedId for live permission invitation."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_invite_permission", previewProof, args.previewToken, "requiredToInvite");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const result = await graph(`${itemMutationBase(current)}/invite`, {
      method: "POST",
      body: JSON.stringify(body)
    });
    const afterPermissions = includePermissionDiff
      ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: current.id }, "compact"))
      : null;
    const permissionDiff = includePermissionDiff && afterPermissions ? diffPermissions(beforePermissions, afterPermissions) : null;
    await writeMutationAudit("onedrive_invite_permission", {
      status: "success",
      target: itemAuditSummary(current),
      invite: inviteAuditSummary(safeInvite),
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      afterPermissions: includePermissionDiff && afterPermissions ? afterPermissions.map(permissionAuditSummary) : undefined,
      permissionDiff: includePermissionDiff && permissionDiff ? permissionDiffAuditSummary(permissionDiff) : undefined
    });
    return {
      dryRun: false,
      confirmed: true,
      item: simplifyItem(current),
      invite: safeInvite,
      permissions: Array.isArray(result.value) ? result.value.map((permission) => simplifyPermission(permission, args.format)) : result,
      verificationIncomplete: includePermissionDiff && !afterPermissions,
      ...(includePermissionDiff ? {
        beforePermissions,
        afterPermissions,
        permissionDiff
      } : {})
    };
  } catch (error) {
    await writeMutationAudit("onedrive_invite_permission", {
      status: "failed",
      target: itemAuditSummary(current),
      invite: inviteAuditSummary(safeInvite),
      beforePermissions: includePermissionDiff ? beforePermissions.map(permissionAuditSummary) : undefined,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function assertPermissionPresent(permissions = [], permissionId) {
  if (!permissions.some((permission) => permission.id === permissionId)) {
    throw new Error(`Permission ${permissionId} was not found on the target item. Refusing to continue.`);
  }
}

function findPermissionById(permissions = [], permissionId) {
  return permissions.find((permission) => permission.id === permissionId) || null;
}

async function preflightRevokePermission(args = {}, options = {}) {
  requireNonRootTarget(args, "Revoke permission");
  const current = await getRawInfo(args);
  if (current.root) throw new Error("Revoke permission refuses to operate on the OneDrive root.");
  assertExpectedItem(current, args, "Revoke permission");
  const includePermissions = options.includePermissions !== false;
  const permissionsForPreflight = await permissionList({ itemId: current.id }, "compact");
  assertPermissionPresent(permissionsForPreflight, args.permissionId);
  const permission = findPermissionById(permissionsForPreflight, args.permissionId);
  const revocable = isRevocablePermission(permission);
  const warnings = revocable ? [] : [`Permission ${args.permissionId} is a ${permission?.permissionKind || "non-revocable"} permission and should not be revoked with this tool.`];
  return {
    targetArgs: args,
    rawItem: current,
    item: simplifyItem(current),
    permission,
    revocable,
    warnings,
    beforePermissions: includePermissions ? permissionsForPreflight : [],
    includePermissions
  };
}

async function revokePermission(args = {}) {
  const preflight = await preflightRevokePermission(args, { includePermissions: args.includePermissions });
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldRevoke: {
      item: preflight.item,
      permissionId: args.permissionId,
      permission: preflight.permission,
      revocable: preflight.revocable
    },
    warnings: preflight.warnings,
    ...(preflight.includePermissions ? {
      beforePermissions: preflight.beforePermissions,
      beforePermissionCount: preflight.beforePermissions.length
    } : {})
  };
  const previewProof = {
    item: { id: preflight.rawItem.id, name: preflight.rawItem.name },
    permissionId: args.permissionId
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_revoke_permission", previewProof),
      requiredToRevoke: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToRevoke: "Provide expectedName or expectedId for live permission revocation."
    };
  }
  if (!preflight.revocable) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToRevoke: "Choose a revocable sharing permission. Owner and inherited permissions are refused."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_revoke_permission", previewProof, args.previewToken, "requiredToRevoke");
  if (previewTokenRequired) return previewTokenRequired;
  try {
    const response = await graph(`${itemIdBase(preflight.rawItem.id)}/permissions/${encodeURIComponent(args.permissionId)}`, {
      method: "DELETE",
      returnResponse: true
    });
    if (!response.ok) throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Revoke permission failed" });
    const afterPermissions = preflight.includePermissions
      ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: preflight.rawItem.id }, "compact"))
      : null;
    const permissionDiff = preflight.includePermissions && afterPermissions ? diffPermissions(preflight.beforePermissions, afterPermissions) : null;
    await writeMutationAudit("onedrive_revoke_permission", {
      status: "success",
      target: itemAuditSummary(preflight.rawItem),
      permissionId: args.permissionId,
      beforePermissions: preflight.includePermissions ? preflight.beforePermissions.map(permissionAuditSummary) : undefined,
      afterPermissions: preflight.includePermissions && afterPermissions ? afterPermissions.map(permissionAuditSummary) : undefined,
      permissionDiff: preflight.includePermissions && permissionDiff ? permissionDiffAuditSummary(permissionDiff) : undefined,
      graphRequestId: response.graphRequestId
    });
    return {
      dryRun: false,
      confirmed: true,
      item: preflight.item,
      permissionId: args.permissionId,
      verificationIncomplete: preflight.includePermissions && !afterPermissions,
      ...(preflight.includePermissions ? {
        beforePermissions: preflight.beforePermissions,
        afterPermissions,
        permissionDiff
      } : {})
    };
  } catch (error) {
    await writeMutationAudit("onedrive_revoke_permission", {
      status: "failed",
      target: itemAuditSummary(preflight.rawItem),
      permissionId: args.permissionId,
      beforePermissions: preflight.includePermissions ? preflight.beforePermissions.map(permissionAuditSummary) : undefined,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function batchRevokePermissions(args = {}) {
  const items = args.items || [];
  const warnings = batchMutationWarnings();
  if (args.dryRun === false) {
    if (args.confirmed !== true) {
      return {
        dryRun: false,
        confirmed: false,
        count: items.length,
        warnings,
        requiredToRevoke: "Set dryRun: false and confirmed: true after explicit user confirmation."
      };
    }
    const missingExpected = items.filter((item) => !hasExpectedIdentity(item));
    if (missingExpected.length) {
      return {
        dryRun: false,
        confirmed: true,
        count: items.length,
        warnings,
        requiredToRevoke: "Provide expectedName or expectedId for every item in a live batch permission revoke.",
        missingExpectedCount: missingExpected.length
      };
    }
  }

  const preflight = [];
  const preflightErrors = [];
  const resolvedPermissionTargets = new Set();
  for (const [index, item] of items.entries()) {
    try {
      const entry = await preflightRevokePermission(item, { includePermissions: args.includePermissions });
      const key = `${entry.rawItem.id}\u0000${entry.targetArgs.permissionId}`;
      if (resolvedPermissionTargets.has(key)) throw new Error(`Batch revoke target duplicates item ${entry.rawItem.id} permission ${entry.targetArgs.permissionId}.`);
      resolvedPermissionTargets.add(key);
      preflight.push(entry);
    } catch (error) {
      preflightErrors.push({ index, target: item, error: safeToolErrorMessage(error) });
    }
  }
  if (preflightErrors.length) {
    return {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: preflightErrors,
      requiredToRevoke: "Fix every preflight error before running a batch permission revoke."
    };
  }
  const nonRevocable = preflight
    .map((entry, index) => entry.revocable ? null : { index, target: entry.targetArgs, permission: entry.permission, warnings: entry.warnings })
    .filter(Boolean);
  if (args.dryRun === false && nonRevocable.length) {
    return {
      dryRun: false,
      confirmed: args.confirmed === true,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: nonRevocable.map((entry) => ({
        index: entry.index,
        target: entry.target,
        error: entry.warnings.join(" ")
      })),
      requiredToRevoke: "Choose only revocable sharing permissions. Owner and inherited permissions are refused."
    };
  }
  const previewProof = {
    items: preflight.map((entry) => ({
      id: entry.rawItem.id,
      name: entry.rawItem.name,
      permissionId: entry.targetArgs.permissionId
    })),
    operation: "batch-revoke"
  };

  if (args.dryRun !== false) {
    return {
      dryRun: true,
      confirmed: args.confirmed === true,
      count: preflight.length,
      warnings,
      ...issuePreviewToken("onedrive_batch_revoke_permissions", previewProof),
      results: preflight.map((entry) => ({
        wouldRevoke: { item: entry.item, permissionId: entry.targetArgs.permissionId, permission: entry.permission, revocable: entry.revocable },
        warnings: entry.warnings,
        ...(entry.includePermissions ? {
          beforePermissions: entry.beforePermissions,
          beforePermissionCount: entry.beforePermissions.length
        } : {})
      }))
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(
    { dryRun: false, confirmed: true, count: preflight.length, warnings },
    "onedrive_batch_revoke_permissions",
    previewProof,
    args.previewToken,
    "requiredToRevoke"
  );
  if (previewTokenRequired) return previewTokenRequired;

  const results = [];
  for (const [index, entry] of preflight.entries()) {
    try {
      const response = await graph(`${itemIdBase(entry.rawItem.id)}/permissions/${encodeURIComponent(entry.targetArgs.permissionId)}`, {
        method: "DELETE",
        returnResponse: true
      });
      if (!response.ok) throw microsoftGraphError(response.body, { headers: response.headers, status: response.status, statusText: "Batch revoke permission failed" });
      const afterPermissions = entry.includePermissions
        ? await bestEffortLocalWrite("post-mutation permission verification", async () => await permissionList({ itemId: entry.rawItem.id }, "compact"))
        : null;
      const permissionDiff = entry.includePermissions && afterPermissions ? diffPermissions(entry.beforePermissions, afterPermissions) : null;
      results.push({
        item: entry.item,
        permissionId: entry.targetArgs.permissionId,
        verificationIncomplete: entry.includePermissions && !afterPermissions,
        ...(entry.includePermissions ? { beforePermissions: entry.beforePermissions, afterPermissions, permissionDiff } : {})
      });
    } catch (error) {
      await writeMutationAudit("onedrive_batch_revoke_permissions", {
        status: "failed",
        targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
        partialResults: results.map((result) => ({ item: itemAuditSummary(result.item), permissionId: result.permissionId })),
        failedIndex: index,
        error: safeErrorInfo(error)
      });
      return {
        dryRun: false,
        confirmed: true,
        failed: true,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        count: preflight.length,
        warnings,
        partialResults: results
      };
    }
  }
  await writeMutationAudit("onedrive_batch_revoke_permissions", {
    status: "success",
    targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
    results: results.map((result) => ({
      item: itemAuditSummary(result.item),
      permissionId: result.permissionId,
      permissionDiff: result.permissionDiff ? permissionDiffAuditSummary(result.permissionDiff) : undefined
    }))
  });
  return { dryRun: false, confirmed: true, count: results.length, warnings, results };
}

async function preflightMoveItem(item = {}, destination) {
  if (item.newName) assertSafeItemName(item.newName, "newName");
  requireNonRootTarget(item, "Move");
  const current = await getRawInfo(item);
  if (current.root) throw new Error("Move refuses to operate on the OneDrive root.");
  assertExpectedItem(current, item, "Move");
  return { targetArgs: item, rawItem: current, item: simplifyItem(current), destination };
}

async function batchMove(args = {}) {
  const items = args.items || [];
  const warnings = batchMutationWarnings();
  if (args.dryRun === false) {
    if (args.confirmed !== true) {
      return {
        dryRun: false,
        confirmed: false,
        count: items.length,
        warnings,
        requiredToMove: "Set dryRun: false and confirmed: true after explicit user confirmation."
      };
    }
    const missingExpected = items.filter((item) => !hasExpectedIdentity(item));
    if (missingExpected.length) {
      return {
        dryRun: false,
        confirmed: true,
        count: items.length,
        warnings,
        requiredToMove: "Provide expectedName or expectedId for every item in a live batch move.",
        missingExpectedCount: missingExpected.length
      };
    }
  }

  const destination = await resolveDestinationParent(args);
  const preflight = [];
  const preflightErrors = [];
  const resolvedIds = new Set();
  for (const [index, item] of items.entries()) {
    try {
      const entry = await preflightMoveItem(item, destination);
      if (resolvedIds.has(entry.rawItem.id)) throw new Error(`Batch move target resolves to duplicate item ID ${entry.rawItem.id}.`);
      resolvedIds.add(entry.rawItem.id);
      preflight.push(entry);
    } catch (error) {
      preflightErrors.push({ index, target: item, error: safeToolErrorMessage(error) });
    }
  }
  if (preflightErrors.length) {
    return {
      dryRun: args.dryRun !== false,
      confirmed: args.confirmed === true,
      destination,
      count: items.length,
      warnings,
      preflightFailed: true,
      errors: preflightErrors,
      requiredToMove: "Fix every preflight error before running a batch move."
    };
  }

  if (args.dryRun !== false) {
    return {
      dryRun: true,
      confirmed: args.confirmed === true,
      destination,
      count: preflight.length,
      warnings,
      results: preflight.map((entry) => ({
        wouldMove: entry.item,
        destination,
        newName: entry.targetArgs.newName || null
      }))
    };
  }

  const results = [];
  for (const [index, entry] of preflight.entries()) {
    try {
      const body = { parentReference: { id: destination.id } };
      if (entry.targetArgs.newName) body.name = entry.targetArgs.newName;
      const result = await graph(itemIdBase(entry.rawItem.id), {
        method: "PATCH",
        headers: mutationMatchHeaders(entry.rawItem),
        body: JSON.stringify(body)
      });
      await cacheMovedOrRenamedItem(entry.rawItem, result);
      results.push({ before: entry.item, moved: simplifyItem(result), newName: entry.targetArgs.newName || null });
    } catch (error) {
      await writeMutationAudit("onedrive_batch_move", {
        status: "failed",
        destination,
        targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
        partialResults: results.map((result) => ({ before: itemAuditSummary(result.before), after: itemAuditSummary(result.moved) })),
        failedIndex: index,
        error: safeErrorInfo(error)
      });
      return {
        dryRun: false,
        confirmed: true,
        failed: true,
        failedIndex: index,
        error: safeToolErrorMessage(error),
        destination,
        count: preflight.length,
        warnings,
        partialResults: results
      };
    }
  }
  await writeMutationAudit("onedrive_batch_move", {
    status: "success",
    destination,
    targets: preflight.map((entry) => itemAuditSummary(entry.rawItem)),
    results: results.map((result) => ({
      before: itemAuditSummary(result.before),
      after: itemAuditSummary(result.moved)
    }))
  });
  return { dryRun: false, confirmed: true, destination, count: results.length, warnings, results };
}

async function deleteItem(args = {}) {
  requireNonRootTarget(args, "Delete");
  const rawItem = await getRawInfo(args);
  if (rawItem.root) throw new Error("Delete refuses to operate on the OneDrive root.");
  assertExpectedItem(rawItem, args, "Delete");
  const item = simplifyItem(rawItem);
  const previewProof = { items: [itemVersionProof(rawItem)], operation: "delete" };
  if (args.dryRun !== false) {
    return previewWithToken({ dryRun: true, wouldDelete: item }, "onedrive_delete", previewProof);
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
  const previewTokenRequired = previewTokenRequiredResult(
    { dryRun: false, confirmed: true, wouldDelete: item },
    "onedrive_delete",
    previewProof,
    args.previewToken,
    "requiredToDelete"
  );
  if (previewTokenRequired) return previewTokenRequired;
  try {
    await graph(itemMutationBase(rawItem), { method: "DELETE", headers: mutationMatchHeaders(rawItem) });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([{ ...rawItem, deleted: {} }]));
    await writeMutationAudit("onedrive_delete", {
      status: "success",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem)
    });
    return { dryRun: false, confirmed: true, deleted: item };
  } catch (error) {
    await writeMutationAudit("onedrive_delete", {
      status: "failed",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem),
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

async function permanentDeleteItem(args = {}) {
  const rawItem = await getRawInfo({ itemId: args.itemId });
  if (rawItem.root) throw new Error("Permanent delete refuses to operate on the OneDrive root.");
  assertExpectedItem(rawItem, args, "Permanent delete");
  const item = simplifyItem(rawItem);
  const previewProof = { items: [itemVersionProof(rawItem)], operation: "permanentDelete" };
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    acknowledgeIrreversible: args.acknowledgeIrreversible === true,
    wouldPermanentlyDelete: item,
    warning: "This bypasses the recycle bin and cannot be undone."
  };
  if (args.dryRun !== false) return previewWithToken(preview, "onedrive_permanent_delete", previewProof);
  if (args.confirmed !== true || args.acknowledgeIrreversible !== true) {
    return {
      ...preview,
      dryRun: false,
      requiredToPermanentlyDelete: "Set dryRun: false, confirmed: true, and acknowledgeIrreversible: true after explicit user confirmation."
    };
  }
  if (!hasExpectedIdentity(args)) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToPermanentlyDelete: "Provide expectedName or expectedId for permanent deletion."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(
    preview,
    "onedrive_permanent_delete",
    previewProof,
    args.previewToken,
    "requiredToPermanentlyDelete"
  );
  if (previewTokenRequired) return previewTokenRequired;
  const driveId = rawItem.parentReference?.driveId || (await graph("/me/drive?$select=id")).id;
  if (!driveId) throw new Error("Could not resolve the OneDrive drive ID for permanent deletion.");
  try {
    await graph(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(rawItem.id)}/permanentDelete`, {
      method: "POST",
      headers: mutationMatchHeaders(rawItem),
      maxRetries: 0
    });
    await bestEffortLocalWrite("metadata cache update", async () => await cacheItems([{ ...rawItem, deleted: { state: "permanentlyDeleted" } }]));
    await writeMutationAudit("onedrive_permanent_delete", {
      status: "success",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem),
      irreversible: true
    });
    return { dryRun: false, confirmed: true, permanentlyDeleted: item, irreversible: true };
  } catch (error) {
    await writeMutationAudit("onedrive_permanent_delete", {
      status: "failed",
      target: itemAuditSummary(rawItem),
      before: itemAuditSummary(rawItem),
      irreversible: true,
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

function validateRestoreArgs(args = {}) {
  if (args.newName) assertSafeItemName(args.newName, "newName");
  if (args.destinationParentPath) assertSafeRemotePath(args.destinationParentPath, "destinationParentPath");
  if (args.destinationParentRelativePath) assertSafeRemotePath(args.destinationParentRelativePath, "destinationParentRelativePath");
  if (args.destinationParentPreset) {
    resolvePresetPath(args, {
      pathField: "destinationParentPath",
      presetField: "destinationParentPreset",
      relativeField: "destinationParentRelativePath"
    });
  }
}

async function restoreDeleted(args = {}) {
  validateRestoreArgs(args);
  if (args.expectedId && args.expectedId !== args.itemId) {
    throw new Error(`Restore expected item ID ${args.expectedId}, but got ${args.itemId}. Refusing to continue.`);
  }
  const hasExplicitDestination = Boolean(args.destinationParentPath || args.destinationParentItemId || args.destinationParentPreset);
  const destinationParent = hasExplicitDestination ? await resolveDestinationParent(args) : null;
  const preview = {
    dryRun: args.dryRun !== false,
    confirmed: args.confirmed === true,
    wouldRestore: {
      itemId: args.itemId,
      destinationParentPath: args.destinationParentPath,
      destinationParentItemId: args.destinationParentItemId,
      destinationParentPreset: args.destinationParentPreset,
      destinationParentRelativePath: args.destinationParentRelativePath,
      resolvedDestinationParent: destinationParent ? simplifyItem(destinationParent) : null,
      newName: args.newName
    },
    permissionNote: "OneDrive Personal restore may require Microsoft Graph delegated Files.ReadWrite.All in addition to the plugin's standard Files.ReadWrite scope."
  };
  const previewProof = {
    itemId: args.itemId,
    destinationParentPath: args.destinationParentPath,
    destinationParentItemId: args.destinationParentItemId,
    destinationParentPreset: args.destinationParentPreset,
    destinationParentRelativePath: args.destinationParentRelativePath,
    destinationParentId: destinationParent?.id || null,
    newName: args.newName
  };
  if (args.dryRun !== false || args.confirmed !== true) {
    return {
      ...previewWithToken(preview, "onedrive_restore_deleted", previewProof),
      requiredToRestore: "Set dryRun: false and confirmed: true after explicit user confirmation."
    };
  }
  if (!args.expectedId) {
    return {
      ...preview,
      dryRun: false,
      confirmed: true,
      requiredToRestore: "Provide expectedId for live restores."
    };
  }
  const previewTokenRequired = previewTokenRequiredResult(preview, "onedrive_restore_deleted", previewProof, args.previewToken, "requiredToRestore");
  if (previewTokenRequired) return previewTokenRequired;
  const body = {};
  if (destinationParent) body.parentReference = { id: destinationParent.id };
  if (args.newName) body.name = args.newName;
  try {
    const result = await graph(`/me/drive/items/${encodeURIComponent(args.itemId)}/restore`, {
      method: "POST",
      body: JSON.stringify(body)
    });
    const restored = simplifyItem(result);
    await writeMutationAudit("onedrive_restore_deleted", {
      status: "success",
      target: { itemId: args.itemId },
      after: itemAuditSummary(restored)
    });
    return { dryRun: false, confirmed: true, restored };
  } catch (error) {
    await writeMutationAudit("onedrive_restore_deleted", {
      status: "failed",
      target: { itemId: args.itemId },
      error: safeErrorInfo(error)
    });
    throw error;
  }
}

export function boundChatgptToolPayload(value, maxBytes = chatgptToolResponseByteLimit) {
  let serialized;
  try {
    serialized = JSON.stringify(value);
  } catch (error) {
    const replacement = {
      responseTruncated: true,
      originalBytes: null,
      maxBytes,
      retryHint: "Retry with a smaller result limit or a more specific item ID.",
      preview: safeToolErrorMessage(error)
    };
    return { value: replacement, truncated: true, originalBytes: null, boundedBytes: Buffer.byteLength(JSON.stringify(replacement), "utf8") };
  }
  if (serialized === undefined) serialized = "null";
  const originalBytes = Buffer.byteLength(serialized, "utf8");
  if (originalBytes <= maxBytes) {
    return { value, truncated: false, originalBytes, boundedBytes: originalBytes };
  }
  const previewLimit = Math.max(1024, Math.min(256 * 1024, Math.floor(maxBytes / 2)));
  const replacement = {
    responseTruncated: true,
    originalBytes,
    maxBytes,
    retryHint: "Retry with a smaller result limit, a narrower range, or a more specific item ID.",
    preview: truncateUtf8(serialized, previewLimit).text
  };
  const boundedBytes = Buffer.byteLength(JSON.stringify(replacement), "utf8");
  return { value: replacement, truncated: true, originalBytes, boundedBytes };
}

function textResult(value, isError = false, structuredContent = undefined) {
  const localWarnings = toolCallContext.getStore()?.localWarnings || [];
  let payload = localWarnings.length
    ? (value && typeof value === "object" && !Array.isArray(value)
        ? { ...value, localWarnings }
        : { message: value, localWarnings })
    : value;
  let boundedStructuredContent = structuredContent;
  if (toolProfile === "chatgpt") {
    const boundedPayload = boundChatgptToolPayload(payload);
    payload = boundedPayload.value;
    if (structuredContent !== undefined) {
      boundedStructuredContent = structuredContent === value && !localWarnings.length
        ? boundedPayload.value
        : boundChatgptToolPayload(structuredContent).value;
    }
    if (boundedPayload.truncated) {
      console.error(JSON.stringify({
        event: "onedrive-response-truncated",
        tool: toolCallContext.getStore()?.toolName || null,
        originalBytes: boundedPayload.originalBytes,
        boundedBytes: boundedPayload.boundedBytes,
        maxBytes: chatgptToolResponseByteLimit
      }));
    }
  }
  const text = typeof payload === "string" ? payload : JSON.stringify(payload, null, toolProfile === "chatgpt" ? 0 : 2);
  return {
    content: [{ type: "text", text }],
    ...(boundedStructuredContent === undefined ? {} : { structuredContent: boundedStructuredContent }),
    isError
  };
}

function structuredToolError(error, tool = undefined) {
  const message = safeToolErrorMessage(error);
  const graphStatus = Number.isInteger(error?.graphStatus) ? error.graphStatus : undefined;
  let code = "internal_error";
  if (graphStatus === 400) code = "invalid_argument";
  else if (graphStatus === 401) code = "unauthenticated";
  else if (graphStatus === 403) code = "permission_denied";
  else if (graphStatus === 404) code = "not_found";
  else if (graphStatus === 409 || graphStatus === 412) code = "conflict";
  else if (graphStatus === 429) code = "rate_limited";
  else if (graphStatus >= 500) code = "service_unavailable";
  else if (/\b(?:timed out|timeout|AbortError|TimeoutError)\b/i.test(message)) code = "deadline_exceeded";
  else if (/\b(?:refus|requires|invalid|must include|must provide|unsupported|not a folder|not a file)\b/i.test(message)) code = "failed_precondition";
  return {
    error: {
      code,
      message,
      ...(tool ? { tool } : {}),
      ...(graphStatus === undefined ? {} : { graphStatus }),
      ...(error?.graphRequestId ? { graphRequestId: redactAuditText(error.graphRequestId) } : {})
    }
  };
}

function toolErrorResult(error, tool = undefined) {
  const structuredContent = structuredToolError(error, tool);
  return textResult(structuredContent.error.message, true, structuredContent);
}

function resultMessage(id, result) {
  return { jsonrpc: "2.0", id, result };
}

function errorMessage(id, code, message) {
  return { jsonrpc: "2.0", id, error: { code, message } };
}

function oauthRequiredResult(error = null) {
  const description = error?.message || "Connect OneDrive to continue.";
  return {
    ...textResult(description, true, {
      error: {
        code: error?.code === "insufficient_scope" ? "permission_denied" : "unauthenticated",
        message: description
      }
    }),
    _meta: {
      "mcp/www_authenticate": [oauthChallenge({
        error: error?.code || "invalid_token",
        description
      })]
    }
  };
}

async function callTool(name, args = {}) {
  if (toolCallContext.getStore()?.authMode === "oauth"
    && new Set(["onedrive_auth_device_start", "onedrive_auth_device_poll", "onedrive_logout"]).has(name)) {
    return textResult(
      "Device-code credential tools are disabled for delegated OAuth requests. Connect or disconnect OneDrive from ChatGPT app settings.",
      true,
      {
        error: {
          code: "failed_precondition",
          message: "Manage the delegated OneDrive connection from ChatGPT app settings.",
          tool: name
        }
      }
    );
  }
  if (previewScopedTools.has(name)) {
    const store = toolCallContext.getStore();
    if (store) store.storageScope = await activeStorageScope();
  }
  switch (name) {
    case "search": {
      const value = await chatgptSearch(args);
      return textResult(value, false, value);
    }
    case "fetch": {
      const value = await chatgptFetch(args);
      return textResult(value, false, value);
    }
    case "onedrive_open_files": {
      const value = await chatgptOpenFiles(args);
      return textResult(value, false, value);
    }
    case "onedrive_preview_actions": {
      const value = await chatgptPreviewActions(args);
      return textResult(value, false, value);
    }
    case "onedrive_config": {
      const status = publicConfig();
      if (args.checkToken) {
        try {
          await getAccessToken();
          status.accessTokenAvailable = true;
        } catch (error) {
          status.accessTokenAvailable = false;
          status.tokenCheckError = safeToolErrorMessage(error);
        }
      }
      return textResult(status);
    }
    case "onedrive_auth_device_start":
      return textResult(await startDeviceLogin(args));
    case "onedrive_auth_device_poll":
      return textResult(await pollDeviceLogin(args));
    case "onedrive_logout": {
      authGeneration += 1;
      deviceLoginGeneration += 1;
      invalidateActiveStorageScope();
      tokenCache = null;
      pendingDevice = null;
      tokenRefreshPromise = null;
      adoptCurrentToolAccountGeneration();
      if (args.deleteKeychainToken === true && args.confirmed !== true) {
        return textResult({
          memoryCleared: true,
          storedCredentialDeleted: false,
          keychainTokenDeleted: false,
          confirmed: false,
          requiredToDelete: "Set confirmed: true after explicit user confirmation to delete the securely stored OneDrive refresh token."
        });
      }
      const storedCredentialDeleted = args.deleteKeychainToken ? deleteKeychainToken() : false;
      return textResult({
        memoryCleared: true,
        storedCredentialDeleted,
        keychainTokenDeleted: authVault(config()).mode === "keychain" && storedCredentialDeleted
      });
    }
    case "onedrive_me":
      return textResult(await graph("/me"));
    case "onedrive_drive":
      return textResult(await graph("/me/drive"));
    case "onedrive_doctor":
      return textResult(await doctor(args));
    case "onedrive_presets":
      return textResult({ pathPresets: pathPresets(), configPath });
    case "onedrive_list":
      return textResult(await list(args));
    case "onedrive_list_all":
      return textResult(await listAll(args));
    case "onedrive_scan":
      return textResult(await scan(args));
    case "onedrive_search":
      return textResult(await search(args));
    case "onedrive_search_all":
      return textResult(await searchAll(args));
    case "onedrive_find":
      return textResult(await find(args));
    case "onedrive_find_all":
      return textResult(await findAll(args));
    case "onedrive_delta":
      return textResult(await delta(args));
    case "onedrive_sync_status":
      return textResult(await syncStatus(args));
    case "onedrive_cache_refresh":
      return textResult(await cacheRefresh(args));
    case "onedrive_cache_clear":
      return textResult(await clearMetadataCache());
    case "onedrive_content_index_refresh":
      return textResult(await contentIndexRefresh(args));
    case "onedrive_content_search":
      return textResult(await contentSearch(args));
    case "onedrive_office_index_refresh":
      return textResult(await officeIndexRefresh(args));
    case "onedrive_office_search":
      return textResult(await officeContentSearch(args));
    case "onedrive_content_index_clear":
      return textResult(await clearContentIndex());
    case "onedrive_office_capabilities":
      return textResult(await officeCapabilities());
    case "onedrive_office_validate":
      return textResult(await inspectRemoteOfficePackage(args, args.expectedKind || null, "validate"));
    case "onedrive_word_get_document":
      return textResult(await inspectRemoteOfficePackage(args, "word"));
    case "onedrive_excel_get_workbook":
      return textResult(await inspectRemoteOfficePackage(args, "excel"));
    case "onedrive_powerpoint_get_presentation":
      return textResult(await inspectRemoteOfficePackage(args, "powerpoint"));
    case "onedrive_word_batch_update":
      return textResult(await officeBatchUpdate(args, "word", "onedrive_word_batch_update"));
    case "onedrive_excel_batch_update":
      return textResult(await officeBatchUpdate(args, "excel", "onedrive_excel_batch_update"));
    case "onedrive_powerpoint_batch_update":
      return textResult(await officeBatchUpdate(args, "powerpoint", "onedrive_powerpoint_batch_update"));
    case "onedrive_office_backups":
      return textResult(await officeBackups(args));
    case "onedrive_office_compare_backup":
      return textResult(await officeCompareBackup(args));
    case "onedrive_office_restore_backup":
      return textResult(await officeRestoreBackup(args));
    case "onedrive_office_batch_transform":
      return textResult(await officeBatchTransform(args));
    case "onedrive_get_info":
      return textResult(await getInfo(args));
    case "onedrive_read_text":
      return textResult(await readText(args));
    case "onedrive_preview":
      return textResult(await preview(args));
    case "onedrive_download":
      return textResult(await download(args));
    case "onedrive_download_excel":
      return textResult(await downloadOffice(args, "excel"));
    case "onedrive_download_word":
      return textResult(await downloadOffice(args, "word"));
    case "onedrive_download_powerpoint":
      return textResult(await downloadOffice(args, "powerpoint"));
    case "onedrive_export_pdf":
      return textResult(await downloadExport(args, "pdf"));
    case "onedrive_export_text":
      return textResult(await downloadExport(args, "text"));
    case "onedrive_upload":
      return textResult(await upload(args));
    case "onedrive_upload_file":
      return textResult(await uploadChatgptFile(args));
    case "onedrive_write_text":
      return textResult(await writeText(args));
    case "onedrive_patch_text":
      return textResult(await patchText(args));
    case "onedrive_versions":
      return textResult(await versions(args));
    case "onedrive_compare_version":
      return textResult(await compareVersion(args));
    case "onedrive_restore_version":
      return textResult(await restoreVersion(args));
    case "onedrive_workspace_list":
      return textResult(await workspaceList());
    case "onedrive_workspace_create":
      return textResult(await workspaceCreate(args));
    case "onedrive_workspace_status":
      return textResult(await workspaceStatus(args));
    case "onedrive_workspace_promote":
      return textResult(await workspacePromote(args));
    case "onedrive_workspace_abandon":
      return textResult(await workspaceAbandon(args));
    case "onedrive_watch_start":
      return textResult(await watchStart(args));
    case "onedrive_watch_status":
      return textResult(await watchStatus(args));
    case "onedrive_watch_stop":
      return textResult(await watchStop(args));
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
    case "onedrive_invite_permission":
      return textResult(await invitePermission(args));
    case "onedrive_revoke_permission":
      return textResult(await revokePermission(args));
    case "onedrive_batch_revoke_permissions":
      return textResult(await batchRevokePermissions(args));
    case "onedrive_permissions":
      return textResult(await permissions(args));
    case "onedrive_batch_get_info":
      return textResult(await batchGetInfo(args));
    case "onedrive_batch_permissions":
      return textResult(await batchPermissions(args));
    case "onedrive_batch_download":
      return textResult(await batchDownload(args));
    case "onedrive_batch_delete":
      return textResult(await batchDelete(args));
    case "onedrive_batch_move":
      return textResult(await batchMove(args));
    case "onedrive_update_file":
      return textResult(await updateFile(args));
    case "onedrive_recent":
      return textResult(await recent(args));
    case "onedrive_large_files":
      return textResult(await largeFiles(args));
    case "onedrive_duplicates":
      return textResult(await duplicates(args));
    case "onedrive_shared_by_me":
      return textResult(await sharingAudit(args, false));
    case "onedrive_public_links":
      return textResult(await sharingAudit(args, true));
    case "onedrive_restore_deleted":
      return textResult(await restoreDeleted(args));
    case "onedrive_permanent_delete":
      return textResult(await permanentDeleteItem(args));
    case "onedrive_audit_recent":
      return textResult(await auditRecent(args));
    case "onedrive_audit_export":
      return textResult(await auditExport(args));
    case "onedrive_audit_clear":
      return textResult(await auditClear(args));
    case "onedrive_delete":
      return textResult(await deleteItem(args));
    default:
      return textResult(`Unknown tool: ${name}`, true);
  }
}

export async function processMcpMessage(message, requestAuth = null) {
  const { id, method, params = {} } = message;
  try {
    if (method === "initialize") {
      const requestedProtocolVersion = params.protocolVersion;
      const supportedProtocolVersions = new Set(["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05"]);
      if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
        console.error(JSON.stringify({ event: "onedrive-initialize", profile: toolProfile, serverVersion: advertisedServerVersion, toolCount: advertisedTools.length }));
      }
      return resultMessage(id, {
        protocolVersion: supportedProtocolVersions.has(requestedProtocolVersion) ? requestedProtocolVersion : "2024-11-05",
        capabilities: { tools: {} },
        instructions: serverInstructions,
        serverInfo: {
          name: "onedrive",
          title: "OneDrive",
          version: advertisedServerVersion,
          description: "Search, read, organize, share, and safely edit files in your personal OneDrive.",
          icons: [{
            src: chatgptIconDataUri,
            mimeType: "image/png",
            sizes: ["256x256"]
          }]
        }
      });
    }
    if (method === "tools/list") {
      if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
        console.error(JSON.stringify({ event: "onedrive-tools-list", profile: toolProfile, serverVersion: advertisedServerVersion, toolCount: advertisedTools.length }));
      }
      return resultMessage(id, { tools: advertisedTools });
    }
    if (method === "tools/call") {
      const toolStartedAt = performance.now();
      if (oauthSettings().mode === "oauth" && requestAuth?.authMode !== "oauth") {
        return resultMessage(id, oauthRequiredResult(requestAuth?.error));
      }
      const validation = validateToolArguments(params.name, params.arguments || {});
      if (!validation.ok) {
        return resultMessage(id, textResult(validation.error, true, {
          error: {
            code: "invalid_argument",
            message: `Invalid arguments for ${params.name}.`,
            tool: params.name,
            details: validation.error.details || []
          }
        }));
      }
      const result = await toolCallContext.run({
        toolName: params.name,
        localWarnings: [],
        lastGraphRequestId: null,
        lastMutationGraphRequestId: null,
        metadataCacheWrites: 0,
        authGeneration,
        storageScopeGeneration,
        ...(requestAuth || {})
      }, async () => {
        try {
          const value = await callTool(params.name, validation.args || {});
          assertToolAccountGeneration(`${params.name} completion`);
          return value;
        } catch (error) {
          return toolErrorResult(error, params.name);
        }
      });
      if (toolProfile === "chatgpt" || process.env.ONEDRIVE_PERFORMANCE_LOG === "1") {
        console.error(JSON.stringify({
          event: "onedrive-tool-complete",
          tool: params.name,
          durationMs: Number((performance.now() - toolStartedAt).toFixed(1)),
          isError: Boolean(result?.isError)
        }));
      }
      return resultMessage(id, result);
    }
    if (method?.startsWith("notifications/")) return null;
    if (id !== undefined) return errorMessage(id, -32601, `Method not found: ${method}`);
    return null;
  } catch (error) {
    if (id !== undefined) return resultMessage(id, toolErrorResult(error, params?.name));
    console.error(error);
    return null;
  }
}

export async function shutdownOneDriveServer() {
  for (const timer of watchTimers.values()) clearTimeout(timer);
  await closeAllExcelSessions().catch(() => null);
}

function startStdioServer() {
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
        void processMcpMessage(JSON.parse(line)).then((response) => {
          if (response) process.stdout.write(`${JSON.stringify(response)}\n`);
        });
      } catch (error) {
        process.stdout.write(`${JSON.stringify(errorMessage(null, -32700, `Parse error: ${error.message}`))}\n`);
      }
    }
  });
  process.stdin.on("end", () => { shutdownOneDriveServer().catch(() => null); });
  for (const signal of ["SIGINT", "SIGTERM"]) {
    process.on(signal, async () => {
      await shutdownOneDriveServer();
      process.exit(0);
    });
  }
}

const isMainModule = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMainModule) {
  startStdioServer();
}
