---
name: onedrive
description: Work with files and folders in OneDrive through the local OneDrive Codex plugin. Use when the user asks to list, search, read, create, upload, download, rename, move, copy, share, or delete OneDrive files.
---

# OneDrive

Use the `onedrive` MCP tools to manage files in the signed-in user's OneDrive through Microsoft Graph.
Prefer remote Microsoft Graph operations through this plugin. Do not inspect, write, or rely on the user's local OneDrive sync folder on the laptop unless the user explicitly asks for local sync-folder access or a local temporary file is required for processing.

## Configuration

The MCP server reads non-secret config from environment variables and `~/.codex/onedrive-plugin/config.json`.
Refresh tokens are stored in macOS Keychain.

- `ONEDRIVE_CLIENT_ID`: Microsoft app registration client ID for a public client app.
- `ONEDRIVE_TENANT`: Microsoft identity tenant. Defaults to `common`.
- `ONEDRIVE_SCOPES`: Space-separated delegated scopes. Defaults to `offline_access User.Read Files.ReadWrite`.
- `ONEDRIVE_KEYCHAIN_SERVICE`: Optional Keychain service name. Defaults to `Codex OneDrive`.

Never ask the user to paste Microsoft passwords, access tokens, or refresh tokens into chat. Use device-code login and Keychain-backed storage.
When the configured tenant is `common`, the plugin can retry Microsoft-account-only auth failures on the `consumers` endpoint automatically.

## Tool Guidance

- Start with `onedrive_config` to check whether a client ID and Keychain token are configured.
- Use `onedrive_doctor` when setup, auth, or Graph access seems questionable, or before a larger workflow where a single health check would save time.
- Use `onedrive_auth_device_start` when no refresh token is stored. Ask the user to open the returned verification URL and enter the returned user code.
- Use `onedrive_auth_device_poll` after the user finishes browser sign-in.
- Use `onedrive_logout` only when the user asks to disconnect or reset OneDrive auth. Do not delete the Keychain token unless the user explicitly asks.
- Use `onedrive_me` to confirm the signed-in account.
- Use `onedrive_drive` to inspect drive metadata such as drive type, name, and quota state.
- Use `onedrive_presets` when the user refers to a friendly location such as documents, desktop, screenshots, or a configured alias.
- Use `onedrive_sync_status` when you need to understand cache age, cache size, delta cursor availability, or local plugin storage paths.
- Use `onedrive_cache_refresh` before broad repeated lookup or cleanup workflows where a warm local metadata cache would save many Graph requests. Use `onedrive_cache_clear` if stale cache data is suspected.
- Use `onedrive_list` before changing or deleting remote files unless the target is already explicit.
- Use compact output by default for list, search, find, scan, and delta results. Pass `format: "full"` only when richer metadata is needed.
- Use `onedrive_list_all` only when the user needs a complete folder listing; set a bounded `maxItems`.
- Use `onedrive_find` for normal file lookup by name, partial name, or fuzzy user wording. It uses the local metadata cache when available, performs live Graph searches, ranks matches in memory, and uses bounded remote scan fallback when confidence is low. Cache-only hits must still be query-relevant and should not be treated as authoritative when live evidence cannot confirm them. Pass `useCache: false` when a fully live lookup is required.
- Use `onedrive_find_all` for broader “look everywhere” or “scan my OneDrive” file-location requests. Keep caps bounded and prefer folder hints when the user gives them.
- Use `onedrive_search` for direct Graph search when the user specifically wants raw Graph search behavior.
- Use `onedrive_search_all` only when the user needs paginated search results; set a bounded `maxItems`.
- Use `onedrive_scan` when the user wants to scan the whole OneDrive or recursively inspect subfolders. Set bounded `maxItems`, `maxFolders`, and `maxResults`; use filters like `nameContains`, `extensions`, and `includeFolders: false` when searching for files.
- Use `onedrive_delta` to answer what changed since a previous scan. Save `deltaLink` when present; use `nextLink` to continue an incomplete delta scan.
- Use `onedrive_get_info` with `includeDeletedItems: true` and an `itemId` when inspecting a deleted item ID on OneDrive Personal.
- Use `onedrive_read_text` only for bounded text files.
- Use `onedrive_preview` for safe bounded previews of text files or Graph-supported document text exports.
- If `onedrive_read_text` refuses likely binary content, use `onedrive_download` instead of forcing text unless the user specifically wants raw text extraction.
- Use `onedrive_download` for binary files or larger files.
- Use `onedrive_download_excel`, `onedrive_download_word`, or `onedrive_download_powerpoint` for Office/document-specific downloads when the user names a document type.
- Use `onedrive_export_pdf` or `onedrive_export_text` when the user wants an Office document converted before local processing. These rely on Microsoft Graph conversion support and may fail cleanly for unsupported file types.
- Use `onedrive_upload` for local-file-to-OneDrive uploads only when the source is truly local and not already in a local OneDrive sync folder. It automatically uses upload sessions for files above the simple upload limit; use `uploadMode: "session"` when explicitly testing resumable upload behavior.
- Use `onedrive_write_text` for creating or replacing small text files.
- Use `onedrive_create_folder` to create remote folders; folder names must be single item names, not paths.
- Use `onedrive_rename` to rename one remote item; include `expectedName` or `expectedId` when the item was resolved earlier. Use `dryRun: true` for preview.
- Use `onedrive_move` and `onedrive_copy` when the user asks to reorganize files; include `expectedName` or `expectedId` when the target was resolved in a previous step. Use `dryRun: true` for preview.
- Use `onedrive_update_file` with `mode: "checkout"` before local editing, then `mode: "commit"` with the edited `localPath`. Commit checks remote eTag/cTag/size/mtime against checkout metadata, creates a local backup by default, and verifies after upload. Checkout refuses to overwrite an existing manifest unless `overwriteManifest: true` is provided. Use `force: true` only when intentionally overwriting a changed remote file or committing without checkout metadata.
- Use `onedrive_batch_get_info` and `onedrive_batch_permissions` for up to 20-item metadata/permission audits. Use `onedrive_batch_download`, `onedrive_batch_delete`, and `onedrive_batch_move` for multi-item workflows; keep destructive batch deletes dry-run until explicitly confirmed. Live `onedrive_batch_move` requires `dryRun: false`, `confirmed: true`, and `expectedName` or `expectedId` for every item.
- Use `onedrive_recent` for recent files, `onedrive_large_files` and `onedrive_duplicates` for cleanup, and `onedrive_shared_by_me` or `onedrive_public_links` for sharing audits.
- Use `onedrive_permissions` to audit sharing/permission state before creating or changing links.
- `onedrive_create_sharing_link` changes access to a file or folder. It defaults to dry-run and includes permission-preview data by default; only set `dryRun: false` and `confirmed: true` after explicit user authorization for the exact item and link scope.
- `onedrive_restore_deleted` restores by deleted item ID and defaults to dry-run; only set `dryRun: false` and `confirmed: true` after explicit user authorization. Microsoft Graph does not provide a normal OneDrive recycle-bin listing API through driveItem file operations, so use `onedrive_delta` deleted changes to discover deleted item IDs when available.
- `onedrive_delete` defaults to `dryRun: true`; only set `dryRun: false` and `confirmed: true` after explicit user authorization. Live deletes require `expectedName` or `expectedId`.
- Uploads and downloads refuse local OneDrive sync-folder paths by default. Pass `allowLocalOneDriveSyncPath: true` only when the user explicitly needs the local synced folder rather than the remote plugin path.

## Maintenance

- Run `scripts/mock-graph-test.mjs` for fast regression coverage without touching OneDrive.
- Run `scripts/prepackage-check.mjs` before reinstalling the plugin cache.
- Run `scripts/prepackage-check.mjs --installed <installed-cache-path>` after reinstalling to catch source/cache drift.
- Run `scripts/beta-test.mjs` only when live OneDrive CRUD verification is intended.
