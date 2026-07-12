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

- Start with `onedrive_config({ checkToken: true })` when authentication state is relevant. A healthy Keychain refresh token must be reused silently.
- Use `onedrive_doctor` when setup, auth, or Graph access seems questionable, or before a larger workflow where a single health check would save time.
- Use `onedrive_auth_device_start` only after the token check reports no stored credential or a genuine Microsoft reauthentication error. The tool verifies existing authentication itself and returns `alreadyAuthenticated: true` without issuing a code when Keychain access is healthy. Never turn a timeout, network failure, throttling response, or other temporary verification problem into a device-login prompt. Pass `forceReauth: true` only when the user explicitly asks to switch accounts, repair consent, or sign in again.
- Use `onedrive_auth_device_poll` after the user finishes browser sign-in.
- Use `onedrive_logout` only when the user asks to disconnect or reset OneDrive auth. Do not delete the Keychain token unless the user explicitly asks.
- Use `onedrive_me` to confirm the signed-in account.
- Use `onedrive_drive` to inspect drive metadata such as drive type, name, and quota state.
- Use `onedrive_presets` when the user refers to a friendly location such as documents, desktop, screenshots, or a configured alias.
- Use `onedrive_sync_status` when you need to understand cache age, cache size, delta cursor availability, unresolved path count, or local plugin storage paths.
- Use `onedrive_cache_refresh` before broad repeated lookup or cleanup workflows where a warm local metadata cache would save many Graph requests. It batches cache writes, can resume incomplete delta `nextLink` refreshes, reconciles pathless delta records through cached parent IDs where possible, and returns progress milestones. Only delta-origin cursors for the matching root are persisted; ordinary list/search pagination cannot seed delta state, and legacy unscoped cursors are cleared. Use `onedrive_cache_clear` if stale cache data is suspected.
- Use `onedrive_content_index_refresh` only when content search is explicitly useful. It is the expensive file-body extraction step and should run with bounded `maxFiles`, `maxBytesPerFile`, and `concurrencyLimit`. Explicit deletes and changed fingerprints evict stale entries, while moves and renames update indexed metadata; bounded partial scans do not globally prune unseen entries.
- Use `onedrive_content_search` for phrase/content lookup after the content index has been built. It searches local indexed text and returns snippets without calling Microsoft Graph.
- Use `onedrive_content_index_clear` if indexed content is stale or the user asks to remove the local content index.
- Use `onedrive_office_index_refresh` for incremental structured research across Word, Excel, and PowerPoint. It prefers stored delta metadata, reuses unchanged eTag/cTag entries, and stores private semantic anchors. Use `onedrive_office_search` for local paragraph/cell/formula/table/shape/notes matches without Graph calls.
- Use `onedrive_list` before changing or deleting remote files unless the target is already explicit.
- Use compact output by default for list, search, find, scan, and delta results. Pass `format: "full"` only when richer metadata is needed.
- Use `onedrive_list_all` only when the user needs a complete folder listing; set a bounded `maxItems`.
- Use `onedrive_find` for normal file lookup by name, partial name, fuzzy user wording, or content/metadata matches returned by Graph. It runs the canonical Graph query first, retains those canonical results even when the match is inside file content, and expands additional terms in bounded concurrent waves only while confidence remains low. Unrelated expansion-only hits remain filtered. It also uses the local metadata cache and content index when available and bounded remote scan fallback when confidence stays low. It never fetches full file contents itself. Cache-only hits must still be query-relevant and should not be treated as authoritative when live evidence cannot confirm them. Pass `useCache: false` when a fully live lookup is required with no metadata-cache reads or writes, `useContentIndex: false` when indexed content should not influence results, tune `searchConcurrency` for Graph term expansion, and tune `scanConcurrency` only for bounded fallback scans. `graphSearchCalls` counts fetched search pages.
- Use `onedrive_find_all` for broader “look everywhere” or “scan my OneDrive” file-location requests. It executes every planned Graph search term instead of stopping after a confident canonical hit. Keep caps bounded and prefer folder hints when the user gives them; duplicate/nested hints are pruned before fallback scanning.
- Use `onedrive_search` for direct Graph search when the user specifically wants raw Graph search behavior.
- Use `onedrive_search_all` only when the user needs paginated search results; set a bounded `maxItems`.
- Use `onedrive_scan` when the user wants to scan the whole OneDrive or recursively inspect subfolders. Set bounded `maxItems`, `maxFolders`, and `maxResults`; use filters like `nameContains`, `extensions`, and `includeFolders: false` when searching for files.
- Use `onedrive_delta` to answer what changed since a previous scan. Save `deltaLink` when present; use `nextLink` to continue an incomplete delta scan.
- Use `onedrive_get_info` with `includeDeletedItems: true` and an `itemId` when inspecting a deleted item ID on OneDrive Personal.
- Use `onedrive_read_text` only for bounded text files.
- Use `onedrive_preview` for safe bounded previews of text files or Graph-supported document text exports.
- If `onedrive_read_text` refuses likely binary content, use `onedrive_download` instead of forcing text unless the user specifically wants raw text extraction.
- Use `onedrive_download` for binary files or larger files. Default download/export/update checkout paths are collision-safe for concurrent calls; do not assume two concurrent default writes return the same local path.
- Use `onedrive_download_excel`, `onedrive_download_word`, or `onedrive_download_powerpoint` for Office/document-specific downloads when the user names a document type.
- Use `onedrive_office_capabilities` before a format-sensitive Office workflow to inspect the runtime, drive type, and available Excel backend. Use `onedrive_office_validate` for package integrity and macro/signature detection without editing.
- Use `onedrive_word_get_document`, `onedrive_excel_get_workbook`, or `onedrive_powerpoint_get_presentation` to ground paragraphs, tables, cells, formulas, slide indexes, shape IDs, notes, and geometry before native edits. Pass `searchText` for bounded structured matches in any Office file; for faster workbook research, also narrow Excel reads with `sheetNames` or `address`. Excel table metadata is included by default.
- Use the matching `onedrive_*_batch_update` tool for native content edits. Always preview first. Live calls require `dryRun: false`, `confirmed: true`, `expectedName` or `expectedId`, and the returned `previewToken`; backups and post-commit validation are enabled by default.
- Use `onedrive_office_batch_transform` for one preview-token-gated plan across up to 25 Office files. It preflights every file before mutation and returns completed, failed, remaining, and recovery-backup details if a later write fails.
- Office edit previews include a semantic diff with operation counts and affected paragraphs, cells, sheets, slides, shapes, or package parts. Use `onedrive_office_backups` to list managed backups and `onedrive_office_compare_backup` to compare one with its original remote item. Restore only with `onedrive_office_restore_backup` after reviewing its dry-run; live restore requires the returned `previewToken`, the original stable `expectedId`, and the current `expectedETag`.
- Word native operations cover text replacement, paragraphs, table insertion/cells, content controls, hyperlinks, and comments; tracked-change markup is refused. Excel reads cover tables, charts, pivots, and optional formula dependencies; Open XML edits cover cells, ranges, conditional formatting, data validation, frozen panes, column widths, styles/number formats, sheet names, defined names, table-row insertion, totals-row configuration, basic clustered bar/column, line, and pie chart creation/update, and recalculation-on-open. `addTableRow` preserves an existing totals row; use `setTableTotals` to enable/disable totals and configure column labels, functions, or custom formulas. Open XML chart sources must be rectangular same-sheet ranges; changing type requires `sourceData`. Supported business `.xlsx` Graph sessions also provide typed table-row and chart create/update operations. PowerPoint covers text, styling, text boxes, geometry, shape deletion, raster-image replacement, table cells, notes, and slide duplicate/delete/reorder.
- Excel `backend: auto` uses a scoped Graph workbook session for supported business `.xlsx` targets and the local Open XML engine for OneDrive Personal, macro-enabled files, or operations Graph cannot represent. Never force Graph for a consumer workbook.
- The optional `office-addin/` companion runs only inside an active Office host, negotiates protocol `codex-office-companion/1` plus host requirement sets, reports only executable typed commands, rejects unknown properties, and has no remote transport. Use it only when the user explicitly wants cursor/selection-aware in-app editing; headless OneDrive editing remains the default.
- Use `onedrive_export_pdf` or `onedrive_export_text` when the user wants an Office document converted before local processing. These rely on Microsoft Graph conversion support and may fail cleanly for unsupported file types.
- Use `onedrive_upload` for local-file-to-OneDrive uploads only when the source is truly local and not already in a local OneDrive sync folder. It automatically uses upload sessions for files above the simple upload limit; use `uploadMode: "session"` when explicitly testing resumable upload behavior.
- Use `onedrive_write_text` for creating or replacing small text files.
- Use `onedrive_create_folder` to create remote folders; folder names must be single item names, not paths.
- Use `onedrive_rename` to rename one remote item. It defaults to dry-run; only set `dryRun: false` with `confirmed: true` and `expectedName` or `expectedId` after explicit user authorization for the exact item.
- Use `onedrive_move` and `onedrive_copy` when the user asks to reorganize files. They default to dry-run; only set `dryRun: false` with `confirmed: true` and `expectedName` or `expectedId` after explicit user authorization for the exact source item and destination.
- Use `onedrive_update_file` with `mode: "checkout"` before local editing, then `mode: "commit"` with the edited `localPath`. Commit checks remote eTag/cTag/size/mtime against checkout metadata, creates a local backup by default, and verifies after upload. Checkout refuses to overwrite an existing manifest unless `overwriteManifest: true` is provided. Use `force: true` only when intentionally overwriting a changed remote file or committing without checkout metadata.
- Use `onedrive_batch_get_info` and `onedrive_batch_permissions` for up to 20-item metadata/permission audits. Use `onedrive_batch_download`, `onedrive_batch_delete`, and `onedrive_batch_move` for multi-item workflows; keep destructive batch deletes dry-run until explicitly confirmed. Live `onedrive_batch_delete` requires `dryRun: false`, `confirmed: true`, `expectedName` or `expectedId` for every item, and the `previewToken` from the dry-run. Live `onedrive_batch_move` requires `dryRun: false`, `confirmed: true`, and `expectedName` or `expectedId` for every item. Live batch mutation responses warn that earlier items may already be changed if a later item fails.
- Graph batch helpers retry only transient individual subrequests and preserve the original result order. Do not retry deterministic per-item failures as though the entire batch failed.
- Use `onedrive_recent` for recent files, `onedrive_large_files` and `onedrive_duplicates` for cleanup, and `onedrive_shared_by_me` or `onedrive_public_links` for sharing audits.
- Use `onedrive_batch_move` for multi-item moves. Preview first; live calls require `dryRun: false`, `confirmed: true`, and `expectedName` or `expectedId` on every item. If any item fails preflight, do not retry live until every target is fixed. Treat partial-state warnings seriously after live failures.
- Use `onedrive_permissions` to audit sharing/permission state before creating links, inviting recipients, or changing access.
- `onedrive_create_sharing_link` changes link access to a file or folder. It supports optional `password` and `expirationDateTime`, defaults to dry-run, and includes permission-preview data by default; only set `dryRun: false` and `confirmed: true` with `expectedName` or `expectedId` and the returned `previewToken` after explicit user authorization for the exact item and link scope. Never repeat link passwords back to the user after a dry-run; refer to `passwordProvided`.
- Use `onedrive_invite_permission` for named-user or group grants. It defaults to a silent direct grant with `sendInvitation: false`, `requireSignIn: true`, and `role: "read"`. Use `sendInvitation: true` and optional `message` only when the user explicitly asks to email the recipient. It defaults to dry-run and includes permission-preview data by default; live calls require `dryRun: false`, `confirmed: true`, `expectedName` or `expectedId`, and the returned `previewToken` after explicit user authorization for the exact item, recipient count/type, role, and email-invite setting.
- Use `onedrive_revoke_permission` to unshare one permission after auditing permissions. It defaults to dry-run and includes before permissions by default; live calls require `dryRun: false`, `confirmed: true`, `permissionId`, `expectedName` or `expectedId`, and the returned `previewToken`.
- Use `onedrive_batch_revoke_permissions` for multi-item unsharing. Preview first; live calls require `dryRun: false`, `confirmed: true`, `permissionId`, `expectedName` or `expectedId` for every item, and the returned `previewToken`. Batch revoke preflights all entries before any permission DELETE.
- `onedrive_restore_deleted` restores by deleted item ID and defaults to dry-run; only set `dryRun: false` and `confirmed: true` with `expectedId` and the returned `previewToken` after explicit user authorization. Microsoft Graph does not provide a normal OneDrive recycle-bin listing API through driveItem file operations, so use `onedrive_delta` deleted changes to discover deleted item IDs when available.
- `onedrive_delete` defaults to `dryRun: true`; only set `dryRun: false` and `confirmed: true` after explicit user authorization. Live deletes require `expectedName` or `expectedId` and the returned `previewToken`.
- Use `onedrive_audit_recent` to review recent live mutation entries, `onedrive_audit_export` to save the JSONL audit log, and `onedrive_audit_clear` only when the user explicitly asks to clear it. Clearing requires `confirmed: true`.
- If a live mutation response contains `localWarnings` or `verificationIncomplete`, treat the remote mutation as successful and the local audit/cache or post-verification step as incomplete. Do not repeat the remote mutation just to repair follow-up state.
- Do not ask users for a vague "yes" before live mutation. First show the dry-run preview, then ask for confirmation of the exact item name/ID, destination or permission ID, and intended change.
- For high-risk live sharing, revocation, restore, delete, and batch delete/revoke calls, pass the `previewToken` from the immediately preceding dry-run preview for the same resolved operation.
- Uploads and downloads refuse local OneDrive sync-folder paths by default. Pass `allowLocalOneDriveSyncPath: true` only when the user explicitly needs the local synced folder rather than the remote plugin path.
- Plugin-managed directories are private to the current user (`0700`) and persisted metadata, extracted content, audit records, downloads, exports, and update manifests are private files (`0600`). Metadata cache, content index, and audit files use atomic/locked persistence so concurrent plugin processes do not lose updates. Remember that the optional content index contains extracted file bodies even though it never sends them outside Microsoft Graph/local plugin storage.

## Maintenance

- Run `scripts/mock-graph-test.mjs` for fast regression coverage without touching OneDrive.
- Run `scripts/benchmark.mjs --query="<query>"` when comparing cold search, warm cache search, content-indexed search, and selected preview timing. It emits progress events to stderr during long steps and exits nonzero if any MCP tool step fails.
- Run `scripts/prepackage-check.mjs` before reinstalling the plugin cache.
- Run `scripts/prepackage-check.mjs --installed <installed-cache-path>` after reinstalling to catch source/cache drift.
- Run `scripts/beta-test.mjs` only when live OneDrive CRUD verification is intended.
- The beta harness rejects unknown, duplicate, and positional CLI options before startup; correct a rejected flag instead of retrying with a guessed alternative.
- Run `scripts/beta-test.mjs --cleanup-stale --stale-days=1` to dry-run cleanup of old `Codex OneDrive Plugin Beta Test ...` folders; add `--confirmed` only after reviewing candidates. Cleanup discovers candidates with bounded search and verifies each item; missing/invalid timestamps fail closed and are skipped.
- Run `scripts/beta-test.mjs --tenant-matrix=common,consumers,organizations` for read-only tenant health checks, and add `--tenant-matrix-live` only when full live beta runs are intended for each tenant.
