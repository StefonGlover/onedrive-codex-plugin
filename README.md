# OneDrive Codex Plugin

Local Codex plugin for OneDrive file operations through Microsoft Graph.

Release `0.5.0+codex.20260714034051` exposes an exact 84-tool contract.

This is an unofficial integration and is not affiliated with, endorsed by, or sponsored by Microsoft.

The plugin is remote-first: it uses Microsoft Graph rather than the laptop's local OneDrive sync folder. Upload and download tools refuse local OneDrive sync-folder paths by default unless `allowLocalOneDriveSyncPath: true` is explicitly provided.

## Setup

This plugin uses Microsoft identity platform device-code login and stores refresh tokens in macOS Keychain. It does not store Microsoft passwords.

1. Create or choose a Microsoft Entra app registration.
2. Enable public client flows for the app.
3. Add delegated Microsoft Graph permissions for `User.Read`, `Files.ReadWrite`, and `offline_access`.
   - Optional: add `Files.ReadWrite.All` if you want to attempt live restore of deleted personal OneDrive items with `onedrive_restore_deleted`.
4. Run the local setup script:

```bash
scripts/configure.zsh
```

5. Start a fresh Codex thread after installing or refreshing the plugin.
6. Ask Codex to call `onedrive_config` with `checkToken: true`. If no reusable credential exists, call `onedrive_auth_device_start`, open the returned verification URL, enter the returned user code, then call `onedrive_auth_device_poll`.

After the first successful login, the refresh token is reused from macOS Keychain. `onedrive_auth_device_start` now checks that credential before contacting Microsoft's device-code endpoint and returns `alreadyAuthenticated: true` without generating a code when authentication is healthy. Use `forceReauth: true` only for an intentional account switch, consent repair, or explicit sign-in reset. Temporary token-check network failures do not trigger a new login flow.

## Configuration

The setup script writes non-secret config to:

```text
~/.codex/onedrive-plugin/config.json
```

Supported environment variables:

```bash
export ONEDRIVE_CLIENT_ID="your-public-client-app-id"
export ONEDRIVE_TENANT="common"
export ONEDRIVE_SCOPES="offline_access User.Read Files.ReadWrite"
export ONEDRIVE_KEYCHAIN_SERVICE="Codex OneDrive"
export ONEDRIVE_STORAGE_ROOT="$HOME/.codex/onedrive-plugin"
export ONEDRIVE_CACHE_ROOT="$HOME/.codex/onedrive-plugin/cache"
export ONEDRIVE_CACHE_TTL_SECONDS="900"
export ONEDRIVE_MAX_SCAN_DEPTH="25"
export ONEDRIVE_MAX_INDEXED_FILE_SIZE="524288"
export ONEDRIVE_INDEX_EXTENSIONS=".txt,.md,.csv,.json,.jsonl,.xml,.yaml,.yml,.html,.css,.js,.mjs,.ts,.tsx,.py,.sql,.log"
export ONEDRIVE_CONCURRENCY_LIMIT="2"
export ONEDRIVE_DELTA_SYNC_ENABLED="true"
export ONEDRIVE_CONTENT_INDEX_ENABLED="true"
export ONEDRIVE_INDEX_OFFICE_EXPORT="false"
```

`ONEDRIVE_TENANT` can be `common`, `consumers`, `organizations`, or a tenant ID. Use `common` for a plugin that may access either personal Microsoft accounts or work/school accounts.
If Microsoft reports that the app is Microsoft-account-only and requires `/consumers`, the plugin retries device-code and refresh-token auth on `consumers` automatically when the configured tenant is `common`.

You can also add friendly path aliases to the config file:

```json
{
  "pathPresets": {
    "job-tracker": "Documents/Job Tracker",
    "screenshots": "Pictures/Screenshots"
  }
}
```

Tools that accept `path` can also accept `preset` plus `relativePath`; upload/write tools accept `remotePreset` plus `remoteRelativePath`.

Optional performance settings can also be placed in the config file:

```json
{
  "storageRoot": "/absolute/path/to/onedrive-plugin",
  "cacheRoot": "/absolute/path/to/onedrive-plugin/cache",
  "settings": {
    "cacheTtlSeconds": 900,
    "maxScanDepth": 25,
    "concurrencyLimit": 2,
    "deltaSyncEnabled": true,
    "contentIndexEnabled": true
  },
  "indexing": {
    "maxFileSize": 524288,
    "supportedExtensions": [".txt", ".md", ".csv", ".json", ".jsonl", ".xml", ".yaml", ".yml"],
    "includeOfficeTextExport": false
  }
}
```

Use absolute paths in `storageRoot` and `cacheRoot` if you override them. Configuration precedence is environment variables, then `~/.codex/onedrive-plugin/config.json`, then server defaults. The checked-in `.mcp.json` does not override tenant or scopes.

## Tools

- `onedrive_config`
- `onedrive_doctor`
- `onedrive_auth_device_start`
- `onedrive_auth_device_poll`
- `onedrive_logout`
- `onedrive_me`
- `onedrive_drive`
- `onedrive_presets`
- `onedrive_list`
- `onedrive_list_all`
- `onedrive_scan`
- `onedrive_search`
- `onedrive_search_all`
- `onedrive_find`
- `onedrive_find_all`
- `onedrive_delta`
- `onedrive_sync_status`
- `onedrive_cache_refresh`
- `onedrive_cache_clear`
- `onedrive_content_index_refresh`
- `onedrive_content_search`
- `onedrive_content_index_clear`
- `onedrive_office_capabilities`
- `onedrive_office_validate`
- `onedrive_office_index_refresh`
- `onedrive_office_search`
- `onedrive_word_get_document`
- `onedrive_excel_get_workbook`
- `onedrive_powerpoint_get_presentation`
- `onedrive_word_batch_update`
- `onedrive_excel_batch_update`
- `onedrive_powerpoint_batch_update`
- `onedrive_office_batch_transform`
- `onedrive_office_backups`
- `onedrive_office_compare_backup`
- `onedrive_office_restore_backup`
- `onedrive_get_info`
- `onedrive_read_text`
- `onedrive_preview`
- `onedrive_download`
- `onedrive_download_excel`
- `onedrive_download_word`
- `onedrive_download_powerpoint`
- `onedrive_export_pdf`
- `onedrive_export_text`
- `onedrive_upload`
- `onedrive_write_text`
- `onedrive_create_folder`
- `onedrive_rename`
- `onedrive_move`
- `onedrive_copy`
- `onedrive_create_sharing_link`
- `onedrive_invite_permission`
- `onedrive_revoke_permission`
- `onedrive_batch_revoke_permissions`
- `onedrive_permissions`
- `onedrive_batch_get_info`
- `onedrive_batch_permissions`
- `onedrive_batch_download`
- `onedrive_batch_delete`
- `onedrive_batch_move`
- `onedrive_update_file`
- `onedrive_recent`
- `onedrive_large_files`
- `onedrive_duplicates`
- `onedrive_shared_by_me`
- `onedrive_public_links`
- `onedrive_restore_deleted`
- `onedrive_audit_recent`
- `onedrive_audit_export`
- `onedrive_audit_clear`
- `onedrive_delete`

## Native Office editing

The plugin can inspect and edit modern Open XML files without requiring Word, Excel, or PowerPoint to be open. Reads expose structured document content and package-safety metadata. Mutation previews include semantic operation counts and affected objects, are bound to file identity and eTag/cTag, backed up locally by default, uploaded with `If-Match`, validated after commit, and recorded in the mutation audit.

Managed Office backups have opaque IDs and manifests containing the original stable item ID and version metadata. Use `onedrive_office_backups` to list them and `onedrive_office_compare_backup` for a bounded semantic comparison with the current remote content. `onedrive_office_restore_backup` defaults to dry-run and requires the preview token, original `expectedId`, and current `expectedETag`; it restores by item ID, creates a rollback backup, audits the mutation, and verifies the restored fingerprint.

- Word exposes durable paragraph, table, and content-control anchors and 21 headless operations, including image insertion/replacement, bookmarks, content controls, table row/column changes, headers/footers, and section properties. Documents containing tracked changes are refused instead of silently changing review semantics.
- Excel exposes worksheet, range, defined-name, and table/row-key anchors and 32 headless operations, including worksheet/table lifecycle, merge/unmerge, sort/filter, hyperlinks, notes, images, chart formatting, passwordless sheet protection, and pivot refresh-on-open. Business `.xlsx` Graph writes use scoped persistent sessions; personal workbooks and unsupported Graph operations use Open XML automatically.
- PowerPoint exposes persistent slide and shape anchors and 25 headless operations, including slides, images/cropping, tables and row/column changes, alternative text, z-order, grouping/ungrouping, and layout application.
- Positional selectors remain valid. An anchor defaults to `rebasePolicy:"unique"`; moved targets re-resolve only when exactly one match exists. Missing, duplicate, or selector/anchor disagreement returns a structured conflict. Every live commit remains bound to the eTag used by its preview.
- Encrypted and legacy binary files are refused. Macro-enabled edits require `allowMacros: true`; signed-package edits are always refused because any edit invalidates the signature.

Run the corresponding structured read tool first, then call the batch-update tool as a dry run. A live commit requires `dryRun: false`, `confirmed: true`, `expectedName` or `expectedId`, and the exact returned `previewToken`.

For faster research inside a known Office file, pass `searchText` to any structured read. Excel can additionally restrict reads to `sheetNames` and a bounded A1 `address`, avoiding a full-workbook response.

For cross-drive research, `onedrive_office_index_refresh` stores structured paragraph, cell, formula, table, content-control, comment, shape, and notes segments with semantic anchors. It reuses unchanged eTag/cTag entries and prefers the existing OneDrive delta cursor before scanning. `onedrive_office_search` searches that private local index without Graph calls. `onedrive_office_batch_transform` preflights every requested file before the first write and returns recovery backup IDs if a later item fails.

## Remote editing workflows

- `onedrive_versions`, `onedrive_compare_version`, and `onedrive_restore_version` expose bounded Graph version history, semantic/text/binary comparisons, and native restore. Restore never falls back to replacing content and requires current identity, eTag, confirmation, and its preview token.
- `onedrive_patch_text` applies bounded unified diffs, RFC 6902 JSON Patch, restricted safe-YAML path operations, or RFC 4180 CSV row-key operations. It preserves supported BOM/encoding, newline style, and trailing-newline state while refusing binary and oversized inputs.
- `onedrive_workspace_*` manages owner-only drafts under `Codex Editing Drafts`. Workspaces record the original stable item ID and base eTag/version, surface source/draft drift, block promotion after source drift, preserve the original item identity/version history on success, and retain failed/conflicted drafts for recovery.
- `onedrive_watch_*` manages auth-context/drive-scoped delta watches with 15–300 second polling, one-hour default expiry, eight-hour maximum, throttling backoff, and a 500-event ring buffer. Events invalidate affected previews and mark source workspaces stale.

## Safety

- Remote mutations that move, rename, copy, expose, invite, restore, delete, or revoke access use a preview-first pattern.
- Rename, move, copy, sharing-link creation, named-recipient invitation, permission revoke, restore, and delete default to dry-run where the operation has a dry-run mode.
- Live rename, move, copy, sharing-link creation, named-recipient invitation, permission revoke, restore, and delete require `dryRun: false`, `confirmed: true`, and stable expected identity (`expectedName` or `expectedId`; restore requires `expectedId`).
- Replacing an existing file with `onedrive_upload` or `onedrive_write_text` is also preview-token gated: review the returned existing item, then repeat with `dryRun: false`, `confirmed: true`, matching expected identity, and the exact preview token.
- Live sharing-link creation, named-recipient invitation, permission revoke, batch permission revoke, restore, delete, and batch delete also require the `previewToken` returned by the immediately preceding dry-run preview for the same resolved operation.
- Batch delete, batch move, and batch permission revoke preflight every item before any mutation and refuse partial execution when a preflight check fails. Live batch responses include a warning that successful earlier items may already be changed if a later item fails.
- Sharing-link creation supports Microsoft Graph link type/scope plus optional password and expiration, and can include a before/after permission diff so the caller can see what changed.
- `onedrive_invite_permission` grants named users or groups access through Microsoft Graph `driveItem: invite`. It defaults to a silent direct grant (`sendInvitation: false`, `requireSignIn: true`); email invitations are opt-in with `sendInvitation: true` and optional `message`.
- Permission revoke uses Microsoft Graph `DELETE /me/drive/items/{item-id}/permissions/{permission-id}` and includes before permissions by default; live revoke includes after permissions and a permission diff.
- Rename, move, copy, share, and delete refuse to operate on the OneDrive root.
- A relative target or destination field is refused unless its matching preset field is present; it never silently resolves relative to the OneDrive root.
- Tool arguments are validated before handlers run, including required fields, unknown properties, enum values, numeric bounds, array bounds, and target `anyOf` rules.
- Text reads are bounded to 5 MB by default.
- Text reads use MIME/extension checks and refuse likely binary files unless `force: true` is set.
- Downloads go to `~/.codex/onedrive-plugin/downloads` unless `localPath` is provided. Concurrent default downloads, exports, update checkouts, backups, and audit exports reserve unique local paths instead of racing on the same filename.
- Downloads and uploads refuse local OneDrive sync-folder paths by default. Use `allowLocalOneDriveSyncPath: true` only for an explicit local sync-folder workflow.
- Uploads use simple upload for smaller files and upload sessions for large files, or when `uploadMode: "session"` is requested.
- List, search, find, scan, and delta tools return compact item summaries by default; pass `format: "full"` for richer metadata.
- Normal list, search, scan, delta, and metadata calls opportunistically maintain a local metadata cache at `~/.codex/onedrive-plugin/cache/metadata-cache.json`.
- Plugin-managed storage, cache, audit, download, export, and update-workflow directories are restricted to the current user (`0700`), and locally persisted metadata, extracted content, audit records, downloads, exports, and update manifests are restricted to the current user (`0600`). Existing managed cache/index/audit files are hardened when loaded or rewritten.
- `onedrive_sync_status` reports cache age, item count, delta cursor availability, resumable delta next-link availability, unresolved path count, and plugin storage locations.
- Metadata cache v4 and content/Office indexes v3 are bound to both an opaque authentication context and the active drive ID. Delta cursors, preview tokens, Office backups, and update manifests use the same scope; legacy unscoped local state is invalidated or retained but refused until safely recreated. Local-only search fails closed on a scope mismatch. Cache, index, and audit files use atomic/locked updates with reload-on-write freshness checks.
- `onedrive_cache_refresh` rebuilds the cache from a bounded recursive scan and uses delta refreshes when a previous cursor exists for the same root. Cache refresh batches metadata-cache writes during scans, persists only delta-origin `nextLink` cursors for continuation, reconciles pathless delta records through cached parent IDs, and returns progress milestones. Ordinary list/search pagination cannot seed delta state. `onedrive_cache_clear` clears the cache.
- `onedrive_content_index_refresh` is the explicit content-reading step. It indexes supported cached text and structured Office content into `content-index.json`, stores normalized text/tokens plus semantic Office anchors, reuses entries when ETag/cTag/mtime/size are unchanged, and applies file-size, segment, concurrency, and per-item failure limits. Explicit metadata deletes and changed fingerprints evict stale entries; moves and renames update indexed metadata; unchanged explicit cTags preserve content entries across metadata-only renames, while changed or omitted content tags with changed ETags invalidate conservatively. Bounded partial scans do not globally prune unseen entries.
- `onedrive_content_search` searches only the local content index and returns lightweight metadata plus snippets. It does not call Microsoft Graph or read file bodies.
- `onedrive_find` and `onedrive_find_all` can merge local content-index hits into ranking, but they never fetch or parse full content themselves. Build or refresh the index first when content search is needed.
- `onedrive_find` is the preferred file lookup helper. It uses the local metadata cache when available, confirms exact strong cache hits with live metadata, runs the canonical Graph query first, and expands additional terms in bounded concurrent waves only while confidence remains low. Canonical Graph results can represent filename, metadata, or file-content matches; unrelated expansion-only results remain gated. Results expose the planned, executed, and skipped search terms. `graphSearchCalls` reports actual Graph search pages fetched, not just term count. Tune expansion with `searchConcurrency` and fallback scans with `scanConcurrency`. Fallback scans prune duplicate and nested folder hints regardless of input order. Cache-only hits must still have query relevance and are not treated as authoritative when live evidence cannot confirm them. Pass `useCache: false` for a fully live lookup with no metadata-cache reads or writes.
- `onedrive_find_all` is the broader locator for “look everywhere” requests. It searches every planned term instead of stopping after the first confident canonical result, searches common folders first, and uses larger bounded scan caps, with cache acceleration and the same duplicate-hint pruning when available.
- `onedrive_preview` returns bounded text previews for text files and Graph-supported document text exports without reading unbounded remote content into memory.
- `onedrive_update_file` provides a checkout/commit edit workflow with a local manifest, eTag/cTag/size/mtime conflict checks, optional backup, and post-commit verification. Checkout refuses to overwrite an existing manifest unless `overwriteManifest: true` is provided.
- `onedrive_batch_get_info` and `onedrive_batch_permissions` use Microsoft Graph batching for up to 20 items. Batch download/delete/move tools provide one result per item with dry-run support where destructive.
- `onedrive_rename`, `onedrive_move`, and `onedrive_copy` support `dryRun: true` previews. Live `onedrive_batch_move` requires `dryRun: false`, `confirmed: true`, and `expectedName` or `expectedId` for every item.
- `onedrive_recent`, `onedrive_large_files`, `onedrive_duplicates`, `onedrive_shared_by_me`, and `onedrive_public_links` provide cleanup and sharing-audit workflows.
- `onedrive_list_all` follows pagination within one folder. Use `onedrive_scan` when you need recursive traversal across subfolders or the whole OneDrive.
- `onedrive_doctor` checks config, auth, profile, drive metadata, presets, and optional root listing in one call.
- `onedrive_export_pdf` and `onedrive_export_text` ask Microsoft Graph to convert supported Office files before saving locally. Microsoft Graph may reject conversions for unsupported file types.
- `onedrive_permissions` audits current sharing/permission grants before changing access.
- `onedrive_delta` can return deleted item changes. Set `maxPages` from 1 to 100 to cap Graph pages in one call while retaining the advanced `nextLink` or terminal `deltaLink`. Microsoft Graph does not expose a normal OneDrive recycle-bin listing endpoint through the driveItem file APIs.
- `onedrive_get_info` supports `includeDeletedItems: true` when targeting an item ID; Microsoft documents this as OneDrive Personal-only.
- `onedrive_restore_deleted` defaults to dry-run and requires a deleted item ID. Live restore may require `Files.ReadWrite.All` for personal OneDrive.
- Live remote mutations are recorded in a local JSONL audit log at `~/.codex/onedrive-plugin/audit/mutations.jsonl`. Audit entries include safe item summaries, before/after summaries when available, permission diffs when relevant, Graph request IDs when available, and safe error details for failed live mutations. They do not log tokens, authorization headers, file contents, raw request bodies, sharing-link web URLs, passwords, invite messages, or recipient identifiers.
- A successful remote mutation remains a success if local cache/audit bookkeeping or a best-effort post-mutation verification later fails. The response reports `localWarnings` and, where applicable, `verificationIncomplete`; do not repeat the mutation merely to repair that follow-up state.
- `onedrive_audit_recent` reads recent audit entries, `onedrive_audit_export` exports the JSONL log to a local file, and `onedrive_audit_clear` requires `confirmed: true`.
- Graph requests retry transient `429`, `500`, `502`, `503`, and `504` responses with `Retry-After` support. Read-only requests also retry transient transport failures, and Microsoft Graph batch helpers retry only the transient individual subrequests while preserving result order.

## Performance Architecture

The plugin separates cheap metadata discovery from expensive content reads:

- Metadata cache: stores IDs, drive/item metadata, paths, web URLs, MIME/type hints, size, timestamps, ETag/cTag, and file/folder status. This avoids repeated full recursive scans when cached metadata is fresh enough for the workflow.
- Delta sync: `onedrive_cache_refresh` prefers stored Microsoft Graph delta cursors when the requested root matches the cached root. If Graph returns a delta `nextLink` before the final `deltaLink`, the plugin stores that incomplete cursor and resumes it on the next refresh. It resolves pathless delta records from cached parent IDs where possible, reports unresolved path counts for parents it cannot hydrate, repaths descendants after folder moves/renames, removes deleted descendants, and rejects legacy non-delta cursors. It falls back to a bounded scan when no cursor exists, the target changed, delta is disabled, or Graph rejects the cursor.
- Search ranking: `onedrive_find` combines exact path/filename evidence, canonical live Graph content/metadata matches, confirmed metadata-cache matches, folder hints, file-type hints, recency-ish modified metadata, and optional content-index hits. Adaptive term execution stops expansion after a confident live result for normal `find`; `find_all` executes all planned terms for exhaustive locator requests. Results include reasons and request-plan counters so ranking and latency are debuggable.
- Content index: indexing is opt-in and explicit. It supports bounded text-like files by default, stores normalized text/tokens for faster repeated queries, keeps only bounded top matches during local content search, and can optionally try Graph `format=text` export for Office-like files. Large files, unsupported binaries, failed exports, and files over the cap are skipped or reported without aborting the whole refresh. Cache reconciliation removes demonstrably stale entries without treating a bounded partial scan as a complete-drive deletion signal.
- Graph optimization: once an item is discovered, the plugin prefers item IDs for follow-up reads/mutations, follows pagination with cycle/page caps, uses `$batch` where useful, retries transient throttling/service errors (including transient batch subresponses), and honors `Retry-After`.

The largest performance risk is any broad recursive scan or content-index refresh over a large OneDrive tree. Keep `maxItems`, `maxFolders`, `maxDepth`, `maxFiles`, and `maxBytesPerFile` bounded, and warm the metadata cache before broad repeated searches.

## Troubleshooting and Benchmarks

Run a health check:

```bash
onedrive_doctor({ "checkRootList": true })
```

Inspect cache/index state:

```bash
onedrive_sync_status({ "includeSamples": true })
```

Refresh metadata using scan/delta:

```bash
onedrive_cache_refresh({ "mode": "auto", "maxItems": 10000, "maxFolders": 2000 })
```

Build the optional content index from cached metadata:

```bash
onedrive_content_index_refresh({ "maxFiles": 100, "maxBytesPerFile": 524288 })
```

For a simple before/after benchmark, compare:

1. Cold search: clear cache, run `onedrive_find` with `useCache: false`.
2. Warm metadata search: run `onedrive_cache_refresh`, then repeat `onedrive_find`.
3. Content-indexed search: run `onedrive_content_index_refresh`, then run `onedrive_content_search` and `onedrive_find` for a phrase inside a file.
4. Selected file read: use `onedrive_preview` or `onedrive_read_text` only after selecting a specific result.

The bundled benchmark script runs those steps through the MCP server with bounded caps:

```bash
scripts/benchmark.mjs --query="project plan" --maxItems=1500 --maxFolders=250 --maxFiles=50 --searchConcurrency=2
```

Add `--clear` when you intentionally want to clear local metadata/content caches before the cold run. The script performs read-only Microsoft Graph operations, writes local cache/index files, and emits progress events to stderr while keeping the final summary JSON on stdout.

Expected improvement: confident canonical matches use one Graph search instead of running every generated term; lower-confidence queries expand terms with bounded concurrency. Warm metadata searches avoid most repeated recursive scans, delta refreshes fetch only changes after the initial scan, and content searches avoid live file reads after the index is built. Benchmark summaries report planned, executed, and skipped terms. Actual speed depends on Microsoft Graph latency, OneDrive size, throttling, and configured caps.

## Safe Example Prompts

- "Find the file named Project Plan, show me its item ID and current permissions, but do not change anything."
- "Dry-run moving this file to Documents/Archive and show the exact item ID I would need to confirm."
- "Dry-run granting person@example.com read access to this item without sending an email."
- "Revoke the anonymous link from this item only after previewing the permission diff and asking me to confirm."
- "Show recent OneDrive mutation audit entries from this plugin."

## Beta Test

Run the mocked Microsoft Graph regression suite first. It does not touch OneDrive, Keychain, or Microsoft services:

```bash
scripts/mock-graph-test.mjs
```

Run the benchmark script when comparing cold search, warm cache search, indexed content search, and selected preview timing. The script exits nonzero if any MCP tool step reports an error:

```bash
scripts/benchmark.mjs --query="project plan"
```

Run the prepackage guard before refreshing the plugin cache:

```bash
scripts/prepackage-check.mjs
```

Preview the exact new versioned cache directory, then install only after reviewing that path. The installer runs the source prepackage gate, creates the version directory atomically, refuses any existing target, preserves older cache versions, and requires full byte/mode/type/symlink parity before succeeding:

```bash
node scripts/install-versioned-cache.mjs
node scripts/install-versioned-cache.mjs --confirmed --target="$HOME/.codex/plugins/cache/personal/onedrive/0.5.0+codex.20260714034051"
```

After both live betas, regenerate the two QA reports, preview their exact sync into that new cache, then apply only those evidence files and re-run parity:

```bash
node scripts/install-versioned-cache.mjs --sync-evidence --target="$HOME/.codex/plugins/cache/personal/onedrive/0.5.0+codex.20260714034051"
node scripts/install-versioned-cache.mjs --sync-evidence --confirmed --target="$HOME/.codex/plugins/cache/personal/onedrive/0.5.0+codex.20260714034051"
```

Office compatibility checks are split by purpose:

```bash
python3 scripts/office-openxml-test.py
python3 scripts/office-security-test.py
node scripts/semantic-anchors-test.mjs
node scripts/text-patch-test.mjs
SOFFICE="$(command -v soffice)" python3 scripts/office-real-fixture-test.py
```

Install the pinned fixture dependencies with `python3 -m pip install -r scripts/requirements-office-test.txt`. The security corpus includes malformed ZIP/XML/relationship cases, deterministic mutation fuzzing, every possible two-run PowerPoint split of a target phrase, and a 5,000-run deck. The real-fixture gate generates packages with `python-docx`, `openpyxl`, and `python-pptx`, edits them, reopens them with their native libraries, and requires LibreOffice PDF conversion without repair/corruption diagnostics.

After installing a refreshed build, compare source with the installed cache:

```bash
scripts/prepackage-check.mjs --installed /path/to/installed/onedrive/cache
```

Running the harness without `--live` is read-only. It prints the exact proposed run ID and folder name. Review those values, then run the live CRUD/regression test from the plugin directory or with an absolute path to the installed plugin:

```bash
scripts/beta-test.mjs --live --confirmed --run-id=codex-beta-20260713t150000z --invite-recipient=person@example.com
```

All four live arguments are required. The test creates the exact named temporary OneDrive folder, exercises CRUD and safety behavior, silently grants then revokes read access for the explicit recipient, creates then revokes an anonymous test link, deletes only that test folder during cleanup, and removes isolated local work on success. Results are recorded as `pass`, `fail`, or `blocked`; a resource limitation is never reported as a pass. Pass `--keep-work` to keep local artifacts for debugging.
Unknown, duplicate, and positional CLI options are rejected before the beta harness starts, so a misspelled safety or mode flag cannot silently fall through to a different test mode.

Find old beta-test folders without deleting them. Cleanup discovery uses bounded Graph search followed by item verification, so it does not recursively scan the entire drive. Candidates with missing or invalid timestamps are skipped, and invalid/overflowing cleanup limits are rejected before any delete:

```bash
scripts/beta-test.mjs --cleanup-stale --stale-days=1
```

Delete the stale candidates only after reviewing the dry-run output:

```bash
scripts/beta-test.mjs --cleanup-stale --stale-days=1 --live --confirmed --run-id=codex-beta-cleanup-20260713
```

Run read-only tenant health checks across personal/work-school tenant endpoints:

```bash
scripts/beta-test.mjs --tenant-matrix=common,consumers,organizations
```

Use `--tenant-matrix-live --live --confirmed --run-id=<exact-id> --invite-recipient=<email>` only when you intentionally want to run the full live beta once per tenant entry.

## Plugin Gallery

The plugin manifest includes a file-manager flow screenshot at `assets/screenshot-file-manager.png` so the Codex plugin page shows the OneDrive search, read, upload, and safety workflow rather than only an icon.

![OneDrive plugin file manager flow](assets/screenshot-file-manager.png)

## CI

The GitHub Actions workflow in `.github/workflows/ci.yml` runs syntax checks, pinned Office fixture/security checks, semantic-anchor and structured-patch tests, task-pane and loopback-broker security tests, the Microsoft Graph regression suite, and the prepackage guard on Node.js 20 and 26 for every push and pull request.

## Microsoft References

- OneDrive files in Microsoft Graph: https://learn.microsoft.com/en-us/graph/api/resources/onedrive
- List folder children: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children
- Large file upload sessions: https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
- Move driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-move
- Copy driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-copy
- Create sharing link: https://learn.microsoft.com/en-us/graph/api/driveitem-createlink
- Invite recipients: https://learn.microsoft.com/en-us/graph/api/driveitem-invite
- Drive recipient resource: https://learn.microsoft.com/en-us/graph/api/resources/driverecipient
- Delta sync: https://learn.microsoft.com/en-us/graph/api/driveitem-delta
- Search drive items: https://learn.microsoft.com/en-us/graph/api/driveitem-search
- JSON batching: https://learn.microsoft.com/en-us/graph/json-batching
- driveItem resource fields: https://learn.microsoft.com/en-us/graph/api/resources/driveitem
- List permissions: https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions
- Get driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-get
- Restore deleted item: https://learn.microsoft.com/en-us/graph/api/driveitem-restore
- Microsoft Graph throttling guidance: https://learn.microsoft.com/en-us/graph/throttling
- Delegated Microsoft Graph auth: https://learn.microsoft.com/en-us/graph/auth-v2-user
- Microsoft identity scopes: https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc
