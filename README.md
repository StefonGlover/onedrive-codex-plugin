# OneDrive Codex Plugin

Local Codex plugin for OneDrive file operations through Microsoft Graph.

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
6. Ask Codex to call `onedrive_auth_device_start`, open the returned verification URL, enter the returned user code, then ask Codex to call `onedrive_auth_device_poll`.

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

## Safety

- Remote mutations that move, rename, copy, expose, invite, restore, delete, or revoke access use a preview-first pattern.
- Rename, move, copy, sharing-link creation, named-recipient invitation, permission revoke, restore, and delete default to dry-run where the operation has a dry-run mode.
- Live rename, move, copy, sharing-link creation, named-recipient invitation, permission revoke, restore, and delete require `dryRun: false`, `confirmed: true`, and stable expected identity (`expectedName` or `expectedId`; restore requires `expectedId`).
- Batch move and batch permission revoke preflight every item before any mutation and refuse partial execution when a preflight check fails.
- Sharing-link creation supports Microsoft Graph link type/scope plus optional password and expiration, and can include a before/after permission diff so the caller can see what changed.
- `onedrive_invite_permission` grants named users or groups access through Microsoft Graph `driveItem: invite`. It defaults to a silent direct grant (`sendInvitation: false`, `requireSignIn: true`); email invitations are opt-in with `sendInvitation: true` and optional `message`.
- Permission revoke uses Microsoft Graph `DELETE /me/drive/items/{item-id}/permissions/{permission-id}` and includes before permissions by default; live revoke includes after permissions and a permission diff.
- Rename, move, copy, share, and delete refuse to operate on the OneDrive root.
- Tool arguments are validated before handlers run, including required fields, unknown properties, enum values, numeric bounds, array bounds, and target `anyOf` rules.
- Text reads are bounded to 5 MB by default.
- Text reads use MIME/extension checks and refuse likely binary files unless `force: true` is set.
- Downloads go to `~/.codex/onedrive-plugin/downloads` unless `localPath` is provided.
- Downloads and uploads refuse local OneDrive sync-folder paths by default. Use `allowLocalOneDriveSyncPath: true` only for an explicit local sync-folder workflow.
- Uploads use simple upload for smaller files and upload sessions for large files, or when `uploadMode: "session"` is requested.
- List, search, find, scan, and delta tools return compact item summaries by default; pass `format: "full"` for richer metadata.
- Normal list, search, scan, delta, and metadata calls opportunistically maintain a local metadata cache at `~/.codex/onedrive-plugin/cache/metadata-cache.json`.
- `onedrive_sync_status` reports cache age, item count, delta cursor availability, and plugin storage locations.
- `onedrive_cache_refresh` rebuilds the cache from a bounded recursive scan and uses delta refreshes when a previous cursor exists. `onedrive_cache_clear` clears the cache.
- `onedrive_find` is the preferred file lookup helper. It uses the local metadata cache when available, runs live Graph search variants, ranks results in memory, and can fall back to bounded recursive remote scans. Cache-only hits must still have query relevance and are confirmed with live evidence before they suppress fallback scanning. Pass `useCache: false` for a fully live lookup.
- `onedrive_find_all` is the broader locator for “look everywhere” requests. It searches common folders first and uses larger bounded scan caps, with cache acceleration when available.
- `onedrive_preview` returns bounded text previews for text files and Graph-supported document text exports without reading unbounded remote content into memory.
- `onedrive_update_file` provides a checkout/commit edit workflow with a local manifest, eTag/cTag/size/mtime conflict checks, optional backup, and post-commit verification. Checkout refuses to overwrite an existing manifest unless `overwriteManifest: true` is provided.
- `onedrive_batch_get_info` and `onedrive_batch_permissions` use Microsoft Graph batching for up to 20 items. Batch download/delete/move tools provide one result per item with dry-run support where destructive.
- `onedrive_rename`, `onedrive_move`, and `onedrive_copy` support `dryRun: true` previews. Live `onedrive_batch_move` requires `dryRun: false`, `confirmed: true`, and `expectedName` or `expectedId` for every item.
- `onedrive_recent`, `onedrive_large_files`, `onedrive_duplicates`, `onedrive_shared_by_me`, and `onedrive_public_links` provide cleanup and sharing-audit workflows.
- `onedrive_list_all` follows pagination within one folder. Use `onedrive_scan` when you need recursive traversal across subfolders or the whole OneDrive.
- `onedrive_doctor` checks config, auth, profile, drive metadata, presets, and optional root listing in one call.
- `onedrive_export_pdf` and `onedrive_export_text` ask Microsoft Graph to convert supported Office files before saving locally. Microsoft Graph may reject conversions for unsupported file types.
- `onedrive_permissions` audits current sharing/permission grants before changing access.
- `onedrive_delta` can return deleted item changes. Microsoft Graph does not expose a normal OneDrive recycle-bin listing endpoint through the driveItem file APIs.
- `onedrive_get_info` supports `includeDeletedItems: true` when targeting an item ID; Microsoft documents this as OneDrive Personal-only.
- `onedrive_restore_deleted` defaults to dry-run and requires a deleted item ID. Live restore may require `Files.ReadWrite.All` for personal OneDrive.
- Live remote mutations are recorded in a local JSONL audit log at `~/.codex/onedrive-plugin/audit/mutations.jsonl`. Audit entries include safe item summaries, before/after summaries when available, permission diffs when relevant, Graph request IDs when available, and safe error details for failed live mutations. They do not log tokens, authorization headers, file contents, raw request bodies, sharing-link web URLs, passwords, invite messages, or recipient identifiers.
- `onedrive_audit_recent` reads recent audit entries, `onedrive_audit_export` exports the JSONL log to a local file, and `onedrive_audit_clear` requires `confirmed: true`.
- Graph requests retry transient `429`, `500`, `502`, `503`, and `504` responses with `Retry-After` support.

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

Run the prepackage guard before refreshing the plugin cache:

```bash
scripts/prepackage-check.mjs
```

After installing a refreshed build, compare source with the installed cache:

```bash
scripts/prepackage-check.mjs --installed /path/to/installed/onedrive/cache
```

Run the live CRUD/regression test from the plugin directory or with an absolute path to the installed plugin:

```bash
scripts/beta-test.mjs
```

The test creates a clearly named temporary OneDrive folder, exercises CRUD and safety behavior, deletes only that test folder during cleanup, and removes local temporary work on success. Pass `--keep-work` to keep local artifacts for debugging.

## Plugin Gallery

The plugin manifest includes a file-manager flow screenshot at `assets/screenshot-file-manager.png` so the Codex plugin page shows the OneDrive search, read, upload, and safety workflow rather than only an icon.

![OneDrive plugin file manager flow](assets/screenshot-file-manager.png)

## CI

The GitHub Actions workflow in `.github/workflows/ci.yml` runs syntax checks, the mocked Microsoft Graph regression suite, and the prepackage guard on every push and pull request.

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
- List permissions: https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions
- Get driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-get
- Restore deleted item: https://learn.microsoft.com/en-us/graph/api/driveitem-restore
- Microsoft Graph throttling guidance: https://learn.microsoft.com/en-us/graph/throttling
- Delegated Microsoft Graph auth: https://learn.microsoft.com/en-us/graph/auth-v2-user
- Microsoft identity scopes: https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc
