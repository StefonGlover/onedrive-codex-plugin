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
- `onedrive_delta`
- `onedrive_get_info`
- `onedrive_read_text`
- `onedrive_download`
- `onedrive_download_excel`
- `onedrive_download_word`
- `onedrive_download_powerpoint`
- `onedrive_upload`
- `onedrive_write_text`
- `onedrive_create_folder`
- `onedrive_rename`
- `onedrive_move`
- `onedrive_copy`
- `onedrive_create_sharing_link`
- `onedrive_permissions`
- `onedrive_restore_deleted`
- `onedrive_delete`

## Safety

- Delete defaults to dry-run.
- Live delete requires `confirmed: true` plus `expectedName` or `expectedId`.
- Sharing-link creation defaults to dry-run and requires both `dryRun: false` and `confirmed: true`.
- Rename, move, copy, share, and delete refuse to operate on the OneDrive root.
- Text reads are bounded to 5 MB by default.
- Text reads use MIME/extension checks and refuse likely binary files unless `force: true` is set.
- Downloads go to `~/.codex/onedrive-plugin/downloads` unless `localPath` is provided.
- Downloads and uploads refuse local OneDrive sync-folder paths by default. Use `allowLocalOneDriveSyncPath: true` only for an explicit local sync-folder workflow.
- Uploads use simple upload for smaller files and upload sessions for large files, or when `uploadMode: "session"` is requested.
- List, search, find, scan, and delta tools return compact item summaries by default; pass `format: "full"` for richer metadata.
- `onedrive_find` is the preferred file lookup helper. It is stateless and remote-first: it runs live Graph search variants, ranks results in memory, and can fall back to bounded recursive remote scans without creating a local index or persistent cache.
- `onedrive_list_all` follows pagination within one folder. Use `onedrive_scan` when you need recursive traversal across subfolders or the whole OneDrive.
- `onedrive_permissions` audits current sharing/permission grants before changing access.
- `onedrive_delta` can return deleted item changes. Microsoft Graph does not expose a normal OneDrive recycle-bin listing endpoint through the driveItem file APIs.
- `onedrive_get_info` supports `includeDeletedItems: true` when targeting an item ID; Microsoft documents this as OneDrive Personal-only.
- `onedrive_restore_deleted` defaults to dry-run and requires a deleted item ID. Live restore may require `Files.ReadWrite.All` for personal OneDrive.
- Graph requests retry transient `429`, `500`, `502`, `503`, and `504` responses with `Retry-After` support.

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

The test creates a clearly named temporary OneDrive folder, exercises CRUD and safety behavior, and deletes only that test folder during cleanup.

## Microsoft References

- OneDrive files in Microsoft Graph: https://learn.microsoft.com/en-us/graph/api/resources/onedrive
- List folder children: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children
- Large file upload sessions: https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
- Move driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-move
- Copy driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-copy
- Create sharing link: https://learn.microsoft.com/en-us/graph/api/driveitem-createlink
- Delta sync: https://learn.microsoft.com/en-us/graph/api/driveitem-delta
- List permissions: https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions
- Get driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-get
- Restore deleted item: https://learn.microsoft.com/en-us/graph/api/driveitem-restore
- Microsoft Graph throttling guidance: https://learn.microsoft.com/en-us/graph/throttling
- Delegated Microsoft Graph auth: https://learn.microsoft.com/en-us/graph/auth-v2-user
- Microsoft identity scopes: https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc
