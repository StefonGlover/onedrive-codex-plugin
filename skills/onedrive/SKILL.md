---
name: onedrive
description: Work with files and folders in OneDrive through the local OneDrive Codex plugin. Use when the user asks to list, search, read, create, upload, download, rename, move, copy, share, or delete OneDrive files.
---

# OneDrive

Use the `onedrive` MCP tools to manage files in the signed-in user's OneDrive through Microsoft Graph.

## Configuration

The MCP server reads non-secret config from environment variables and `~/.codex/onedrive-plugin/config.json`.
Refresh tokens are stored in macOS Keychain.

- `ONEDRIVE_CLIENT_ID`: Microsoft app registration client ID for a public client app.
- `ONEDRIVE_TENANT`: Microsoft identity tenant. Defaults to `common`.
- `ONEDRIVE_SCOPES`: Space-separated delegated scopes. Defaults to `offline_access User.Read Files.ReadWrite`.
- `ONEDRIVE_KEYCHAIN_SERVICE`: Optional Keychain service name. Defaults to `Codex OneDrive`.

Never ask the user to paste Microsoft passwords, access tokens, or refresh tokens into chat. Use device-code login and Keychain-backed storage.

## Tool Guidance

- Start with `onedrive_config` to check whether a client ID and Keychain token are configured.
- Use `onedrive_auth_device_start` when no refresh token is stored. Ask the user to open the returned verification URL and enter the returned user code.
- Use `onedrive_auth_device_poll` after the user finishes browser sign-in.
- Use `onedrive_me` to confirm the signed-in account.
- Use `onedrive_presets` when the user refers to a friendly location such as documents, desktop, screenshots, or a configured alias.
- Use `onedrive_list` before changing or deleting remote files unless the target is already explicit.
- Use compact output by default for list, search, and delta results. Pass `format: "full"` only when richer metadata is needed.
- Use `onedrive_list_all` only when the user needs a complete folder listing; set a bounded `maxItems`.
- Use `onedrive_search` for file lookup before broad manual listing.
- Use `onedrive_search_all` only when the user needs paginated search results; set a bounded `maxItems`.
- Use `onedrive_delta` to answer what changed since a previous scan. Save `deltaLink` when present; use `nextLink` to continue an incomplete delta scan.
- Use `onedrive_get_info` with `includeDeletedItems: true` and an `itemId` when inspecting a deleted item ID on OneDrive Personal.
- Use `onedrive_read_text` only for bounded text files.
- If `onedrive_read_text` refuses likely binary content, use `onedrive_download` instead of forcing text unless the user specifically wants raw text extraction.
- Use `onedrive_download` for binary files or larger files.
- Use `onedrive_download_excel`, `onedrive_download_word`, or `onedrive_download_powerpoint` for Office/document-specific downloads when the user names a document type.
- Use `onedrive_upload` for local-file-to-OneDrive uploads. It automatically uses upload sessions for files above the simple upload limit; use `uploadMode: "session"` when explicitly testing resumable upload behavior.
- Use `onedrive_write_text` for creating or replacing small text files.
- Use `onedrive_move` and `onedrive_copy` when the user asks to reorganize files; include `expectedName` or `expectedId` when the target was resolved in a previous step.
- Use `onedrive_permissions` to audit sharing/permission state before creating or changing links.
- `onedrive_create_sharing_link` changes access to a file or folder. It defaults to dry-run; only set `dryRun: false` and `confirmed: true` after explicit user authorization for the exact item and link scope.
- `onedrive_restore_deleted` restores by deleted item ID and defaults to dry-run; only set `dryRun: false` and `confirmed: true` after explicit user authorization. Microsoft Graph does not provide a normal OneDrive recycle-bin listing API through driveItem file operations, so use `onedrive_delta` deleted changes to discover deleted item IDs when available.
- `onedrive_delete` defaults to `dryRun: true`; only set `dryRun: false` and `confirmed: true` after explicit user authorization. Live deletes require `expectedName` or `expectedId`.

## Maintenance

- Run `scripts/mock-graph-test.mjs` for fast regression coverage without touching OneDrive.
- Run `scripts/prepackage-check.mjs` before reinstalling the plugin cache.
- Run `scripts/prepackage-check.mjs --installed <installed-cache-path>` after reinstalling to catch source/cache drift.
- Run `scripts/beta-test.mjs` only when live OneDrive CRUD verification is intended.
