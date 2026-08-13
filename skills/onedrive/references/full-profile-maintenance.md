# Full-profile maintenance

Read this reference only when `tools/list` advertises the relevant tool. Hidden full-profile tools are not callable from the focused ChatGPT profile.

## Authentication and diagnosis

- Check `onedrive_config({checkToken: true})` when authentication state matters. Reuse a healthy stored refresh token silently.
- Run `onedrive_doctor` for setup, authentication, Graph access, or preset failures. Use `checkPresets: false` only for a deliberate auth-only check.
- Start `onedrive_auth_device_start` only after a genuine missing-credential or Microsoft reauthentication result. Never convert timeouts, throttling, or transient network failures into login prompts. Poll with `onedrive_auth_device_poll` after browser sign-in.
- Use `onedrive_logout` only when the user explicitly requests disconnect or credential reset. Never ask for passwords, access tokens, refresh tokens, or encryption keys in chat.
- Use `onedrive_me`, `onedrive_drive`, and `onedrive_presets` to verify the account, drive, and friendly paths.

Non-secret settings resolve from environment variables, then `~/.codex/onedrive-plugin/config.json`, then defaults. Relevant settings include `ONEDRIVE_CLIENT_ID`, `ONEDRIVE_TENANT`, `ONEDRIVE_SCOPES`, `ONEDRIVE_KEYCHAIN_SERVICE`, `ONEDRIVE_TOKEN_STORE`, `ONEDRIVE_TOKEN_FILE`, and `ONEDRIVE_TOKEN_ENCRYPTION_KEY_FILE`. Keep tokens in Keychain on macOS or the encrypted file store with an owner-only key on Linux/NAS.

## Cache, index, and monitoring

- Inspect `onedrive_sync_status` before repairing local state. Use `onedrive_cache_refresh` for repeated broad discovery and `onedrive_cache_clear` only when stale metadata is suspected.
- Build extracted-text search deliberately with bounded `onedrive_content_index_refresh`; query it with `onedrive_content_search`; clear it with `onedrive_content_index_clear` when requested or stale. Remember that the local index contains extracted file bodies.
- Use `onedrive_office_index_refresh` and `onedrive_office_search` for structured paragraph, cell, formula, table, shape, and notes research.
- Use `onedrive_watch_start`, `onedrive_watch_status`, and `onedrive_watch_stop` for bounded delta monitoring.
- Use `onedrive_delta` for change history and persist only the returned continuation or delta link for the same scoped root.

## Full-profile file and Office operations

- Use `onedrive_find` or `onedrive_find_all` for ranked broad lookup, `onedrive_scan` for bounded recursive traversal, and paginated `onedrive_list_all` or `onedrive_search_all` only when completeness is required.
- Use `onedrive_download`, `onedrive_download_word`, `onedrive_download_excel`, or `onedrive_download_powerpoint` for local processing. Reject local sync-folder paths unless the user explicitly permits them.
- Use `onedrive_word_get_document`, `onedrive_excel_get_workbook`, or `onedrive_powerpoint_get_presentation` before the matching batch update. Require a unique durable anchor or an exact positional selector; refresh on missing or ambiguous anchors.
- Use `onedrive_office_validate` before format-sensitive writes. Treat markup-compatibility errors as fail-closed.
- Use `onedrive_office_backups` and `onedrive_office_compare_backup` for recovery evidence. Preview `onedrive_office_restore_backup`; require the original stable ID and current eTag for restore.
- Use `onedrive_versions` and `onedrive_compare_version` before preview-gated `onedrive_restore_version`.

## Audit and maintenance

- Use `onedrive_audit_recent` or `onedrive_audit_export` to inspect mutation history. Clear with `onedrive_audit_clear` only on an explicit request and with confirmation.
- Run `scripts/mock-graph-test.mjs` for fast non-live regression coverage and `scripts/prepackage-check.mjs` before packaging. Keep live beta artifacts under the ignored `work/` tree.
- Treat plugin-managed caches, indexes, audit data, exports, backups, downloads, and workspaces as private per-account/per-drive state. Do not copy state between authentication scopes.
