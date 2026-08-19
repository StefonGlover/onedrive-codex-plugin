---
name: onedrive
description: Route OneDrive work through the local Microsoft Graph plugin. Use for finding, reading, uploading, organizing, sharing, deleting, restoring, reviewing, versioning, or working with Word, Excel, and PowerPoint files in OneDrive or SharePoint document libraries.
---

# OneDrive

Operate on the signed-in user's remote OneDrive through the plugin. Do not inspect or modify a local OneDrive sync folder unless the user explicitly requests local sync-folder access.

## Route the request

- Use `onedrive_read_actions` for batched discovery and metadata reads. Use `search` for descriptive lookup; use `list` for a known folder; use `recent`, `versions`, or `compareVersion` for timeline evidence.
- Use `onedrive_open_files` once for up to five exact filenames, known root-relative file paths, or observed HTTPS OneDrive/SharePoint file links. Pass exactly one of `names` or `urls`. Prefer a known path or observed link for deeply nested files; each resolves directly, while filename-only index misses share one bounded traversal. Present `displayLink` as returned so the visible hyperlink text is only the resolved filename; never print the bare URL or synthesize a link. Use `fetch` for a selected search result and pass its opaque `id` unchanged.
- Use `$onedrive-word`, `$onedrive-excel`, or `$onedrive-powerpoint` for format-aware Office creation or editing.
- Use `$onedrive-review` for comments, notes, versions, visual review, or evidence-backed feedback.
- Use `onedrive_write_text` or `onedrive_patch_text` for text/code files. Preview replacements and patches before any live write.
- Use `onedrive_create_folder` for a conflict-safe folder create. Use `onedrive_create_office_file` for bounded structured Word, Excel, or PowerPoint authoring directly in OneDrive, and `onedrive_upload_file` for a genuinely local or chat-provided file.
- Use `onedrive_preview_actions` and `onedrive_commit_actions` for rename, move, copy, permission revocation, and version restore. Use the dedicated preview-gated tool for other mutations such as sharing, delete, restore, text patch, upload replacement, or Office transforms.

## Preserve mutation safety

1. Resolve the exact remote target and retain its stable ID, current eTag, name, and drive ID when returned.
2. Run a dry-run or `onedrive_preview_actions` once.
3. Show the exact item, destination or recipient, and intended change. Ask for exact confirmation; do not accept a vague earlier “yes.”
4. Pass the preview token and expected identity back unchanged to the live call. Do not reuse a proof for a different target or operation.
5. Read or inspect the result again. Treat partial completion and verification warnings as real; never repeat a successful remote mutation merely to repair local follow-up state.

Return a link only when the provider response contains `webUrl`. Render it as `[resolved filename](webUrl)` so only the filename is visible. Never synthesize or guess a OneDrive or SharePoint URL.

## Work across enterprise libraries

Use `onedrive_read_actions` with `drives` to discover accessible drives, `enterpriseSearch` to search OneDrive and SharePoint, and `libraryList` to inspect a selected library. Preserve the returned `driveId` with every subsequent library target. Do not silently fall back to `/me/drive`, guess a site/library, or use a deprecated shared-with-me route. If tenant consent does not permit enterprise search, report the scope/capability limitation instead of broadening access.

## Load references conditionally

- Read [references/tool-profiles.yaml](references/tool-profiles.yaml) when validating routing or checking which focused tools and composite operations are allowed.
- Read [references/full-profile-maintenance.md](references/full-profile-maintenance.md) only when the runtime advertises the full Codex/local profile and the request concerns authentication, cache/index management, watches, backups, audit, diagnostics, or plugin maintenance.
