---
name: onedrive-review
description: Review OneDrive Office files with structured evidence, rendered previews, comments or notes, recent activity, and version comparison or restore. Use for document feedback, comment management, visual QA, audit-style review, or guarded version recovery.
---

# OneDrive Review

Ground every review statement in current structured content, a rendered page/slide/sheet, or version metadata.

## Review content and layout

1. Inspect the target with `onedrive_office_inspect` using the correct `kind` and bounded selectors. Retain item ID, eTag, drive ID, coordinates, and durable anchors.
2. Render relevant pages, sheets, or slides with `onedrive_render_preview`. Use `onedrive_download_file` only when a materialized original or PDF is required for local QA.
3. Tie feedback to observable evidence: paragraph/table coordinates, sheet and range, slide and shape, page/slide image, or version timestamp.
4. Re-inspect after any accepted edit and distinguish verified changes from suggestions.

## Manage comments and notes

Use `onedrive_office_review` to list current review artifacts before adding or deleting one. Supply the exact inspected target and durable evidence anchor; never infer a location from prose alone. Preview mutating review operations when the tool exposes a preview proof, show the exact file/location/text or deletion target, obtain confirmation, then commit and list again.

Respect the tool's returned `limitations`. Do not claim support for replies, resolution state, author impersonation, PowerPoint threaded comments, or deletion on a format unless the current capability response explicitly advertises it. Do not simulate an unsupported comment by inserting visible body text.

## Inspect and restore versions

1. Use `onedrive_read_actions` with `versions`; compare a candidate with `compareVersion` before proposing restore.
2. Read the current file identity and eTag again immediately before preview.
3. Preview `restoreVersion` with `onedrive_preview_actions`. Show the exact current item and selected version; explain that restore changes the live file.
4. Obtain exact confirmation, then pass the unchanged action, preview token, stable identity, and current eTag to `onedrive_commit_actions`.
5. Inspect the live item and versions again. Never retry when the restore succeeded but local verification was incomplete.

Use `recent` in `onedrive_read_actions` for recent-file evidence. For enterprise libraries, discover with `drives` or `enterpriseSearch`, retain the selected `driveId`, and never silently substitute the personal drive.

Read [references/evidence-and-limitations.md](references/evidence-and-limitations.md) before publishing review findings or mutating comments/versions.
