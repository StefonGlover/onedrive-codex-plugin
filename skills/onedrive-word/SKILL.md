---
name: onedrive-word
description: Create, inspect, copy from templates, edit, visually verify, download, and upload Microsoft Word documents in OneDrive. Use for .docx authoring, paragraph/table/header/footer changes, structured Word readback, or layout-sensitive Word work.
---

# OneDrive Word

Keep Word documents native `.docx` files and use structured evidence before editing.

## Create a document

1. Use the installed Documents artifact capability to author and validate a local `.docx`; do not fake a Word document with renamed plain text.
2. Upload it with `onedrive_upload_file`. Preview first if the destination could replace an existing file.
3. Read it back with `onedrive_office_inspect` using `kind: "word"` and bounded paragraph/table selectors.
4. Render representative pages with `onedrive_render_preview` when layout, pagination, tables, images, or headers matter.
5. Return only the observed provider `webUrl` from the upload/readback response.

## Start from a template or reference

Preview a `copy` with `onedrive_preview_actions`, show the exact source and destination, obtain confirmation, and apply it with `onedrive_commit_actions`. Inspect and edit the copy only. Preserve section setup, styles, headers/footers, content controls, images, and table structure unless the user explicitly requests changes.

## Edit an existing document

1. Inspect with `onedrive_office_inspect`; retain item ID, eTag, paragraph/table coordinates, and durable anchors.
2. Query `onedrive_office_capabilities` with `kind: "word"` and the intended `operation` when support is uncertain.
3. Preview `onedrive_office_batch_transform` with exact operations and the current identity. Review its semantic diff.
4. Ask for confirmation of the exact document and changes. Apply with the preview token and identity unchanged.
5. Inspect the changed ranges again. Render before and after when visual fidelity matters; use `onedrive_download_file` only when a materialized original or PDF is needed for local QA.

Never guess an anchor after ambiguity or drift; refresh the structured inspection. Do not claim tracked-change authoring, comment replies/resolution, or another unsupported feature unless the capability response explicitly advertises it.

Read [references/word-operations.md](references/word-operations.md) for operation targeting and validation rules.
