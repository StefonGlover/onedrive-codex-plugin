---
name: onedrive-powerpoint
description: Create, inspect, copy from templates, edit, visually verify, download, and upload Microsoft PowerPoint presentations in OneDrive. Use for .pptx authoring, slides, shapes, text, notes, images, tables, layout preservation, or structured deck readback.
---

# OneDrive PowerPoint

Keep presentations native `.pptx` files and treat visual verification as part of completion for layout-sensitive work.

## Create a presentation

1. Use the installed Presentations artifact capability to author and validate a local `.pptx`; do not upload images or plain text under a PowerPoint extension.
2. Upload it with `onedrive_upload_file`. Preview first if the destination could replace an existing file.
3. Read it back with `onedrive_office_inspect` using `kind: "powerpoint"` and bounded slide/shape selectors.
4. Render representative slides with `onedrive_render_preview` and inspect text fit, clipping, overlap, hierarchy, contrast, and image crop.
5. Return only the observed provider `webUrl`.

## Start from a template or reference

Preview a `copy` with `onedrive_preview_actions`, show the exact source and destination, obtain confirmation, and apply it with `onedrive_commit_actions`. Inspect and edit the copy only. Preserve master-derived layouts, theme, slide sizes, placeholders, geometry, notes, alt text, and z-order unless explicitly asked to change them.

## Edit an existing presentation

1. Inspect with `onedrive_office_inspect`; retain item ID, eTag, slide indexes, shape IDs, geometry, notes, and anchors.
2. Query `onedrive_office_capabilities` with `kind: "powerpoint"` and the intended `operation` for its exact schema.
3. Preview `onedrive_office_batch_transform` and review its slide/shape semantic diff.
4. Ask for confirmation of the exact deck and changes. Apply with the preview token and identity unchanged.
5. Inspect edited slides again and render them for visual QA. Use `onedrive_download_file` only when a materialized original or PDF is needed for local inspection.

Never claim animation, master editing, threaded comments, or another unsupported feature unless the capability response explicitly advertises it.

Read [references/powerpoint-operations.md](references/powerpoint-operations.md) for targeting and visual QA rules.
