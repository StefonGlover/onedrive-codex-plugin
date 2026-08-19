---
name: onedrive-excel
description: Create, inspect, copy from templates, edit, visually verify, download, and upload Microsoft Excel workbooks in OneDrive. Use for .xlsx authoring, ranges, formulas, tables, sheets, notes, charts, or structured workbook readback.
---

# OneDrive Excel

Keep workbooks native `.xlsx` files and preserve formulas, styles, validations, tables, and charts unless the user requests otherwise.

## Create a workbook

1. Use `onedrive_create_office_file` with `kind: "excel"` for bounded structured sheets and rows; it builds and validates a real `.xlsx`, previews the exact destination, and verifies the uploaded package.
2. For advanced formatting, charts, validation, print layout, or model design, use the installed Spreadsheets artifact capability to create and validate a local `.xlsx`, then upload it with `onedrive_upload_file`. Do not upload CSV under an Excel extension.
3. Read it back with `onedrive_office_inspect` using `kind: "excel"` and explicit sheet/range bounds.
4. Render representative sheets with `onedrive_render_preview` when print layout, charts, widths, merged cells, or formatting matter.
5. Return only the observed provider `webUrl`.

## Start from a template or reference

Preview a `copy` with `onedrive_preview_actions`, show the exact source and destination, obtain confirmation, and apply it with `onedrive_commit_actions`. Inspect and edit the copy only. Preserve sheet names/order, formulas, named ranges, tables, validation, conditional formatting, charts, and print settings unless explicitly asked to change them.

## Edit an existing workbook

1. Inspect with `onedrive_office_inspect`; retain item ID, eTag, sheet identity, ranges, table names, and anchors.
2. Query `onedrive_office_capabilities` with `kind: "excel"` and the intended `operation` for its exact schema.
3. Preview `onedrive_office_batch_transform` and review affected sheets, cells, formulas, tables, and backend choice.
4. Ask for confirmation of the exact workbook and edits. Apply with the preview token and identity unchanged.
5. Inspect the affected ranges again. Render before and after when visual fidelity matters; use `onedrive_download_file` only for a materialized original or PDF needed by local QA.

Do not assume formulas were evaluated server-side. Treat formula-cache clearing and recalculation-on-open as distinct from a computed result. Never force the Graph workbook backend for a consumer workbook.

Read [references/excel-operations.md](references/excel-operations.md) before table-row, formula, validation, merge, or backend-sensitive edits.
