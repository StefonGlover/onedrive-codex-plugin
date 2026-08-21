# Excel operation guidance

Use `onedrive_office_capabilities` with `kind: "excel"` and one `operation` to obtain its exact schema and example.

## Table rows

Use a two-dimensional rows array for `addTableRow.values`, even for one row:

```json
{
  "type": "addTableRow",
  "table": "Applications",
  "values": [["Applied", "Example Co"]]
}
```

Omit `index` to append. The server accepts a one-dimensional single-row shorthand but normalizes it to this canonical shape. Preview and inspect the table after commit. Refuse to guess a table, column order, or row index.

## Safety checks

- Preserve cell styles, row attributes, formula relationships, conditional-format ranges, and data-validation ranges unless the operation preview says otherwise.
- Inspect formulas and table/range identities before row deletion. Fail closed for the final data row, moved shared-formula masters, and array/data-table formulas.
- Use `backend: "auto"`. Let supported business `.xlsx` files use a scoped Graph workbook session and let consumer, macro-enabled, or unsupported operations use Open XML.
- Use `recalculate` only when calculation is part of the requested change. A successful Business/SharePoint Graph calculation can prove that Microsoft's workbook engine ran; a Personal/OpenXML recalculation only clears caches and marks calculation-on-open, so `calculationVerified` remains false.
- Verify changed cell values and formulas by structured readback. Do not claim a formula result without observed recalculation evidence.
- Render bounded sheets after chart, print-layout, size, merge, or formatting changes.
