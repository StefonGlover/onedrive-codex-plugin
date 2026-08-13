# Word operation guidance

Use `onedrive_office_capabilities` with `kind: "word"` and one `operation` to retrieve its exact schema and example before constructing an unfamiliar edit.

- Ground text edits in inspected paragraph indexes or unique durable anchors.
- Ground table edits in table, row, and column coordinates returned by inspection.
- Preserve formatting and structure by choosing the narrowest operation instead of replacing whole package parts.
- Treat a supplied position and anchor as a consistency check: both must resolve to the same object.
- Fail closed when an anchor is missing or ambiguous, the eTag changed, package validation fails, or markup-compatibility errors are present.
- Inspect affected paragraphs, tables, headers, footers, bookmarks, content controls, or images after commit.
- Render bounded pages after layout-sensitive changes and compare them with the pre-edit preview.

Use `$onedrive-review` rather than embedding review feedback into body text.
