# PowerPoint operation guidance

Use `onedrive_office_capabilities` with `kind: "powerpoint"` and one `operation` to retrieve its exact schema and example.

- Ground every change in an inspected slide index and shape ID or a unique durable anchor.
- Preserve geometry, alt text, notes, grouping, z-order, crop, and layout relationships unless the request targets them.
- Use the narrowest operation. Avoid replacing whole slide XML or flattening editable content.
- Fail closed on missing/ambiguous anchors, eTag drift, package-validation failures, or markup-compatibility errors.
- Inspect changed shape text, geometry, alt text, table cells, images, and notes after commit.
- Render bounded affected slides before and after changes. Check clipping, overflow, collisions, alignment, margins, contrast, and crop; correct failures through another preview-confirm cycle.
