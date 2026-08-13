# Review evidence and limitations

## Evidence hierarchy

1. Prefer a current structured inspection with stable ID and eTag.
2. Use exact paragraph/table, sheet/range, or slide/shape coordinates plus a unique anchor.
3. Use a bounded rendered preview for layout observations and identify the rendered page, sheet, or slide.
4. Use version IDs and timestamps returned by Microsoft Graph for history claims.
5. Label an inference as an inference. Refresh stale or ambiguous evidence instead of guessing.

## Mutation guards

- List before add/delete and list again after commit.
- Keep comment/note text faithful to the user's approved wording.
- Confirm the exact artifact, location, and action immediately after preview.
- Require a current eTag for version restore and reject target drift.
- Treat partial completion and verification warnings as unresolved follow-up, not permission to repeat a successful mutation.

## Unsupported behavior

Treat the capability/limitations response as authoritative. Report unavailable replies, resolution, deletion, PowerPoint comments, tracked-change authoring, or other review features plainly. Never emulate them with body text, hidden shapes, cells, or fabricated metadata.
