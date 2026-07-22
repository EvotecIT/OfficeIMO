---
title: Migration from the preview APIs
description: Adopt the split Google Workspace packages and explicit fidelity, import, and replacement contracts.
order: 90
---

The completed package family intentionally breaks misleading preview shapes.

## Package ownership

- Keep sessions, transport, scopes, diagnostics, and fidelity contracts in `OfficeIMO.GoogleWorkspace`.
- Add `OfficeIMO.GoogleWorkspace.Drive` for direct Drive work, `Auth.GoogleApis` only for Google SDK credentials, and `Sync` only for change tracking or plan/apply.
- Use `Word.GoogleDocs`, `Excel.GoogleSheets`, and `PowerPoint.GoogleSlides` for format mapping.

## Replace removed promise-only options

Old booleans such as Docs flatten/rasterize/comment switches and Sheets chart/pivot switches did not execute the promised behavior. Use `UnsupportedFeatures`, `GoogleWorkspaceFidelityPolicy`, the format support catalogs, and the executed fallback modes instead.

## Existing-file updates

Import/read the target first. Docs and Slides require the observed API revision; Sheets requires the observed Drive version. Pass it through `Replace`. Choose an overwrite mode only when last-writer-wins is an application decision.

## Imports and image safety

Use native import for editable target semantics and Drive Office-format export for broad fallback. Docs temporary image publication is explicit; Slides owns short-lived image leases internally. Do not retain or create permanent public staging objects.

Regenerate the [support matrix](/docs/google-workspace/support/) after changing any code-owned feature catalog.
