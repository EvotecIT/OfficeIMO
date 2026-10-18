---
title: "Review, Protect, and Deliver Word Documents"
description: "Inspect revisions, compare Word documents, resolve approved changes, protect output, update fields, and prepare DOCX files for delivery."
layout: docs
---

Treat review and delivery as explicit stages after content generation or modification. Inspection should not accept revisions accidentally, and protection should not be used as a substitute for validation.

## Review changes

`Compare-OfficeWordDocument` produces comparison evidence between a reference and a candidate. `Get-OfficeWordReview` exposes comments and revisions. Use `Resolve-OfficeWordRevision` only after the acceptance rule is known; filter by author, revision type, or review decision instead of accepting everything by default.

## Finalize document state

Before delivery:

1. update fields and the table of contents;
2. inspect comments and unresolved revisions;
3. verify required controls, links, and document properties;
4. apply protection when recipients should not freely edit the result;
5. save to the delivery path and reopen it read-only.

`Protect-OfficeWordDocument` and `Unprotect-OfficeWordDocument` make the protection boundary explicit. `Update-OfficeWordFields` and `Update-OfficeWordTableOfContents` handle calculated presentation state.

The [approval-checklist recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-ApprovalChecklist.ps1) demonstrates content controls and a watermark. Combine it with [read and inspect](/docs/pswriteoffice/word-read-inspect/) for delivery readback.
