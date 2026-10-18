---
title: "Update Existing Word Documents"
description: "Replace content and modify targeted Word structures while preserving the rest of an existing DOCX document."
layout: docs
---

Updating an existing document is different from rebuilding it. Use the smallest command that owns the intended change, preserve the original as input when auditability matters, and reopen the result to verify the change.

## Replace text and link metadata

`Update-OfficeWordText` can update visible text and, when explicitly selected, hyperlink text, URI, anchor, and tooltip metadata.

```powershell
Update-OfficeWordText -Path '.\Input\FY24-Report.docx' `
    -OldValue 'FY24' -NewValue 'FY25' `
    -IncludeHyperlinkText -IncludeHyperlinkUri `
    -IncludeHyperlinkAnchor -IncludeHyperlinkTooltip
```

Use `-WhatIf` before a broad replacement. `Find-OfficeWord` can count the old and new values before and after the operation.

## Modify live objects

For structural work, open an editable document and pipe the specific object to commands such as `Set-OfficeWordParagraph`, `Set-OfficeWordText`, `Set-OfficeWordTableCell`, `Set-OfficeWordImage`, or `Set-OfficeWordDocumentProperty`. Save or close with `-Save` only after all changes succeed.

The [update-existing recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-UpdateExisting.ps1) demonstrates link-aware replacement. [Example-WordModifyExistingObjects.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Example-WordModifyExistingObjects.ps1) shows object-level changes.

## Preserve intent

A targeted update should not silently restructure unrelated content. For large template changes, create a new version and use [Word comparison and review](/docs/pswriteoffice/word-review-delivery/) instead of treating a complex rewrite as a text replacement.
