---
title: "Convert Between Markdown and Word"
description: "Create reviewable Word documents from Markdown and return DOCX content to Markdown with explicit image and formatting policies."
layout: docs
---

A Markdown-to-Word workflow is useful when authors prefer text-based source control but reviewers need comments, tracked changes, or a familiar DOCX file. The return conversion is best treated as semantic round-trip, not pixel-perfect fidelity.

## Create the review copy

`Convert-OfficeDocument -To Word` accepts a Markdown path, Markdown text (`-Text -From Markdown`), or a parsed document (`-Document`). Pass `New-OfficeMarkdownWordOptions` through `-Options` for template insertion, front matter, and image policies.

```powershell
$options = New-OfficeMarkdownWordOptions -Template .\Template.docx -BookmarkName MainContent
$report = Convert-OfficeDocument -Path .\Review.md -To Word -OutputPath .\Review.docx -Options $options
$report.Diagnostics
```

## Return to Markdown

`Convert-OfficeDocument -To Markdown` accepts a Word path or loaded document. Use `New-OfficeWordMarkdownOptions` for image export mode and directory, underline, and highlight output. Those formatting extensions are opt-in because Markdown targets represent them differently.

```powershell
$options = New-OfficeWordMarkdownOptions -ImageExportMode File -ImageDirectory .\Review-Images
Convert-OfficeDocument -Path .\Review.docx -To Markdown -OutputPath .\Review.md -Options $options
```

The [Word round-trip recipe](https://github.com/EvotecIT/OfficeIMO/blob/8efa0cdf9065c309e704080a560f2a686a144502/PowerShell/PSWriteOffice/Examples/Markdown/Recipe-Markdown-WordRoundTrip.ps1) creates Markdown, converts it to DOCX, and converts the Word file back to Markdown.

For heavily designed Word templates, define which semantic elements must survive before promising a round-trip. Tables, headings, lists, links, and basic emphasis are a better contract than exact pagination.
