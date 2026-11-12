---
title: "Read and Inspect Word Documents"
description: "Open existing DOCX files, enumerate document structures, search content, and collect statistics without changing the source document."
layout: docs
---

Use the Word read surface when the document already exists and the job is to understand it: inventory paragraphs and tables, locate terms, inspect links or controls, or decide whether a later update is safe.

## Open read-only

`Open-OfficeWord -ReadOnly` returns a document that can be passed to the targeted `Get-OfficeWord*` commands, so the file is opened once even when several structures are inspected. Close it when finished.

```powershell
$document = Open-OfficeWord -Path '.\Policy.docx' -ReadOnly
$document | Get-OfficeWordParagraph
$document | Get-OfficeWordTable
$document | Get-OfficeWordHyperlink
(Get-OfficeDocumentInfo -Document $document -Detailed).Details.Statistics
Close-OfficeWord -Document $document
```

## Find content deliberately

Use `Get-OfficeWordParagraph -Pattern 'Quarter' -SimpleMatch` for literal text; omit `-SimpleMatch` for regular expressions. The inspection family also covers fields, bookmarks, comments and revisions, footnotes and endnotes, images, content controls, tables of contents, and document properties.

The complete [inspect-existing recipe](https://github.com/EvotecIT/OfficeIMO/blob/37d1d94ab3d6e7f84c6f3853ef07264061660d47/PowerShell/PSWriteOffice/Examples/Word/Recipe-Word-InspectExisting.ps1) creates a representative file, opens it once, and reports sections, paragraphs, tables, word count, and matches.

## Choose the next step

- Use [targeted updates](/docs/pswriteoffice/word-update-existing/) when the structure should stay intact.
- Use [merge and mail merge](/docs/pswriteoffice/word-merge-mailmerge/) when several documents or data records produce the result.
- Use the cross-format [Reader](/docs/pswriteoffice/reader/) when the same search must include Word, Excel, PowerPoint, PDF, Markdown, email, and other supported files.
