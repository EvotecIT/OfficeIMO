---
title: "Open Formats and Text Automation"
description: "Use the smaller PSWriteOffice families for text, interchange, open formats, and managed message workflows. Includes examples and cmdlet links."
layout: docs
aliases:
  - /docs/pswriteoffice/markdown/
---

PSWriteOffice is not limited to the three desktop Office formats. Smaller command families expose the same managed-engine approach for text, interchange, open document, and email workflows.

## Markdown

Twenty-six commands build and inspect typed Markdown. Add headings, paragraphs, lists, task lists, tables, code, callouts, details, front matter, images, quotes, definition lists, and tables of contents. Reader commands expose headings, nodes, tables, and front matter; converters bridge HTML and Word workflows.

Start with the [operations runbook](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-OperationsRunbook.ps1) for operational content or the [release notes recipe](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-ReleaseNotes.ps1) for publishable change documentation. The [DSL cookbook](/docs/pswriteoffice/dsl-cookbook/) includes both and shows how the same data can also produce Word, Excel, PowerPoint, and PDF output.

The [object composition recipe](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-ObjectComposition.ps1) shows explicit document targets, while the [definition guide](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Markdown/Recipe-Markdown-DefinitionGuide.ps1) demonstrates definition lists and reusable reference content.

Use the focused Markdown guides for complete workflows:

- [Read and parse Markdown](/docs/pswriteoffice/markdown-read-parse/)
- [Convert and publish Markdown](/docs/pswriteoffice/markdown-convert-publish/)
- [Convert between Markdown and Word](/docs/pswriteoffice/markdown-word-roundtrip/)

## RTF

Six canonical commands create, load, update, convert, inspect, and configure PDF export for Rich Text Format documents. Use RTF when a lightweight rich-text interchange file is the required source or destination, and keep loss-aware conversion diagnostics for complex content.

See [update and convert RTF](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Rtf/Recipe-Rtf-UpdateAndConvert.ps1) for an end-to-end example.

## CSV

Five commands convert, import, export, and inspect CSV through OfficeIMO.CSV. Use the CSV family for delimited-data contracts; use Excel when worksheet formatting, formulas, charts, or workbook structure are part of the outcome.

See [safe CSV export](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Csv/Recipe-Csv-SafeExport.ps1) for delimiter, quoting, and round-trip validation choices.

## OpenDocument

OpenDocument commands create, read, convert, and save ODT, ODS, and ODP artifacts through OfficeIMO.OpenDocument. These are native managed workflows rather than LibreOffice automation. Creation is composable from PowerShell: ODT supports headings and paragraphs, ODS supports sheets and typed cells, and ODP supports slides and positioned text boxes.

```powershell
New-OfficeOpenDocument -Kind Text -Path .\Report.odt -Content {
    Add-OfficeOpenDocumentHeading -Text 'Service report' -Level 1
    Add-OfficeOpenDocumentParagraph -Text 'Generated without desktop Office or LibreOffice.'
}

New-OfficeOpenDocument -Kind Spreadsheet -Path .\Status.ods -Content {
    Add-OfficeOpenDocumentSheet -Name Services -Content {
        Set-OfficeOpenDocumentCell -Row 0 -Column 0 -Value 'Service'
        Set-OfficeOpenDocumentCell -Row 0 -Column 1 -Value 'Healthy'
        Set-OfficeOpenDocumentCell -Row 1 -Column 0 -Value 'Directory'
        Set-OfficeOpenDocumentCell -Row 1 -Column 1 -Value $true
    }
}
```

Use `New-OfficeWordOpenDocumentOptions`, `New-OfficeExcelOpenDocumentOptions`, or `New-OfficePowerPointOpenDocumentOptions` when conversion needs explicit fidelity or resource controls. The raw OfficeIMO option parameters remain available as an advanced escape hatch.

## Email

Four artifact commands load and save messages and mailbox files through OfficeIMO.Email, and five option builders expose their safety and fidelity policies. The underlying engine covers multiple message, personal-information, store, and address-book families; exact support and diagnostics belong to the generated command/API reference.

The module boundary is deliberate. PSWriteOffice treats email as document content: it reads or writes supported artifacts and lets OfficeIMO.Reader normalize mail sources for mixed-format search and reporting. [Mailozaurr](https://github.com/EvotecIT/Mailozaurr) owns transport, authentication, mailbox/store lifecycle, PST/OST import and conversion, querying, export, and delivery. Workflows can use Mailozaurr to acquire or deliver content and pass ordinary paths or attachments to PSWriteOffice without either module duplicating the other's operational responsibilities. The [PDF delivery recipe](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/Integrations/Recipe-Mailozaurr-PdfDelivery.ps1) is a runnable handoff: PSWriteOffice creates the attachment and Mailozaurr sends it. It uses `-WhatIf` unless `-Send` is explicitly supplied.

Advanced safety and fidelity controls are also PowerShell-native. Use `New-OfficeEmailReaderOptions`, `New-OfficeEmailWriterOptions`, `New-OfficeEmailStoreReaderOptions`, `New-OfficeEmailMailboxReaderOptions`, and `New-OfficeEmailMailboxWriterOptions`; their output binds directly to the matching `-Options` or `-StoreOptions` parameter. You do not need a hashtable, `New-Object`, or an OfficeIMO constructor.

## AsciiDoc and LaTeX

Each family provides four bounded interoperability commands for reading, saving, and bridging through Markdown. These are explicit profiles, not a claim to implement every extension or package in the wider AsciiDoc or TeX ecosystems.

## HTML review

Office format families provide focused HTML conversion commands, and `Export-OfficeHtmlImage` supports extracted assets. Use the [HTML review examples](https://github.com/EvotecIT/PSWriteOffice/tree/v3.0.6/Examples) to build a browser-review step without discarding the source document.
