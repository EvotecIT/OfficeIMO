---
title: "Document Conversion Guides for .NET"
description: "Choose .NET routes for DOC, DOCX, XLS, XLSX, XLSB, PPT, PPTX, PDF, and web formats, with documented PSWriteOffice routes called out separately."
layout: solution
meta.eyebrow: "Conversion workflows"
meta.outcome: "Move documents without hiding fidelity decisions"
meta.primary_label: "Check format support"
meta.primary_url: "/compatibility/"
meta.social_card_badge: "Office conversion"
---

OfficeIMO conversion is built for .NET applications that need more than a file-extension switch. Each focused package owns a document family, while adapters add PDF, HTML, Markdown, OpenDocument, Google Workspace, and normalized extraction only when the workflow needs them.

## Start with the route you need

| Input and output | Best starting point | What you can inspect |
|---|---|---|
| [DOC and DOCX](/convert/doc-docx/) | `OfficeIMO.Word` | Feature-level compatibility, preservation, and conversion diagnostics |
| [XLS and XLSX](/convert/xls-xlsx/) | `OfficeIMO.Excel` | Worksheets, formulas, styles, charts, and legacy BIFF decisions |
| [XLSB and XLSX](/convert/xlsb-xlsx/) | `OfficeIMO.Excel` | Native binary-workbook parsing and write diagnostics |
| [PPT and PPTX](/convert/ppt-pptx/) | `OfficeIMO.PowerPoint` | Slides, shapes, text, media, notes, charts, and fallback decisions |
| [Word to PDF](/convert/word-to-pdf/) | `OfficeIMO.Word.Pdf` | Publishing warnings, fonts, images, tables, sections, and output status |
| [Excel to PDF](/convert/excel-to-pdf/) | `OfficeIMO.Excel.Pdf` | Worksheet selection, pagination, page setup, styles, and export warnings |

For HTML, Markdown, RTF, OpenDocument, Google Docs, Sheets, and Slides, use the focused adapter shown in the [package catalog](/downloads/). This keeps dependencies and deployment behavior explicit.

## PowerShell routes are documented separately

PSWriteOffice exposes dedicated DOC/DOCX and XLS/XLSX conversion cmdlets and a much broader set of authoring, inspection, validation, and publishing commands. It does not imply that every .NET conversion route in the table is available as a cmdlet. Start with the [PSWriteOffice product page](/products/pswriteoffice/), then verify the exact command and parameters in the [PowerShell API reference](/api/powershell/).

## Conversion is a policy decision

A document can be readable even when every source feature is not editable in the destination. OfficeIMO distinguishes:

- **Native** content represented directly by the destination format.
- **Editable approximation** when the closest destination feature is still useful.
- **Visual fallback** when appearance matters more than editability.
- **Preserved source** when recovery is safer than pretending a conversion was complete.
- **Blocked output** when continuing would silently misrepresent the result.

The [format compatibility page](/compatibility/) exposes those states from the same contracts used by the libraries. For governed migrations, analyze first, decide which states your workflow accepts, then write the output.

## Browser, server, and desktop choices

The browser playground intentionally supports a smaller modern-format route set because files stay in the browser. Legacy DOC, XLS, and PPT migration belongs in a .NET service, desktop application, or CLI. PSWriteOffice additionally provides the documented DOC/DOCX and XLS/XLSX cmdlet routes described above. That boundary is explicit so a browser demo—or the .NET package matrix—is never mistaken for the PowerShell command surface.

## Need a broader workflow?

- [Browse all solution guides](/solutions/)
- [Modernize a legacy Office archive](/solutions/legacy-office-modernization/)
- [Publish Office content to PDF and web formats](/solutions/document-publishing/)
- [Normalize documents for search and AI](/solutions/document-ingestion/)
- [Automate recurring document jobs with PowerShell](/solutions/powershell-document-automation/)
