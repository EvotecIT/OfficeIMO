---
title: "OpenDocument"
description: "Create, read and edit ODT, ODS and ODP files in .NET without LibreOffice, and convert them to Word, Excel, PowerPoint and PDF with a report of anything lost."
order: 38
verified: 2026-11-14
commands:
  - officeimo hidden-data inspect
  - officeimo hidden-data remove
meta.seo_title: "OpenDocument formats and APIs | OfficeIMO"
packages:
  - OfficeIMO.OpenDocument
cmdlets:
  - New-OfficeOpenDocument
  - Open-OfficeOpenDocument
related:
  - /docs/open-document/create-read/
  - /docs/open-document/convert-office/
  - /docs/open-document/convert-pdf/
  - /docs/security/sign-documents/
limits: Arbitrary producer-specific ODF signature profiles are inspected and preserved but not created; legacy Blowfish encryption fails before content is exposed.
search_terms:
  - odt
  - ods
  - odp
  - odg
  - libreoffice
  - open document format
---

`OfficeIMO.OpenDocument` creates and edits ODT, ODS, ODP and ODG files directly. It depends only on `OfficeIMO.Core` and never invokes LibreOffice, Microsoft Office or UNO. Install the adapter packages when a workflow crosses into Office formats or PDF; each brings only the engine it needs.

| Package | Adds |
|---|---|
| `OfficeIMO.OpenDocument` | `OdtDocument`, `OdsDocument`, `OdpPresentation`, `OdgDocument`, flat XML variants, encryption |
| `OfficeIMO.Word.OpenDocument` | ODT to and from DOCX, ODT page images |
| `OfficeIMO.Excel.OpenDocument` | ODS to and from XLSX, ODS page images |
| `OfficeIMO.PowerPoint.OpenDocument` | ODP to and from PPTX, ODP slide images |
| `OfficeIMO.OpenDocument.Odt.Pdf`, `.Ods.Pdf`, `.Odp.Pdf` | Each format to and from PDF |

```sh
dotnet add package OfficeIMO.OpenDocument
```

The task guides cover [creating, reading and editing files](/docs/open-document/create-read/), [Office conversion](/docs/open-document/convert-office/) and [PDF and image export](/docs/open-document/convert-pdf/).

## Create an ODT document

### C#

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created without LibreOffice or Office.");
document.Save("summary.odt");
```

### PowerShell

```powershell
New-OfficeOpenDocument -Path ./summary.odt {
    Add-OfficeOpenDocumentHeading -Text 'Summary' -Level 1
    Add-OfficeOpenDocumentParagraph -Text 'Created without LibreOffice or Office.'
}
```

## Remove hidden data before sharing

`OdfDocument.Load` opens any OpenDocument kind and returns the matching document type. The engine detects annotations, tracked changes, scripts, event listeners, external links, embedded objects, encryption and signatures, and never runs active content or fetches external resources.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OpenDocument;

using OdfDocument document = OdfDocument.Load("Report.odt");
Console.WriteLine($"{document.GetType().Name}");

OfficeHiddenDataReport report = document.RemoveHiddenData();
foreach (OfficeHiddenDataFinding finding in report.Findings) {
    Console.WriteLine($"{finding.Outcome} {finding.Kind} at {finding.Location}");
}
document.Save("Report-shared.odt");
```

`RemoveHiddenData` removes identity metadata, annotations, speaker notes, the thumbnail, printer settings and image metadata, anonymizes the authors of tracked changes you keep, and reports each item.

### CLI

```sh
officeimo hidden-data inspect Report.odt
officeimo hidden-data remove Report.odt --output Report-shared.odt
```

## Choose a format

| Format | Typical content |
|---|---|
| ODT / FODT | Reports and policies: paragraphs, headings, lists, tables, images, fields, notes, headers, footers, tracked paragraph changes |
| ODS / FODS | Typed spreadsheet data, formulas with bounded recalculation, merges, named ranges, validations, charts, print ranges |
| ODP / FODP | Presentations: text, shapes, images, tables, masters and layouts, speaker notes, transitions, basic animations |
| ODG / FODG | Drawings: shapes, paths, connectors, text boxes, images, layers |

Unknown XML, vendor extensions and embedded content are preserved when their part is not replaced. Conversions report what they could not carry, so a pipeline can enforce its own fidelity policy. The [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OpenDocument/README.md) lists the full editing surface.
