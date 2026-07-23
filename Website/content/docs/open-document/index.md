---
title: "OpenDocument"
description: "Create and edit ODT, ODS, and ODP files and connect them to Office and PDF adapters."
order: 38
meta.seo_title: "OpenDocument formats and APIs | OfficeIMO"
---

## Install

```shell
dotnet add package OfficeIMO.OpenDocument
```

## Create an ODT document

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created without LibreOffice or Office.");
document.Save("summary.odt");
```

The package also exposes `OdsDocument` for spreadsheets and `OdpDocument` for presentations. Add the focused Word, Excel, or PowerPoint OpenDocument adapter when the workflow needs bidirectional conversion; add `OfficeIMO.OpenDocument.Pdf` for PDF output.

## Choose the artifact

| Format | Typical workflow |
|---|---|
| ODT / FODT | Policies, reports, and document interchange with paragraphs, lists, tables, images, headers, footers, and tracked paragraph changes |
| ODS / FODS | Typed spreadsheet data, formulas, merges, named ranges, validation, and print ranges |
| ODP / FODP | Presentations with text, shapes, images, tables, notes, transitions, and basic animations |

Focused adapters connect ODT to Word, ODS to Excel, and ODP to PowerPoint. Conversion reports list rewritten or lossy entries so a pipeline can reject unexpected changes rather than relying on a successful save alone.

## Inspect before changing

OpenDocument can detect annotations, tracked changes, extension namespaces, scripts, event listeners, external links, embedded objects, formulas, validations, transitions, animations, encryption, and signatures. The engine does not execute active content or fetch external resources.

## Runtime boundary

The native package depends only on `OfficeIMO.Drawing`. It does not invoke LibreOffice, Microsoft Office, or UNO. Conversion diagnostics identify unsupported or approximated features so callers can enforce their own fidelity policy.

Browse the [OpenDocument API reference](/api/open-document/) for the complete model.
