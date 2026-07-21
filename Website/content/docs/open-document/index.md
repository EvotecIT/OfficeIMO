---
title: "OpenDocument"
description: "Create and edit ODT, ODS, and ODP files and connect them to Office and PDF adapters."
order: 38
meta.seo_title: "OpenDocument formats and APIs | OfficeIMO"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/docs/open-document/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/docs/open-document/" />'
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

## Runtime boundary

The native package depends only on `OfficeIMO.Drawing`. It does not invoke LibreOffice, Microsoft Office, or UNO. Conversion diagnostics identify unsupported or approximated features so callers can enforce their own fidelity policy.

Browse the [OpenDocument API reference](/api/open-document/) for the complete model.
