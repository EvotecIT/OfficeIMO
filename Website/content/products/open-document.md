---
title: "OfficeIMO.OpenDocument"
hero_title_html: "OfficeIMO.<wbr>OpenDocument"
description: "Create and edit ODT, ODS, and ODP files and connect them to Office and PDF conversion workflows."
layout: product
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/open-document/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/open-document/" />'
product_label: "OpenDocument engine"
product_color: "#0f766e"
install: "dotnet add package OfficeIMO.OpenDocument"
nuget: "OfficeIMO.OpenDocument"
docs_url: "/docs/open-document/"
api_url: "/api/open-document/"
---

## OpenDocument without LibreOffice

Create and edit ODT text documents, ODS spreadsheets, and ODP presentations directly from .NET. The package does not invoke LibreOffice, Microsoft Office, or UNO.

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");
document.Save("summary.odt");
```

## Compose only what you need

| Package | Purpose |
|---|---|
| `OfficeIMO.OpenDocument` | Native ODT, ODS, and ODP models |
| `OfficeIMO.Word.OpenDocument` | Word and ODT conversion |
| `OfficeIMO.Excel.OpenDocument` | Excel and ODS conversion |
| `OfficeIMO.PowerPoint.OpenDocument` | PowerPoint and ODP conversion |
| `OfficeIMO.OpenDocument.Pdf` | PDF output through the matching Office adapters |

Converters report unsupported or approximated features so a workflow can validate fidelity before it publishes the result.
