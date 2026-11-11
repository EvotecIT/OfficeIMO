---
title: "OfficeIMO.Xps"
description: "Read, create, edit and save XPS and OpenXPS documents, and convert fixed pages to SVG, images and tagged vector PDF without Windows desktop APIs."
layout: product
meta.seo_title: "XPS and OpenXPS processing for .NET"
product_color: "#475569"
product_label: "XPS and OpenXPS engine"
install: "dotnet add package OfficeIMO.Xps"
nuget: "OfficeIMO.Xps"
docs_url: "/docs/xps/"
api_url: "/api/xps/"
meta.software.name: "OfficeIMO.Xps"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.Xps"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Fixed-page documents on every platform

`OfficeIMO.Xps` handles Microsoft XPS and ECMA-388 OpenXPS packages with managed code. It reads multiple fixed documents, creates pages with paths, embedded fonts, text and images, edits page order and native page XML, and saves in the original dialect. `OfficeIMO.Xps.Pdf` converts the result to vector PDF.

```csharp
using OfficeIMO;
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("invoice.xps");
foreach (XpsPage page in document.Pages)
    Console.WriteLine(page.ExtractText());

OfficeConversionReport report = document.SaveAsPdf("invoice.pdf");
```

## Useful for

- Converting printed XPS and OpenXPS output to searchable, tagged PDF.
- Rendering pages to SVG or PNG for previews.
- Extracting text and native logical structure for indexing.
- Assembling, splitting and reordering fixed documents while keeping links and outlines aligned.
- Producing fixed-layout pages with embedded fonts and images.

## Strict by default

Conversion rejects features outside the documented rendering profile instead of silently dropping them; an explicit partial SVG mode returns diagnostics. Loading is bounded, external resources and DTDs are rejected, and nothing is fetched from the network.

## Current limits

Signed packages can be read but not saved, and protected packages are not supported. Saves keep native page markup and opaque parts but rebuild ZIP packaging, so output is not byte-identical. PDF output does not claim PDF/UA or PDF/X conformance.

Start with the [XPS guide](/docs/xps/) or review the [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Xps/SUPPORT.md).
