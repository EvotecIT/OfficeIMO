---
title: "OfficeIMO.Rtf"
description: "Read, create, edit, preserve, and convert Rich Text Format documents with bounded parsing and explicit fidelity reports."
layout: product
meta.seo_title: "RTF editing and conversion for .NET"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/rtf/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/rtf/" />'
product_color: "#0f766e"
product_label: "Rich Text Format engine"
install: "dotnet add package OfficeIMO.Rtf"
nuget: "OfficeIMO.Rtf"
docs_url: "/docs/rtf/"
---

## One engine for normalized and source-preserving RTF work

`OfficeIMO.Rtf` exposes an editable semantic document for normal authoring and a lossless syntax tree for surgical changes that must retain untouched producer data.

```csharp
using OfficeIMO.Rtf;

RtfReadResult read = RtfDocument.Load("input.rtf");
RtfDocument document = read.Document;
document.ReplaceText("Draft", "Approved");
File.WriteAllText("approved.rtf", document.ToRtf());
```

## Useful for

- Generating RTF reports, letters, and interchange documents.
- Editing text, bookmarks, images, metadata, headers, and selected syntax without Microsoft Word.
- Processing uploaded RTF with limits on depth, tokens, text, binary payloads, images, objects, and output blocks.
- Converting between RTF and Word, Markdown, HTML, PDF, or Reader results.
- Enforcing a no-loss policy from structured conversion diagnostics.

The parser supports legacy code pages used by Word and Outlook producers and retains unknown syntax on the lossless path. It never fetches external resources.

Start with the [RTF guide](/docs/rtf/) or inspect the [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.rtf-support-matrix.md).
