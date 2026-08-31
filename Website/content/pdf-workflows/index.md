---
title: "PDF tools for browsers and .NET"
description: "Inspect, compare, organize, optimize, protect, unlock, and redact PDF files locally in your browser or through the OfficeIMO.Pdf API."
layout: page
slug: index
---

OfficeIMO gives the browser workspace and .NET applications the same first-party PDF engine. The browser tools run in WebAssembly, keep selected files in the current tab, and return downloadable artifacts with operation reports. Application code can use the corresponding `OfficeIMO.Pdf` APIs without a browser or Microsoft Office.

Choose the job you need. Each guide explains the browser workflow, the equivalent public API, the output it creates, and the limits that matter.

{{< pdf-workflows >}}

## What browser-local means

The PDF workspace does not upload selected files to an OfficeIMO service. Work is bounded by browser memory and the published input, page, comparison, split, and output limits. Server applications can choose their own `PdfLoadOptions` and resource policy.

For the complete engine, including authoring, rendering, forms, annotations, cryptographic signatures, and validation, see [OfficeIMO.Pdf](/products/pdf/) and the [PDF guides](/docs/pdf/).
