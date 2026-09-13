---
title: "PDF tools for browsers and .NET"
description: "Inspect, compare, organize, optimize, protect, unlock, and redact PDF files locally in your browser or through the OfficeIMO.Pdf API."
layout: conversion
slug: index
meta.eyebrow: "PDF workflows"
meta.outcome: "Choose a focused operation and know what it changes before you run it"
meta.source_format: "PDF files"
meta.destination_format: "Report, PDF copy, ZIP, or comparison gallery"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_label: "Open PDF tools in the browser"
meta.primary_url: "/convert/?workspace=pdf"
meta.secondary_label: "Read the .NET PDF guides"
meta.secondary_url: "/docs/pdf/"
meta.summary_title: "PDF workflow scope"
meta.limit: "Browser tools use published file, page, comparison, split, and output limits; .NET applications choose their own bounded load policy."
meta.related_label: "Explore OfficeIMO.Pdf"
meta.related_url: "/products/pdf/"
---

OfficeIMO gives the browser workspace and .NET applications the same first-party PDF engine. The browser tools run in WebAssembly, keep selected files in the current tab, and return downloadable artifacts with operation reports. Application code can use the corresponding `OfficeIMO.Pdf` APIs without a browser or Microsoft Office.

Choose the job you need. Each guide explains the browser workflow, the equivalent public API, the output it creates, and the limits that matter.

{{< pdf-workflows >}}

## What browser-local means

The PDF workspace does not upload selected files to an OfficeIMO service. Work is bounded by browser memory and the published input, page, comparison, split, and output limits. Server applications can choose their own `PdfLoadOptions` and resource policy.

For the complete engine, including authoring, rendering, forms, annotations, cryptographic signatures, and validation, see [OfficeIMO.Pdf](/products/pdf/) and the [PDF guides](/docs/pdf/).
