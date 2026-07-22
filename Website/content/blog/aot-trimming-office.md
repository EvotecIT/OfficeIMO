---
title: "Shipping OfficeIMO Applications with NativeAOT"
description: "A customer-focused guide to publishing native OfficeIMO applications, choosing typed APIs, and validating real document workflows."
date: 2025-11-01
tags: [aot, trimming, performance]
author: OfficeIMO Team
---

NativeAOT is useful for small command-line tools, containerized document workers, and services where startup time and deployment size matter. OfficeIMO's standard in-process workflows are designed for that model: select the focused packages you need, publish your application for a concrete runtime, and test the documents your service will actually handle.

## What works as a native executable

OfficeIMO maintains native test applications for the most common customer paths. All eight scenarios currently pass on both Windows and Linux:

- Word creates, saves, reopens, and inspects a DOCX.
- Excel writes and reloads a typed table.
- PowerPoint creates and duplicates a chart slide with its related parts.
- Markdown and CSV compose or parse real content.
- Reader registers the complete local-format preset and performs structured extraction.
- HTML rendering produces SVG, PNG, and searchable PDF output.

The tests publish and execute on `win-x64` and `linux-x64` rather than stopping when compilation succeeds. This catches missing metadata, relationship, serialization, and startup behavior that a project setting alone cannot prove. Each scenario uses isolated SDK artifacts so results do not depend on a previous build from another operating system.

## Add NativeAOT to your application

```xml
<PropertyGroup>
  <TargetFramework>net10.0</TargetFramework>
  <PublishAot>true</PublishAot>
</PropertyGroup>
```

Then publish for the target that will run the application:

```powershell
dotnet publish -c Release -r linux-x64
```

Use the normal typed OfficeIMO APIs. If your code asks OfficeIMO or another dependency to discover arbitrary runtime model members, the .NET analyzer may require an explicit model mapping or preservation rule. That warning belongs at the application boundary where the dynamic type is selected.

## Optional integrations remain optional

An AOT document worker does not need a browser, cloud SDK, or OCR engine unless the application selects one. OfficeIMO keeps these capabilities in focused packages:

- Google Workspace and other network clients bring the authentication and HTTP provider chosen by the application.
- Tesseract and process-based OCR execute an external program even when the OfficeIMO host is native.
- WPF/WebView2 is a desktop rendering integration with its own runtime deployment model.

Keeping those boundaries explicit makes a small native Word, Excel, PowerPoint, Markdown, CSV, Reader, or PDF tool practical without pretending every third-party runtime is compiled into the same binary.

## Test the workflow, not just startup

For production acceptance, generate or ingest a representative document and verify the useful result. Check paragraph text, table values, formulas, slide relationships, conversion diagnostics, or searchable PDF text. Repeat the test for each operating system, architecture, font set, and optional provider you intend to ship.

The [NativeAOT deployment guide](/docs/advanced/aot-trimming/) includes the current package-family guidance, repository verification command, and alternatives such as trimming, ReadyToRun, and single-file deployment.
