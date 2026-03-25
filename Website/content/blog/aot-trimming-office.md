---
title: "AOT and Trimming: Current State in OfficeIMO"
description: "A practical look at which OfficeIMO packages are lower-risk fits for NativeAOT and trimming, and where you should still test carefully."
date: 2025-11-01
tags: [aot, trimming, performance]
categories: [Deep Dive]
author: "Przemyslaw Klys"
---

.NET 8 and newer made NativeAOT much more accessible, which makes trimming behavior and startup cost matter more for library authors. But OfficeIMO is a package family, not one uniform runtime story, so the right answer depends on which package you are using.

## The Short Version

- **Lower-risk starting points for AOT-sensitive workloads:** `OfficeIMO.Markdown` and `OfficeIMO.CSV`
- **Test carefully with your own scenarios:** `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, and `OfficeIMO.Reader`
- **Treat separately:** `OfficeIMO.Word.Pdf`, because it adds PDF/layout dependencies and host-font concerns

## What the Repo Proves Today

From the project files in this repository:

- `OfficeIMO.Markdown`, `OfficeIMO.Word`, and `OfficeIMO.Excel` explicitly enable `EnableTrimmingPolyfills`.
- `OfficeIMO.Markdown` has no package dependencies.
- `OfficeIMO.CSV` is also a lightweight package with no runtime-heavy dependency graph.
- `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, and `OfficeIMO.Reader` target modern TFMs but still sit on top of Open XML-oriented code paths.
- `OfficeIMO.Word.Pdf` brings in QuestPDF and SkiaSharp.

That means the repo supports a **strong trimming/AOT story for Markdown and CSV**, but it does **not** prove that every OfficeIMO package is uniformly NativeAOT-safe across all code paths.

## Why Markdown and CSV Are the Strongest Candidates

These two packages are the simplest fit for aggressive deployment modes because they avoid the broader Open XML document stack and keep their runtime surface small.

- `OfficeIMO.Markdown` has zero external dependencies and a typed in-memory model.
- `OfficeIMO.CSV` is similarly focused and lightweight for read/write/validation workflows.

If you are building small utilities, serverless helpers, or trimmed CLI tools, these are the safest packages to start with.

## Where You Should Test More Carefully

### Open XML document packages

`OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, and higher-level packages like `OfficeIMO.Reader` ultimately rely on richer document-processing stacks. Those are absolutely usable in modern deployments, but you should validate your own code paths with trimming or `PublishAot` instead of assuming blanket support.

Typical areas to test:

- Read/modify scenarios instead of write-only generation.
- Reflection warnings emitted by upstream dependencies.
- Package combinations such as Reader, converters, or HTML/Markdown bridges.

### PDF conversion

`OfficeIMO.Word.Pdf` is cross-platform, but it is not as lightweight as the Markdown or CSV packages. It uses QuestPDF and SkiaSharp, so host fonts and platform packaging matter, especially in containers.

## A Reasonable Publishing Baseline

If you want to test trimming or AOT yourself, start with a minimal app and verify the exact features you call:

```xml
<PropertyGroup>
  <PublishTrimmed>true</PublishTrimmed>
  <TrimMode>link</TrimMode>
</PropertyGroup>
```

For NativeAOT experiments:

```xml
<PropertyGroup>
  <PublishAot>true</PublishAot>
</PropertyGroup>
```

```bash
dotnet publish -c Release -r linux-x64
```

## Practical Guidance

1. Prefer `OfficeIMO.Markdown` and `OfficeIMO.CSV` first when AOT or trimming is a hard requirement.
2. Treat Word, Excel, PowerPoint, Reader, and converters as scenario-driven validation work.
3. Run publish-time checks early in the project instead of waiting until deployment.
4. Test on the same OS, architecture, and runtime you intend to ship.

## Conclusion

OfficeIMO is friendly to modern deployment patterns, but the current state is nuanced. The lightweight packages already fit trimming-oriented workflows well, while the richer Open XML and PDF packages still deserve targeted validation in your own environment. That is a much more useful rule of thumb than pretending the whole family has the same AOT profile.
