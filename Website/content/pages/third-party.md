---
title: "Third-Party Dependencies"
description: "Runtime dependency and upstream license notes for the public OfficeIMO package family."
layout: page
---

OfficeIMO packages are published under the [MIT License](https://github.com/EvotecIT/OfficeIMO/blob/master/LICENSE), but several packages intentionally build on upstream open-source components. This page exists to make those relationships easier to review during OSS approval, procurement, and redistribution checks.

## Scope

- Covers shipped runtime dependencies for the public OfficeIMO package families documented on this site.
- Focuses on user-visible dependencies rather than test-only packages, benchmark harnesses, sample apps, or framework/reference-only helpers.
- Exact package ranges live in the linked `.csproj` files and NuGet metadata.
- This page is informational only and is not legal advice.

## Package Family Overview

| OfficeIMO package or family | Upstream components used in the repo today | Why they are there |
|---|---|---|
| `OfficeIMO.Word` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`, `SixLabors.ImageSharp` `2.1.11` | OOXML document model, packaging, colors, and image handling |
| `OfficeIMO.Excel` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`, `SixLabors.ImageSharp` `2.1.11`, `SixLabors.Fonts` `1.0.1` | Workbook model, image support, and font measurement/layout work |
| `OfficeIMO.PowerPoint` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)` | Presentation OOXML model and packaging |
| `OfficeIMO.Word.Html` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`, `AngleSharp` `1.3.0`, `AngleSharp.Css` `1.0.0-beta.157` | HTML and CSS parsing for Word conversion workflows |
| `OfficeIMO.Markdown.Html` | `AngleSharp` `1.3.0` | HTML parsing for Markdown conversion and bridge scenarios |
| `OfficeIMO.Word.Pdf` | `QuestPDF` `2026.2.0`, `SkiaSharp` `3.119.2` | PDF document layout and graphics rendering |
| `OfficeIMO.Visio` | `SixLabors.ImageSharp` `2.1.11`, `System.IO.Packaging` `10.0.3` | Image handling plus OPC packaging support for `.vsdx` files |
| `OfficeIMO.Markdown` | No third-party runtime package references | Core package is intentionally dependency-light |
| `OfficeIMO.CSV` | No third-party runtime package references | Core package is intentionally dependency-light |
| `OfficeIMO.Reader` | Composes first-party OfficeIMO packages | Its effective upstream surface follows the format packages it wraps |

Additional Microsoft compatibility helpers may appear on older target frameworks, but the table above focuses on the upstream packages that most teams will care about during license review.

## Upstream License Notes

| Upstream project | License or model | OfficeIMO packages that use it | What to know |
|---|---|---|---|
| [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/3.5.1) | MIT | Word, Excel, PowerPoint, Word.Html | This is Microsoft's official Open XML SDK and the main OOXML building block in the Office document packages. |
| [SixLabors.ImageSharp](https://www.nuget.org/packages/SixLabors.ImageSharp/2.1.11) | Six Labors Split License 1.0 | Word, Excel, Visio | Review the upstream terms carefully for commercial or enterprise usage. This is one of the main dependencies that deserves explicit OSS review. |
| [SixLabors.Fonts 1.0.1](https://www.nuget.org/packages/SixLabors.Fonts/1.0.1) | Apache License 2.0 | Excel | OfficeIMO currently pins the older `1.0.1` line. If this dependency is upgraded in future, re-check its license because newer Six Labors package lines may use different terms. |
| [AngleSharp](https://www.nuget.org/packages/AngleSharp/1.3.0) | MIT | Markdown.Html, Word.Html | Used for HTML parsing and DOM work. |
| [AngleSharp.Css](https://www.nuget.org/packages/AngleSharp.Css/1.0.0-beta.157) | MIT | Word.Html | Adds CSS parsing on top of AngleSharp for HTML conversion flows. |
| [QuestPDF](https://www.questpdf.com/license/guide.html) | Hybrid model: community option plus paid tiers | Word.Pdf | Review the QuestPDF license guide for your usage context. It is not a plain one-size-fits-all MIT dependency. |
| [SkiaSharp](https://www.nuget.org/packages/SkiaSharp/3.119.2) | MIT | Word.Pdf | Graphics and drawing backend used by the PDF conversion layer. |
| [System.IO.Packaging](https://www.nuget.org/packages/System.IO.Packaging/10.0.3) | MIT | Visio | Microsoft packaging primitives for OPC-style containers. |

## What We Recommend Teams Check

1. Review the exact `PackageReference` list for the OfficeIMO packages you ship, not just the repo root license.
2. Treat `SixLabors.ImageSharp` and `QuestPDF` as the first dependencies to review explicitly during commercial approval.
3. Re-check upstream terms whenever a dependency version changes, especially around PDF and imaging stacks.
4. Keep a copy of the upstream notices or license URLs in your own release/compliance workflow if your organization requires that.

## How We Address This In The Website

- We keep the first-party OfficeIMO package license visible across the repo and site.
- We document the public dependency surface here instead of hiding it in project files.
- We keep the page scoped to real shipped package dependencies so it stays reviewable and current.

If you need the exact current references, start with the repository project files such as [`OfficeIMO.Word.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word/OfficeIMO.Word.csproj), [`OfficeIMO.Excel.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel/OfficeIMO.Excel.csproj), [`OfficeIMO.Word.Pdf.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Pdf/OfficeIMO.Word.Pdf.csproj), and [`OfficeIMO.Visio.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Visio/OfficeIMO.Visio.csproj).
