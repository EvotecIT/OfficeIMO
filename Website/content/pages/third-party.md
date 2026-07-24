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
| `OfficeIMO.Drawing.HarfBuzz` | `HarfBuzzSharp` `14.2.1.1` plus matching Windows, Linux, macOS, and WebAssembly native assets | Optional full OpenType GSUB/GPOS shaping through the first-party Drawing provider contract; Drawing and PDF core remain independent of HarfBuzz |
| `OfficeIMO.Word` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`; `Microsoft.Bcl.AsyncInterfaces` `10.0.9` on legacy targets | OOXML document model and compatibility helpers; colors and image metadata use first-party `OfficeIMO.Drawing` |
| `OfficeIMO.Excel` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`; `Microsoft.Bcl.AsyncInterfaces` `10.0.9` and `System.Text.Json` `[10.0.7, 11.0.0)` on legacy targets | Workbook model, first-party image metadata, and compatibility helpers for older target frameworks |
| `OfficeIMO.PowerPoint` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`; `Microsoft.Bcl.AsyncInterfaces` `10.0.9` on legacy targets | Presentation OOXML model, packaging, and compatibility helpers |
| `OfficeIMO.Word.Html` | `DocumentFormat.OpenXml` `[3.5.1, 4.0.0)`; `AngleSharp` `1.5.2` and `AngleSharp.Css` `1.0.0-beta.216` through `OfficeIMO.Html` | OOXML plus shared HTML and CSS parsing for Word conversion workflows |
| `OfficeIMO.Markdown.Html` | `AngleSharp` `1.5.2` and `AngleSharp.Css` `1.0.0-beta.216` through `OfficeIMO.Html` | Shared HTML and CSS parsing for Markdown conversion and bridge scenarios |
| `OfficeIMO.Word.Pdf` | First-party `OfficeIMO.Word` and `OfficeIMO.Pdf` project references | Word-to-PDF conversion through the OfficeIMO PDF engine |
| `OfficeIMO.Excel.Pdf` | First-party `OfficeIMO.Excel` and `OfficeIMO.Pdf` project references | Excel-to-PDF conversion through the OfficeIMO PDF engine |
| `OfficeIMO.Visio` | `System.IO.Packaging` `10.0.8`; `Microsoft.Bcl.AsyncInterfaces` `10.0.9` on `net472` | OPC packaging support for `.vsdx` files and legacy async compatibility; colors and image metadata use first-party `OfficeIMO.Drawing` |
| `OfficeIMO.Markdown` | No third-party runtime package references | Keeps its runtime surface self-contained |
| `OfficeIMO.CSV` | `System.Buffers` `4.5.1` on legacy targets | Compatibility buffer primitives; shared document primitives come from first-party `OfficeIMO.Drawing` |
| `OfficeIMO.Reader.Core`, selective `OfficeIMO.Reader.*` adapters, and `OfficeIMO.Reader.All` | Core uses `System.Text.Json` `[10.0.7, 11.0.0)` only on legacy targets; adapters compose their named first-party format packages | Core stays format-neutral, selective adapters add only their owning engines, and All deliberately composes every local managed adapter |
| `OfficeIMO.Security` | `BouncyCastle.Cryptography` `[2.6.2, 3.0.0)` | One neutral CMS, S/MIME, RFC 3161, and X.509 engine shared directly by Email and PDF without format-specific cryptography packages |

Additional Microsoft compatibility helpers may appear on older target frameworks, but the table above focuses on the upstream packages that most teams will care about during license review.

## Upstream License Notes

| Upstream project | License or model | OfficeIMO packages that use it | What to know |
|---|---|---|---|
| [HarfBuzzSharp](https://www.nuget.org/packages/HarfBuzzSharp/14.2.1.1) | MIT | Drawing.HarfBuzz | Optional .NET bindings and platform-native HarfBuzz assets for full OpenType shaping; not referenced by Drawing or PDF core packages. |
| [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/3.5.1) | MIT | Word, Excel, PowerPoint, Word.Html | This is Microsoft's official Open XML SDK and the main OOXML building block in the Office document packages. |
| [AngleSharp](https://www.nuget.org/packages/AngleSharp/1.5.2) | MIT | Html, Markdown.Html, Word.Html, Reader.Html | Used for shared HTML parsing and DOM work. |
| [AngleSharp.Css](https://www.nuget.org/packages/AngleSharp.Css/1.0.0-beta.216) | MIT | Html, Markdown.Html, Word.Html, Reader.Html | Adds CSS parsing on top of the shared HTML engine. |
| [Microsoft.Bcl.AsyncInterfaces](https://www.nuget.org/packages/Microsoft.Bcl.AsyncInterfaces/10.0.9) | MIT | Word, Excel, PowerPoint, and Visio legacy targets | Provides async interface compatibility for older target frameworks. |
| [System.Text.Json](https://www.nuget.org/packages/System.Text.Json/10.0.7) | MIT | Excel and Reader.Core legacy targets | Provides JSON support where it is not supplied by the target framework. |
| [System.IO.Packaging](https://www.nuget.org/packages/System.IO.Packaging/10.0.8) | MIT | Visio | Microsoft packaging primitives for OPC-style containers. |
| [System.Buffers](https://www.nuget.org/packages/System.Buffers/4.5.1) | MIT | CSV legacy targets | Provides buffer primitives for older target frameworks. |
| [BouncyCastle.Cryptography](https://www.nuget.org/packages/BouncyCastle.Cryptography/2.6.2) | MIT | Security; transitively Email and PDF security workflows | Implements the neutral CMS/X.509 primitives behind the OfficeIMO-owned security contracts without exposing vendor types. |

## What We Recommend Teams Check

1. Review the exact `PackageReference` list for the OfficeIMO packages you ship, not just the repo root license.
2. Re-check upstream terms whenever a dependency version changes, especially around parsing, packaging, or compatibility helper libraries.
3. Keep a copy of the upstream notices or license URLs in your own release/compliance workflow if your organization requires that.

## How We Address This In The Website

- We keep the first-party OfficeIMO package license visible across the repo and site.
- We document the public dependency surface here instead of hiding it in project files.
- We keep the page scoped to real shipped package dependencies so it stays reviewable and current.

If you need the exact current references, start with the repository project files such as [`OfficeIMO.Word.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word/OfficeIMO.Word.csproj), [`OfficeIMO.Excel.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel/OfficeIMO.Excel.csproj), [`OfficeIMO.Word.Pdf.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Pdf/OfficeIMO.Word.Pdf.csproj), [`OfficeIMO.Excel.Pdf.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Pdf/OfficeIMO.Excel.Pdf.csproj), and [`OfficeIMO.Visio.csproj`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Visio/OfficeIMO.Visio.csproj).
