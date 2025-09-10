# OfficeIMO — Open XML utilities for .NET (Word, Excel, PowerPoint, Visio)

[![CI](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml/badge.svg?branch=master)](https://github.com/EvotecIT/OfficeIMO/actions/workflows/dotnet-tests.yml)
[![codecov](https://codecov.io/gh/EvotecIT/OfficeIMO/branch/master/graph/badge.svg)](https://codecov.io/gh/EvotecIT/OfficeIMO)
[![license](https://img.shields.io/github/license/EvotecIT/OfficeIMO.svg)](LICENSE)

If you would like to contact me you can do so via Twitter or LinkedIn.

[![twitter](https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social)](https://twitter.com/PrzemyslawKlys)
[![blog](https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg)](https://evotec.xyz/hub)
[![linked](https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn)](https://www.linkedin.com/in/pklys)
[![discord](https://img.shields.io/discord/508328927853281280?style=flat-square&label=discord%20chat)](https://evo.yt/discord)

OfficeIMO is a family of lightweight, cross‑platform .NET libraries that make working with Office file formats easier using the Open XML SDK — no Office/COM required.

- Word: create and edit .docx documents with a friendly API
- Excel: fast read/write helpers, tables, styles, ranges, fluent composers
- PowerPoint: build .pptx slides programmatically
- Visio: basic .vsdx diagrams

Each project ships as its own NuGet package under the MIT license.


## Project READMEs

- Word → `OfficeIMO.Word/README.md`
- Excel → `OfficeIMO.Excel/README.md`
- PowerPoint → `OfficeIMO.PowerPoint/README.md`
- Visio → `OfficeIMO.Visio/README.md`
- Converters:
  - `OfficeIMO.Word.Html` — HTML ↔ Word
  - `OfficeIMO.Word.Markdown` — Markdown ↔ Word
  - `OfficeIMO.Word.Pdf` — PDF export for Word

## Targets

- Word, PowerPoint, Visio: netstandard2.0, net472, net8.0 (Linux/macOS: net8.0); select projects also net9.0
- Excel: netstandard2.0, net472, net48, net8.0/net9.0 (cross‑platform)

## Build & Coverage

- CI (Windows/Linux/macOS): single workflow badge above
- Coverage: Codecov dashboard linked above

## Licenses

All OfficeIMO packages are MIT‑licensed. See individual project READMEs for third‑party dependency licenses (Open XML SDK, ImageSharp, AngleSharp, Markdig, QuestPDF, SkiaSharp, etc.).

## Dependencies at a glance

Below are product‑centric graphs. Arrows point from a package to what it depends on.

### Word

```mermaid
flowchart TD
  WCore[OfficeIMO.Word]
  WHtml[OfficeIMO.Word.Html]
  WMd[OfficeIMO.Word.Markdown]
  WPdf[OfficeIMO.Word.Pdf]

  OXml[DocumentFormat.OpenXml]
  Angle[AngleSharp]
  AngleCss[AngleSharp.Css]
  Markdig[Markdig]
  Quest[QuestPDF]
  Skia[SkiaSharp]

  %% OfficeIMO package relationships
  WHtml --> WCore
  WMd --> WCore
  %% Word.Markdown currently references Word.Html in this repo
  WMd --> WHtml
  WPdf --> WCore

  %% External dependencies
  WCore --> OXml
  WHtml --> Angle
  WHtml --> AngleCss
  WMd   --> Markdig
  WPdf  --> Quest
  WPdf  --> Skia
```

### Excel

```mermaid
flowchart TD
  Xl[OfficeIMO.Excel]
  OXml[DocumentFormat.OpenXml]
  ImgSharp[SixLabors.ImageSharp]
  Fonts[SixLabors.Fonts]

  Xl --> OXml
  Xl --> ImgSharp
  Xl --> Fonts
```

### PowerPoint

```mermaid
flowchart TD
  Ppt[OfficeIMO.PowerPoint]
  OXml[DocumentFormat.OpenXml]

  Ppt --> OXml
```

### Visio

```mermaid
flowchart TD
  Vsdx[OfficeIMO.Visio]
  ImgSharp[SixLabors.ImageSharp]
  Pkg[System.IO.Packaging]

  Vsdx --> ImgSharp
  Vsdx --> Pkg
```

### When do I need what?

- Only editing/creating Word (.docx): add `OfficeIMO.Word`.
- Word → PDF export: add `OfficeIMO.Word` + `OfficeIMO.Word.Pdf` (pulls QuestPDF + SkiaSharp).
- Word ↔ HTML: add `OfficeIMO.Word` + `OfficeIMO.Word.Html` (pulls AngleSharp + AngleSharp.Css).
- Word ↔ Markdown: add `OfficeIMO.Word` + `OfficeIMO.Word.Markdown` (pulls Markdig; also uses `OfficeIMO.Word.Html`).
- Excel read/write, tables, styles: add `OfficeIMO.Excel` (pulls ImageSharp + Fonts for sizing and header images).
- PowerPoint slides: add `OfficeIMO.PowerPoint`.
- Visio drawings: add `OfficeIMO.Visio` (uses ImageSharp and System.IO.Packaging).

## Dependency versions (high‑level)

- DocumentFormat.OpenXml: 3.3.x (constraints: [3.3.0, 4.0.0))
- SixLabors.ImageSharp: 2.1.x; SixLabors.Fonts: 1.0.x (Excel)
- AngleSharp: 1.3.x; AngleSharp.Css: 1.0.0‑beta.154 (Word.Html)
- Markdig: 0.41.x (Word.Markdown)
- QuestPDF: 2025.7.x; SkiaSharp: 3.119.x (Word.Pdf)

We keep package ranges conservative to avoid breaking changes; see each project’s csproj for exact ranges.

## Licenses

- OfficeIMO.Word, OfficeIMO.Excel, OfficeIMO.PowerPoint, OfficeIMO.Word.Html, OfficeIMO.Word.Markdown, OfficeIMO.Word.Pdf: MIT
- OfficeIMO.Visio: License TBD (not MIT yet)

Third‑party dependency licenses: see their upstream repos (Open XML SDK, SixLabors, AngleSharp, Markdig, QuestPDF, SkiaSharp).
## Support This Project

If you find this project helpful, please consider supporting its development.
Your sponsorship will help the maintainers dedicate more time to maintenance and new feature development for everyone.

It takes a lot of time and effort to create and maintain this project.
By becoming a sponsor, you can help ensure that it stays free and accessible to everyone who needs it.

To become a sponsor, you can choose from the following options:

- [Become a sponsor via GitHub Sponsors :heart:](https://github.com/sponsors/PrzemyslawKlys)
- [Become a sponsor via PayPal :heart:](https://paypal.me/PrzemyslawKlys)

Your sponsorship is completely optional and not required for using this project.
We want this project to remain open-source and available for anyone to use for free,
regardless of whether they choose to sponsor it or not.

If you work for a company that uses our .NET libraries or PowerShell modules, please consider sponsoring.
Thank you for considering support!

## Please share with the community

Please consider sharing a post about OfficeIMO and the value it provides. It really does help!

[![Share on reddit](https://img.shields.io/badge/share%20on-reddit-red?logo=reddit)](https://reddit.com/submit?url=https://github.com/EvotecIT/OfficeIMO&title=OfficeIMO)
[![Share on hacker news](https://img.shields.io/badge/share%20on-hacker%20news-orange?logo=ycombinator)](https://news.ycombinator.com/submitlink?u=https://github.com/EvotecIT/OfficeIMO)
[![Share on twitter](https://img.shields.io/badge/share%20on-twitter-03A9F4?logo=twitter)](https://twitter.com/share?url=https://github.com/EvotecIT/OfficeIMO&t=OfficeIMO)
[![Share on facebook](https://img.shields.io/badge/share%20on-facebook-1976D2?logo=facebook)](https://www.facebook.com/sharer/sharer.php?u=https://github.com/EvotecIT/OfficeIMO)
[![Share on linkedin](https://img.shields.io/badge/share%20on-linkedin-3949AB?logo=linkedin)](https://www.linkedin.com/shareArticle?url=https://github.com/EvotecIT/OfficeIMO&title=OfficeIMO)

## Features
See individual project READMEs for detailed capability lists and code samples.
