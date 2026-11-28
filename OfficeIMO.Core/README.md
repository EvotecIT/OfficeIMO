# OfficeIMO.Core

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Core)](https://www.nuget.org/packages/OfficeIMO.Core)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Core?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Core)

The shared foundation every OfficeIMO package builds on: lengths, colors and themes, conversion reports, document lifecycle, package safety, tabular row mapping, fonts and text measurement, vector drawing, charts, raster image codecs, SVG, ink and math. It has no third-party dependencies. You rarely install it directly; the format packages bring it in.

## Install

```sh
dotnet add package OfficeIMO.Core
```

## Quick start

Lengths and colors are the values you meet in every format's API:

```csharp
using OfficeIMO;

OfficeLength width = OfficeLength.Parse("2.5cm");
OfficeLength total = width + OfficeLength.Points(6) * 2;
Console.WriteLine(total.ToString("cm"));      // "2.9233cm"

OfficeTheme theme = OfficeTheme.Modern.WithPalette(
    OfficeTheme.Modern.Palette.With(accent1: OfficeColor.Parse("#3458A4")));
OfficeColor accent = OfficeColor.Accent1;     // a theme role, resolved when written
OfficeColor rgb = accent.Resolve(theme);
```

`OfficeLength` stores whole EMUs and makes units explicit. `OfficeColor` holds literal colors and theme roles in one type, so output follows the document's theme.

## What it does

- **Themes, colors and lengths** shared by Word, Excel, PowerPoint, PDF and the other engines, with the built-in `Modern`, `Classic`, `Minimal`, `Boardroom` and `Dark` themes.
- **Conversion reports and diagnostics**: `OfficeConversionReport`, loss kinds and `ThrowIfLoss`, used by every converter.
- **Document lifecycle and package safety**: shared load, create and persistence options, bounded package and embedded payload reads, and atomic saves.
- **Data**: per-column row mapping, `RowsAs<T>`, JSON data and JSON Schema, and sequence and text comparison.
- **Text**: TrueType, OpenType and variable font measurement, managed multilingual shaping, vertical text and hyphenation.
- **Drawing and charts**: vector shapes, gradients, transforms and chart snapshots that the format packages and PDF render.
- **Images**: PNG, JPEG, TIFF, WebP and SVG encoding and decoding, metadata reading and editing, frame selection, placement optimization and scanned-image preparation.
- **Ink and math**: reusable ink strokes and structured equations with MathML and LaTeX conversion.
- **Safety and provenance**: SVG safety checks, concealed-content inspection before AI ingestion, and C2PA and IPTC provenance inspection and removal.
- **Entity recognition**: `OfficeIMO.Entities` validates identifiers and finds and masks identifiers, email addresses and URLs in text. See the [entity recognition guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.entities.md).
- **Security contracts** that Word, PDF and email use for signatures; the cryptography itself is in the optional [OfficeIMO.Security](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Security/README.md).

The [Core guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.core-guide.md) shows each of these with examples, and [managed raster workflows](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.core-raster-workflows.md) covers image pipelines.

## Limits

Core describes conversion routes but doesn't run them; the format packages own document semantics, layout and export APIs. The [content safety matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.content-safety-support-matrix.md) and the [image export capability matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.image-export-capability-matrix.md) list what the shared inspectors and renderers cover.

## Learn more

- [Core guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.core-guide.md) and the [API design](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.api-design.md) rules every package follows.
- [Package catalog](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.package-catalog.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: none. The package embeds reference data, such as Unicode tables, font metrics and codec tables, under their own notices; see the [third-party notices](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Core/THIRD-PARTY-NOTICES.md).
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Inspect | 2 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 0 | 0 | 0 | 0 | 0 |
| Remove | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Core` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
