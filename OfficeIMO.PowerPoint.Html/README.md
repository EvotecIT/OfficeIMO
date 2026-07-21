# OfficeIMO.PowerPoint.Html

First-party HTML adapter for OfficeIMO.PowerPoint. It exports semantic slide HTML and positioned review HTML using the shared OfficeIMO.Html profile contracts and the public PowerPoint slide model.

## Semantic round trips

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;

using PowerPointPresentation presentation = PowerPointPresentation.Load("briefing.pptx");
string html = presentation.ToHtml();

HtmlToPowerPointResult result = html.ToPowerPointPresentationResult();
using PowerPointPresentation imported = result.GetArtifactOrThrow();
using FileStream output = File.Create("briefing-roundtrip.pptx");
imported.Save(output);
```

Semantic output carries a versioned OfficeIMO envelope and keeps slide order and visibility, unified drawing order across text boxes, tables, pictures, and charts, shape geometry and transforms, presenter notes, table merge spans, embedded pictures, and supported chart data. Generic HTML `rowspan` and `colspan` values become native PowerPoint table merges.

`ToPowerPointPresentation()` is the convenience API. It throws `HtmlConversionException` when no semantic `section.officeimo-slide` envelope exists. Use `ToPowerPointPresentationResult()` to inspect diagnostics and loss classification, and `ToHtmlResult()` for export evidence.

To turn ordinary HTML sections into slides, select the shared generic path:

```csharp
HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
    .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
        Mode = HtmlImportMode.Auto
    });
```

`Semantic` remains the strict round-trip default. `Auto` uses a supported semantic envelope when present and otherwise groups ordinary headings, text, lists, tables, and embedded images into slides; `Generic` always uses that projection. `HtmlToPowerPointOptions.Limits` bounds slides, shapes, tables, cells, images, chart data, metadata, and geometry before native allocations. `MaxTableCells` remains as a forwarding compatibility property.

Path, stream, and async save/import methods use UTF-8 without a byte-order mark. Stream overloads leave caller-owned streams open.

## Positioned review

Set `Profile = OfficeHtmlConversionProfile.PowerPointVisualReview` for a positioned visual representation. Visual-review HTML is intended for inspection, while semantic slide HTML is the importable contract.

## Targets

`netstandard2.0`, `net8.0`, and `net10.0`; `net472` is included when building on Windows.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.PowerPoint`, `OfficeIMO.Html`, and `OfficeIMO.Drawing` own the slide model, HTML source, mapping, visual review, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
