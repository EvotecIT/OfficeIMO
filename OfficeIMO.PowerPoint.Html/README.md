# OfficeIMO.PowerPoint.Html

First-party HTML adapter for OfficeIMO.PowerPoint. It exports semantic slide HTML and positioned review HTML using the shared OfficeIMO.Html profile contracts and the public PowerPoint slide model.

## Semantic round trips

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint.Html;

using PowerPointPresentation presentation = PowerPointPresentation.Load("briefing.pptx");
string html = presentation.ToHtml();

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
HtmlToPowerPointResult result = source.ToPowerPointPresentationResult();
using PowerPointPresentation imported = result.RequireValue();
using FileStream output = File.Create("briefing-roundtrip.pptx");
imported.Save(output);
```

Semantic output carries a versioned OfficeIMO envelope and keeps slide order and visibility, unified drawing order across text boxes, tables, pictures, charts, SmartArt, and media, shape geometry and transforms, presenter notes, table merge spans, embedded pictures, supported chart data, master/layout inventory, poster frames, and supported picture adjustments. SmartArt and advanced effects use static snapshots or diagnosed fallbacks; media is never executed. Generic HTML `rowspan` and `colspan` values become native PowerPoint table merges.

`ToPowerPointPresentation()` is the convenience API. It throws `HtmlConversionException` when no semantic `section.officeimo-slide` envelope exists. Use `ToPowerPointPresentationResult()` to inspect diagnostics and loss classification, and `ToHtmlResult()` for export evidence. Master/layout projection, SmartArt fallback, inert media, advanced effects, unavailable pictures or charts, and visual-renderer fallbacks are represented in the immutable operation report so `RequireNoLoss()` cannot accept a simplified review silently.

To turn ordinary HTML sections into slides, select the shared generic path:

```csharp
HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
    .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
        Mode = HtmlImportMode.Auto
    });
```

`Semantic` remains the strict round-trip default. `Auto` uses a supported semantic envelope when present and otherwise groups ordinary headings, text, lists, tables, and embedded images into slides; `Generic` always uses that projection. `HtmlToPowerPointOptions.Limits` bounds slides, shapes, tables, cells, images, chart data, metadata, and geometry before native allocations. `MaxTableCells` remains as a forwarding compatibility property.

On the ordinary HTML path, bounded positioned, floating, flex, and grid regions become editable slide text boxes and DrawingML pictures at rendered geometry. Solid fills, supported background/image layers, picture opacity, and the first box-shadow layer use native PowerPoint constructs. Additional shadows and unsupported effects receive stable diagnostics. Set `ImportEditableLayoutRegions = false` to retain semantic flow only.

`SaveAsHtml` and `SaveAsHtmlAsync` write UTF-8 without a byte-order mark to paths or caller-owned streams. For import I/O, use `HtmlConversionDocument.Load(...)` or `LoadAsync(...)`, then call `ToPowerPointPresentation()` or `ToPowerPointPresentationResult()` on the prepared document. Stream overloads leave caller-owned streams open.

## Positioned review

Use `PowerPointHtmlSaveOptions.CreateVisualReviewProfile()` or set `ExportProfile = PowerPointHtmlExportProfile.VisualReview` for a positioned visual representation. `SharedProfile` exposes the corresponding generic engine lane. `DocumentOutput` controls full-document versus fragment output, title, language, theme, default styles, and newlines. Visual-review HTML is intended for inspection, while semantic slide HTML is the importable contract.

## Targets

`netstandard2.0`, `net8.0`, and `net10.0`; `net472` is included when building on Windows.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.PowerPoint`, `OfficeIMO.Html`, and `OfficeIMO.Core` own the slide model, HTML source, mapping, visual review, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
