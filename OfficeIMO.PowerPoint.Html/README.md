# OfficeIMO.PowerPoint.Html

First-party HTML adapter for OfficeIMO.PowerPoint. It exports semantic slide HTML and positioned review HTML using the shared OfficeIMO.Html profile contracts and the public PowerPoint slide model.

## Semantic round trips

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint.Html;

using PowerPointDocument presentation = PowerPointDocument.Load("briefing.pptx");
string html = presentation.ToHtml();

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
using FileStream output = File.Create("briefing-roundtrip.pptx");
OfficeConversionReport report = source.SaveAsPowerPoint(output);
```

Semantic output carries a versioned OfficeIMO envelope and keeps slide order and visibility, unified drawing order across text boxes, tables, pictures, charts, SmartArt, and media, shape geometry and transforms, presenter notes, table merge spans, embedded pictures, supported chart data, master/layout inventory, poster frames, and supported picture adjustments. SmartArt and advanced effects use static snapshots or diagnosed fallbacks; media is never executed. Generic HTML `rowspan` and `colspan` values become native PowerPoint table merges.

`ToPowerPoint(options, out report)` returns the presentation with its `OfficeConversionReport` and throws `HtmlConversionException` when no semantic `section.officeimo-slide` envelope exists; `SaveAsPowerPoint` writes it and returns the report. On export, `ToHtml(options, out report)` and `SaveAsHtml` return the `OfficeConversionReport`; set `FailOnLoss` to make reported loss an error. Master/layout projection, SmartArt fallback, inert media, advanced effects, unavailable pictures or charts, and visual-renderer fallbacks are represented in the immutable operation report so `ThrowIfLoss()` cannot accept a simplified review silently.

To turn ordinary HTML sections into slides, select the shared generic path:

```csharp
using PowerPointDocument presentation = HtmlConversionDocument.Parse(html)
    .ToPowerPoint(new HtmlPowerPointExportOptions {
        Mode = HtmlImportMode.Auto
    }, out OfficeConversionReport report);
```

`Semantic` remains the strict round-trip default. `Auto` uses a supported semantic envelope when present and otherwise groups ordinary headings, text, lists, tables, and embedded images into slides; `Generic` always uses that projection. `HtmlPowerPointExportOptions.Limits` bounds slides, shapes, tables, cells, images, chart data, metadata, and geometry before native allocations. `MaxTableCells` remains as a forwarding compatibility property.

On the ordinary HTML path, bounded positioned, floating, flex, and grid regions become editable slide text boxes and DrawingML pictures at rendered geometry. Solid fills, supported background/image layers, picture opacity, and the first box-shadow layer use native PowerPoint constructs. Additional shadows and unsupported effects receive stable diagnostics. Set `ImportEditableLayoutRegions = false` to retain semantic flow only.

`SaveAsHtml` and `SaveAsHtmlAsync` write UTF-8 without a byte-order mark to paths or caller-owned streams and return the `OfficeConversionReport`. `ToHtmlAsync` returns the HTML with its report as `OfficeConversionResult<string>`; the async forms run the CPU-bound conversion on the thread pool. For import I/O, use `HtmlConversionDocument.Load(...)` or `LoadAsync(...)`, then call `ToPowerPoint` or `SaveAsPowerPoint` on the prepared document. Stream overloads leave caller-owned streams open.

## Generic document flow

Generic HTML import divides long paragraphs into editable text boxes across slides and retains supported inline formatting and hyperlinks. Tables with suitable row geometry continue as native tables with repeated header rows. Native header links use the header text color, including on repeated table fragments, while authored run colors remain explicit. Readers supporting the native hyperlink text-color extension retain these colors; older readers may use their theme hyperlink color. Captions remain separate editable text. Oversized cells and unsupported list pagination produce explicit approximation reports; authored slide geometry remains authoritative.

Ordinary HTML runs use shared computed font styles, including inherited families
and resolved CSS custom properties. Heading titles retain their own text and font.

## Positioned review

Use `PowerPointHtmlExportOptions.CreateVisualReviewProfile()` or set `ExportProfile = PowerPointHtmlExportProfile.VisualReview` for a positioned visual representation. `SharedProfile` exposes the corresponding generic engine lane. `DocumentOutput` controls full-document versus fragment output, title, language, theme, default styles, and newlines. Visual-review HTML is intended for inspection, while semantic slide HTML is the importable contract.

Both profiles start with an `h1` heading showing `Title`, or a profile name when no title is set; set `IncludeHeading = false` to omit it. Each slide starts with an `h2` such as "Slide 1"; set `IncludeSlideHeadings = false` to omit those. Visual review lists each slide's render diagnostics as visible text below the slide; set `IncludeRenderDiagnostics = false` to leave them out of the HTML. They are still reported in the conversion report. The following settings omit the presentation and slide headings, snapshot diagnostic lists, master inventory, and extraction proof; enabled SmartArt or media review sections remain:

```csharp
var options = PowerPointHtmlExportOptions.CreateVisualReviewProfile();
options.IncludeHeading = false;
options.IncludeSlideHeadings = false;
options.IncludeRenderDiagnostics = false;
options.IncludeMasterInventory = false;
options.IncludeExtractionProof = false;
string html = presentation.ToHtml(options, out OfficeConversionReport report);
```

Each visual-review slide embeds an SVG snapshot whose text is measured and drawn with installed platform fonts. Set `Fonts` to supply TrueType faces that take precedence. Supply every face required by the slides for snapshots that do not depend on installed fonts; missing faces can still resolve through platform font fallback. It is null by default, which keeps platform font resolution.

```csharp
using OfficeIMO.Drawing;

var options = PowerPointHtmlExportOptions.CreateVisualReviewProfile();
options.Fonts = new OfficeFontFaceCollection()
    .Add("Calibri", File.ReadAllBytes("Fonts/Calibri.ttf"))
    .Add("Calibri", File.ReadAllBytes("Fonts/Calibri-Bold.ttf"), OfficeFontStyle.Bold);
string html = presentation.ToHtml(options);
```

## Targets

`netstandard2.0`, `net8.0`, and `net10.0`; `net472` is included when building on Windows.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.PowerPoint`, `OfficeIMO.Html`, and `OfficeIMO.Core` own the slide model, HTML source, mapping, visual review, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint.Html` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
