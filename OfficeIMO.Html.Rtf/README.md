# OfficeIMO.Html.Rtf

`OfficeIMO.Html.Rtf` provides the optional semantic bridge between `HtmlConversionDocument` and `RtfDocument`. Plain HTML and plain RTF applications remain independent.

```powershell
dotnet add package OfficeIMO.Html.Rtf
```

The public APIs remain in the familiar `OfficeIMO.Html` namespace:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Rtf;

HtmlConversionDocument html = HtmlConversionDocument.Parse(
    "<p>Hello <strong>RTF</strong></p>");
RtfDocument rtf = html.ToRtfDocument();

RtfToHtmlResult roundTrip = rtf.ToHtmlResult(
    RtfToHtmlOptions.CreateWebSafeProfile());
roundTrip.RtfReport.RequireNoLoss();
```

For a complete responsive and print-aware review document, select the named print profile and a shared theme:

```csharp
string reviewHtml = rtf.ToHtml(
    RtfToHtmlOptions.CreatePrintReviewProfile(OfficeVisualThemeKind.WordLike));
```

`CreateWebSafeProfile()` is the bounded semantic publishing profile. `CreateRoundTripProfile()` emits a complete HTML document with trusted private metadata and embedded payloads for editable HTML/RTF workflows. The document shell retains resource tables and document-level settings; set `FragmentOnly = true` when the consumer requires a fragment. `CreatePrintReviewProfile()` emits a complete static document with the shared OfficeIMO stylesheet; it never enables script execution or remote browser behavior. `RtfHtmlExportProfile` prevents selecting another adapter's profile, while `SharedProfile` exposes the generic engine mapping. `DocumentOutput` composes full-document versus fragment output, title, language, theme, default styles, and newlines. Conversion results retain per-construct preserved, simplified, omitted, and rejected diagnostics.

The bridge preserves supported structure and reports approximation or loss. Native RTF editing and exact unchanged-source preservation remain in `OfficeIMO.Rtf`.

Bounded, single-surface positioned and floating HTML regions map to editable page-anchored RTF paragraph frames with native size, offsets, wrap controls, and solid backgrounds. Repeated or fragmented paged regions stay in semantic flow; background image layers, shadows, and stacking metadata without an RTF frame equivalent are diagnosed. Set `HtmlToRtfOptions.ImportEditableLayoutRegions = false` to retain semantic flow only.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Html`, and `OfficeIMO.Rtf`.

## ARIA tables and stylesheet diagnostics

Structurally supported ARIA tables become native editable RTF tables, including spans bounded to their row groups. Unsupported structures remain in text flow with an approximation diagnostic. Source-node/depth limits are checked before synthetic native elements are introduced. Active embedded stylesheets and applicable external stylesheet links that the RTF importer does not apply are reported as loss; inline styling remains available.

Ordinary HTML tables with unspecified column widths fit their default grid to the page text area or containing table cell. Authored preferred widths and private RTF round-trip boundaries are preserved.

Unstyled images larger than the effective document, section, or table-cell text area are reduced
proportionally, using embedded image DPI (96 DPI when absent). The default text area
is 6 inches wide by 9 inches high. Explicit image
dimensions are preserved. The result reports `HtmlRtfImageFittedToPage` or
`HtmlRtfImageFittedToCell` when fitting changes an image. Imported HTML table rows stay together; explicit RTF round-trip
metadata preserves rows that allow splitting, including identifiable older exports.
Legacy fragments with no RTF metadata or document wrapper follow ordinary HTML row defaults.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Html.Rtf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
