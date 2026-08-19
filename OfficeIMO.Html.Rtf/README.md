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
roundTrip.Report.RequireNoLoss();
```

For a complete responsive and print-aware review document, select the named print profile and a shared theme:

```csharp
string reviewHtml = rtf.ToHtml(
    RtfToHtmlOptions.CreatePrintReviewProfile(OfficeVisualThemeKind.WordLike));
```

`CreateWebSafeProfile()` is the bounded semantic publishing profile. `CreateRoundTripProfile()` enables trusted private metadata and embedded payloads for editable HTML/RTF workflows. `CreatePrintReviewProfile()` emits a complete static document with the shared OfficeIMO stylesheet; it never enables script execution or remote browser behavior. `RtfHtmlExportProfile` prevents selecting another adapter's profile, while `SharedProfile` exposes the generic engine mapping. `DocumentOutput` composes full-document versus fragment output, title, language, theme, default styles, and newlines. Conversion results retain per-construct preserved, simplified, omitted, and rejected diagnostics.

The bridge preserves supported structure and reports approximation or loss. Native RTF editing and exact unchanged-source preservation remain in `OfficeIMO.Rtf`.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Html`, and `OfficeIMO.Rtf`.
