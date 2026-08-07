# OfficeIMO.Html.Rtf

`OfficeIMO.Html.Rtf` provides the optional semantic bridge between `HtmlConversionDocument` and `RtfDocument`. Plain HTML and plain RTF applications remain independent.

```powershell
dotnet add package OfficeIMO.Html.Rtf
```

The public APIs remain in the familiar `OfficeIMO.Html` namespace:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Rtf;

HtmlConversionDocument html = HtmlConversionDocument.Parse(
    "<p>Hello <strong>RTF</strong></p>");
RtfDocument rtf = html.ToRtfDocument();

RtfToHtmlResult roundTrip = rtf.ToHtmlResult(
    RtfToHtmlOptions.CreateWebSafeProfile());
roundTrip.Report.RequireNoLoss();
```

The bridge preserves supported structure and reports approximation or loss. Native RTF editing and exact unchanged-source preservation remain in `OfficeIMO.Rtf`.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Html`, and `OfficeIMO.Rtf`.
