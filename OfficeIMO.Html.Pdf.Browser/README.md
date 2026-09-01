# OfficeIMO.Html.Pdf.Browser

`OfficeIMO.Html.Pdf.Browser` is the optional Chromium-backed route for pages that need browser layout, JavaScript, or live website navigation. `OfficeIMO.Html.Pdf` remains the managed, dependency-light default for static HTML.

## Install

```powershell
dotnet add package OfficeIMO.Html.Pdf.Browser
```

HtmlTinkerX owns browser lifecycle, navigation, waits, credentials, cookies, origin-scoped headers and storage, viewport/device emulation, network policy, capture limits, cancellation, and diagnostics. This package opens the resulting bytes through `OfficeIMO.Pdf`, so extraction, inspection, preflight, and mutation use the same APIs as any other PDF source.

```csharp
using HtmlTinkerX;
using OfficeIMO.Html.Pdf.Browser;
using OfficeIMO.Pdf;

await using var renderer = new HtmlBrowserPdfRenderer(
    new HtmlBrowserPdfRendererOptions(
        new HtmlBrowserPdfDeviceEmulation(
            deviceScaleFactor: 1),
        viewportWidth: 1440,
        viewportHeight: 900));
var request = new HtmlBrowserPdfRequest(
    HtmlBrowserPdfSource.FromUrl("https://example.com"),
    pdfOptions: new HtmlBrowserPdfOptions(tagged: false));

PdfDocumentConversionResult result =
    await renderer.CapturePdfDocumentResultAsync(request);

PdfDocumentPreflight preflight = result.Value.Preflight();
PdfMutationPlan stamping = result.Value.PlanMutation(
    PdfMutationOperation.ModifyPageContent);
string text = result.Value.Read().Text;
```

Use `tagged: false` when later full-rewrite operations are important. Use `tagged: true` when Chromium-generated accessibility structure is required. In both cases, treat `PdfDocument.Preflight()` and `PdfDocument.PlanMutation(...)` as authoritative: a valid, readable PDF can still have operation-specific rewrite blockers.

The bridge never falls back from managed rendering to Chromium automatically. Browser capture executes page code and may access network resources, so callers must opt in and configure HtmlTinkerX security policy explicitly.
