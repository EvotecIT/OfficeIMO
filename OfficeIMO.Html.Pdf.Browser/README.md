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

## Provisioning and deployment

The managed [HTML renderer](../OfficeIMO.Html/PDF.md) runs without Chromium.
This optional package uses HtmlTinkerX's Playwright driver and Chromium runtime.
HtmlTinkerX installs missing components on first use; provisioning needs download
access and writable cache locations. The page's network policy does not govern
installer downloads. Provision during deployment when the running application
must work without download access:

```csharp
await HtmlBrowser.EnsureInstalledAsync(HtmlBrowserEngine.Chromium);
```

On Linux the browser also needs its operating-system libraries. HtmlTinkerX owns
installation and browser-version selection; keep dependency installation in the
image or deployment setup. Its `HTMLTINKERX_PLAYWRIGHT_WITH_DEPS=1` provisioning
option installs the selected browser's Linux dependencies and requires suitable
build-time privileges. It does not make a bare runtime image browser-ready
without that preparation.

Use `PLAYWRIGHT_DRIVER_SEARCH_PATH` and `PLAYWRIGHT_BROWSERS_PATH` when provisioning
and runtime identities must share an installation. Ensure the runtime identity
can read the driver and browser and write to its temporary directories. Set up
the complete installation before disabling download access.

## Package and live capture checks

The repository's browser package smoke defaults to API checks. The opt-in capture
also executes JavaScript in offline HTML, opens the resulting PDF through
OfficeIMO, and checks the resulting text and tagged structure:

```powershell
./Build/Test-HtmlPdfBrowserPackages.ps1 -Capture -PublishPath /path/to/browser-smoke-publish
```

`PublishPath` must be a new directory. The script qualifies local packages in an
isolated package cache and publishes the consumer for .NET 10. To test a Linux
container, use that published directory as the Docker build context:

```sh
docker build -f Build/PackageSmoke/OfficeIMO.Html.Pdf.Browser/Dockerfile \
  -t officeimo-browser-smoke /path/to/browser-smoke-publish
docker run --rm --network none officeimo-browser-smoke
docker image rm officeimo-browser-smoke
```

The image provisions Chromium and its Linux dependencies during build, then runs
the capture as the image's `app` user with runtime network access disabled. A
successful API-only smoke does not establish browser startup or container
acceptance; record the actual capture and container results separately.
