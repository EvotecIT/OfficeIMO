# Four application classes through one document workflow

The same OfficeIMO host/page, structured-action, capture and output API passed
four independently authored fixture classes: a vanilla ES-module application,
a production-built React 18 application, Preact 10.29.8, and a static legacy
page with an inline event handler. The caller selected the final state through
accessible-name actions. The runtime session closed before the captured
document produced screen PNG, print PDF and screen-to-page PDF. All eight PDFs
reopened with searchable text and tagged content.

| Case | Final state | OfficeIMO screen | Chromium screen | OfficeIMO print | Chromium print | Screen-to-page PDF |
| --- | --- | --- | --- | --- | --- | --- |
| Vanilla module | `/review`, Quarterly / South, Total: 25 | [PNG](vanilla/officeimo-screen.png) | [PNG](vanilla/chromium-screen.png) | [PDF](vanilla/officeimo-print.pdf), [preview](vanilla/officeimo-print-preview.png) | [PDF](vanilla/chromium-print.pdf), [preview](vanilla/chromium-print-preview.png) | [PDF](vanilla/officeimo-screen-to-page.pdf) |
| React build | `/review`, Quarterly / South, Total: 21 | [PNG](react-build/officeimo-screen.png) | [PNG](react-build/chromium-screen.png) | [PDF](react-build/officeimo-print.pdf), [preview](react-build/officeimo-print-preview.png) | [PDF](react-build/chromium-print.pdf), [preview](react-build/chromium-print-preview.png) | [PDF](react-build/officeimo-screen-to-page.pdf) |
| Preact | Quarterly / South, Total: 18 | [PNG](preact/officeimo-screen.png) | [PNG](preact/chromium-screen.png) | [PDF](preact/officeimo-print.pdf), [preview](preact/officeimo-print-preview.png) | [PDF](preact/chromium-print.pdf), [preview](preact/chromium-print-preview.png) | [PDF](preact/officeimo-screen-to-page.pdf) |
| Legacy | Approved | [PNG](legacy/officeimo-screen.png) | [PNG](legacy/chromium-screen.png) | [PDF](legacy/officeimo-print.pdf), [preview](legacy/officeimo-print-preview.png) | [PDF](legacy/chromium-print.pdf), [preview](legacy/chromium-print-preview.png) | [PDF](legacy/officeimo-screen-to-page.pdf) |

[The browser manifest](chromium-manifest.json) records Chromium 151.0.7922.34
through Playwright 1.62.1 on Linux x64, the 816 by 900 CSS-pixel viewport,
resource hashes and final semantic states. The reproducible harness is in
[BrowserReference](../../../../../OfficeIMO.Html.Runtime.Tests/BrowserReference/README.md).
The OfficeIMO output test is in
[RuntimeApplicationDocumentWorkflowTests.cs](../../../../../OfficeIMO.Html.Runtime.Tests/RuntimeApplicationDocumentWorkflowTests.cs).
OfficeIMO artifacts here were rendered on Windows x64 under .NET 10. The PDFs
were previewed with Poppler 24.02.0. These are cross-engine, cross-platform
references, so font availability is part of the observed difference.

The screen images preserve the selected content and broad layout, with visible
font, form-control and spacing differences. The Preact page has no author CSS;
Chromium uses a serif browser default while OfficeIMO currently uses a sans
default. The legacy table columns and React table widths also differ. OfficeIMO
continuous screen images stop at content height while Chromium's full-page
screenshot is at least viewport height. Print previews retain the same content,
but are not pixel matches. These are named workflow results, not claims of
general framework or website compatibility.

The full runtime suite at `089a7c64df` passed 362/362 on Windows x64 under
.NET 8 and .NET 10 and on Linux x64 and macOS Arm64 under .NET 10. The
evidence-output hook added after that commit passed its four selected Windows
.NET 10 cases. The inputs are explicitly trusted and offline. Arbitrary public
scripts still require the separate OS isolation profile in the roadmap.
