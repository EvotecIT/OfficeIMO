# React application acceptance

This fixture runs the unmodified React and React DOM 18.3.1 production bundles
inside the retained OfficeIMO runtime provider. An independently authored report
app loads JSON with `fetch`, mounts with `createRoot`, changes hook state through
accessible-name actions, and captures an immutable OfficeIMO document. The
captured report renders after the worker exits. React is a test input, not an
OfficeIMO runtime dependency.

The source and license provenance are recorded with the
[fixture](../../../../OfficeIMO.Html.Runtime.Tests/Fixtures/React18/README.md).
Run the complete acceptance path with:

```sh
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj -c Release -f net10.0 --filter FullyQualifiedName~React18ReportLoadsAndCapturesFetchedData
```

The same test runs under `net8.0`. It checks fetched data, initial and edited
captures, selected region, adjustment, controlled title input, review-state
transition, semantic headings/table/status, retained live form value, capture
independence, styled screen PNG, searchable tagged print PDF, and searchable
tagged screen-to-page PDF. A separate regression checks passive listener
cancellation and nested-event behavior. The generated artifacts are available
for visual review:

| Intent | Retained artifact | Observed result |
| --- | --- | --- |
| Screen full page | [PNG](react-screen.png) | Blue heading, pale screen background, white report card, one South row, total 21 |
| Browser-style print | [PDF](react-print.pdf), [first-page raster](react-print-pdf.png) | Print media removes screen background and card padding; report remains readable |
| Screen-to-page PDF | [PDF](react-screen-to-page.pdf), [first-page raster](react-screen-to-page-pdf.png) | Screen media colors and card survive on an A4 page |

The PDF previews were rasterized with Poppler 26.07.0 on Windows. Both PDFs
reopen through `PdfReadDocument`, expose tagged content, and retain the selected
report text. The screen PNG is decoded through `OfficePngReader`; its background
pixel matches the stylesheet's `#f3f6fa`.

| Platform | Runtime | Full runtime suite | React acceptance |
| --- | --- | --- | --- |
| Windows x64 | .NET 8 | 353/353 | Pass |
| Windows x64 | .NET 10 | 353/353 | Pass |

This qualifies the selected React 18 application workflow, not arbitrary React
apps or websites. The fixture uses classic UMD scripts and no hydration,
React Router, shadow DOM, canvas, service worker, cookie-backed request,
credentialed fetch, or hostile-script isolation. Accessible-name targeting and
tagged PDF checks cover a bounded semantic path; they are not an assistive
technology reading test. Browser-reference interaction and pixel comparisons
for a production-built React app remain open on the roadmap.
