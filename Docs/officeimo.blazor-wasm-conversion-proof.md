# Browser-local conversion: performance and limits

Measured: 2026-08-16

Use the [OfficeIMO browser document workspace](https://officeimo.com/convert/) to convert supported documents and run focused PDF workflows without uploading them to a conversion service. The static `net10.0` WebAssembly application runs the same OfficeIMO byte and stream APIs that are available to .NET applications.

## Available routes

The browser supports eleven conversion routes: DOCX, XLSX, PPTX, and HTML to PDF; Markdown to HTML or DOCX; HTML to Markdown; and PDF to DOCX, XLSX, PPTX, or HTML. The [conversion map](https://officeimo.com/docs/capabilities/conversions/) identifies every browser route and the focused NuGet package for .NET-only routes.

The PDF workspace adds twelve bounded operations through `OfficeIMO.Pdf`: inspect, compare, merge, split, extract, delete, reorder, rotate, lossless optimize, protect, unlock, and verified literal-text redaction. Each successful operation returns an artifact and a machine-readable report.

Files remain in the current browser tab. OfficeIMO does not send them to a server. A support bundle excludes the source document and converted PDF unless you explicitly choose to include that content.

## Browser limits

- File uploads are limited to 25 MiB and 5,000 package parts.
- Multi-file PDF tools accept no more than ten PDFs or 75 MiB combined; visual comparison is capped at 25 pages.
- Text input is limited to 500,000 characters.
- The XLSX browser-safe preview can limit processing to 250 rows and omit sheet layout and media.
- Conversion reports identify substitutions, approximations, blocked content, and other fidelity warnings.
- Font availability, document complexity, browser memory, and device performance can affect output and conversion time.

Use the focused .NET package when a document exceeds the browser limits, requires a route that is not exposed in WebAssembly, or must run under your own server-side resource policy. OCR/searchable-PDF output, lossy scan compression, and cryptographic signing are intentionally outside the browser workspace.

## Measured baseline

The following Windows Release measurements use native WebAssembly linking and Chromium through the pinned Playwright CLI 0.1.17. They are reproducible regression measurements, not performance guarantees for every device or document.

| Measurement | Observed |
| --- | ---: |
| Published app | 80,158,821 bytes |
| Startup to interactive converter | 1,101 ms |
| Maximum observed browser heap | 151,187,486 bytes |

| Conversion | Time | Observed browser heap | Peak retained PDF buffers | Result |
| --- | ---: | ---: | ---: | ---: |
| DOCX to PDF | 506 ms | 53,489,577 bytes | 12,694 bytes | 200,643 bytes |
| XLSX to PDF | 468 ms | 102,785,222 bytes | 3,805 bytes | 440,298 bytes |
| PPTX to PDF | 490 ms | 151,187,486 bytes | 30 bytes | 769 bytes |

The browser-heap value is Chromium's highest sampled JavaScript heap before, during, or immediately after a conversion. Peak retained PDF buffers measure page-content and object buffers retained by the converter; they are not whole-process memory.

## Regression ceilings

The automated browser gate rejects:

- a published application above 84 MiB;
- startup above 25 seconds;
- observed browser heap above 512 MiB;
- any measured DOCX, XLSX, or PPTX conversion above 15 seconds;
- retained PDF buffers above 64 MiB;
- an empty result, missing memory samples, or browser console errors.

These ceilings protect the checked samples from regressions. They do not imply that every OfficeIMO package runs in WebAssembly or that arbitrary document sizes fit the browser profile.

## Reproduce the measurements

From the repository root:

```powershell
pwsh Website/build.ps1 -CI -SkipBuildTool -Only build,dotnet-publish,overlay

pwsh Website/scripts/Test-ConverterPerformance.ps1 `
  -SiteRoot Website/_site `
  -ReportPath Website/_reports/browser-converter-performance.json
```

Record the SDK, browser, operating system, and hardware when comparing results across machines.
