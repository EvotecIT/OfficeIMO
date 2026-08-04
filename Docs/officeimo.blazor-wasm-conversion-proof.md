# Browser-local conversion: performance and limits

Measured: 2026-08-04

Use the [OfficeIMO browser converter](https://officeimo.com/convert/) to convert supported documents without uploading them to a conversion service. The static `net10.0` WebAssembly application runs the same OfficeIMO byte and stream APIs that are available to .NET applications.

## Available routes

The browser currently supports DOCX, XLSX, and PPTX to PDF, plus selected HTML and Markdown routes. The [conversion map](https://officeimo.com/docs/capabilities/conversions/) identifies every browser route and the focused NuGet package for .NET-only routes.

Files remain in the current browser tab. OfficeIMO does not send them to a server. A support bundle excludes the source document and converted PDF unless you explicitly choose to include that content.

## Browser limits

- File uploads are limited to 25 MiB and 5,000 package parts.
- Text input is limited to 500,000 characters.
- The XLSX browser-safe preview can limit processing to 250 rows and omit sheet layout and media.
- Conversion reports identify substitutions, approximations, blocked content, and other fidelity warnings.
- Font availability, document complexity, browser memory, and device performance can affect output and conversion time.

Use the focused .NET package when a document exceeds the browser limits, requires a route that is not exposed in WebAssembly, or must run under your own server-side resource policy.

## Measured baseline

The following Windows Release measurements use native WebAssembly linking and Chromium through the pinned Playwright CLI 0.1.17. They are reproducible regression measurements, not performance guarantees for every device or document.

| Measurement | Observed |
| --- | ---: |
| Published app | 80,063,848 bytes |
| Startup to interactive converter | 1,206 ms |
| Maximum observed browser heap | 150,725,160 bytes |

| Conversion | Time | Observed browser heap | Peak retained PDF buffers | Result |
| --- | ---: | ---: | ---: | ---: |
| DOCX to PDF | 468 ms | 53,694,154 bytes | 12,694 bytes | 200,643 bytes |
| XLSX to PDF | 463 ms | 102,924,132 bytes | 3,805 bytes | 440,298 bytes |
| PPTX to PDF | 492 ms | 150,725,160 bytes | 30 bytes | 769 bytes |

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
