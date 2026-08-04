# Blazor WebAssembly conversion proof

Date: 2026-08-04

OfficeIMO ships a static `net10.0` Blazor WebAssembly converter for representative DOCX, XLSX, and PPTX to PDF workflows. The app references the local OfficeIMO projects directly, uses byte/stream APIs, embeds a deterministic browser font pack, and performs conversion locally without uploading documents.

## Reproducible gate

The website CI pipeline publishes the native-linked WebAssembly app, verifies its static deployment shape, and runs `Website/scripts/Test-ConverterPerformance.ps1`. The Playwright probe loads the production publish through a local HTTP server, waits for the real Blazor surface, converts the checked sample for each route, and rejects:

- a publish above 84 MiB;
- startup above 25 seconds;
- observed browser heap above 512 MiB;
- a conversion above 15 seconds;
- retained PDF serialization buffers above 64 MiB;
- an empty result or browser console error.

The committed budgets are regression ceilings, not target marketing numbers. SDK, browser, operating system, and hardware must be recorded when comparing measurements.

## Current local baseline

The table below records a Windows Release publish produced with the repository's .NET 10 SDK and native WebAssembly linking. The repository-pinned Playwright CLI 0.1.17 drove its explicit Chromium engine on the same machine.

| Measurement | Observed |
| --- | ---: |
| Published app | 79,984,480 bytes |
| Startup to interactive converter | 1,531 ms |
| Maximum observed browser heap | 150,168,583 bytes |

| Conversion | Time | Observed browser heap | Peak retained PDF buffers | Result |
| --- | ---: | ---: | ---: | ---: |
| DOCX to PDF | 624 ms | 53,922,871 bytes | 12,694 bytes | 200,643 bytes |
| XLSX to PDF | 667 ms | 102,638,381 bytes | 3,805 bytes | 440,298 bytes |
| PPTX to PDF | 453 ms | 150,168,583 bytes | 30 bytes | 769 bytes |

`peakBrowserHeapBytes` is Chromium's highest sampled total JavaScript heap before, during, or immediately after each representative conversion. The gate requires at least two non-zero samples per route and fails closed when Chromium's memory API is unavailable. `peakRetainedBytes` is the converter's high-water retained PDF page-content plus object-buffer evidence; it is not presented as whole-process memory.

## Commands

```powershell
dotnet publish Website/Apps/OfficeIMO.Web.Converter/OfficeIMO.Web.Converter.csproj `
  -c Release `
  -o Website/_temp/converter-publish `
  -p:BaseHref=/apps/officeimo-converter/

pwsh Website/scripts/Test-ConverterPerformance.ps1 `
  -SiteRoot Website/_site `
  -ReportPath Website/_reports/browser-converter-performance.json
```

The pipeline owns the publish overlay into `Website/_site/apps/officeimo-converter`; the performance command expects that exact static-site layout.

## Boundary

This evidence proves the three checked browser routes and samples. It does not imply that every OfficeIMO conversion package is WebAssembly-compatible, that arbitrary document sizes fit the browser profile, or that the measured Windows numbers transfer unchanged to every device. Package-size and input/package-part limits remain enforced by the browser conversion service, while conversion reports expose fidelity warnings and the support bundle records the selected profile and performance evidence.
