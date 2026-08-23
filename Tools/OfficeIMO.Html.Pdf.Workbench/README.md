# OfficeIMO HTML to PDF Workbench

A focused loopback-only operator surface for exercising the two OfficeIMO HTML-to-PDF lanes:

- **Managed** uses `OfficeIMO.Html.Pdf` for dependency-free parsing, layout, pagination, tagging, forms, diagnostics, and PDF writing.
- **Chromium** uses the pooled `HtmlTinkerX` renderer through `OfficeIMO.Html.Pdf.Browser` for browser layout, readiness, lifecycle, security policy, and capture diagnostics.

The public OfficeIMO converter remains a static WebAssembly application. This tool is intentionally server-hosted because a browser tab cannot start or pool Chromium locally.

## Run from current sources

The workbench consumes the new fail-closed `HtmlBrowserNetworkPolicy.Offline` contract. Until that HtmlTinkerX version is published, point the build at a current local HtmlTinkerX checkout instead of adding compatibility code:

```powershell
$root = if ($env:EVOTEC_GITHUB_ROOT) { $env:EVOTEC_GITHUB_ROOT } else { 'C:\Support\GitHub' }
$htmlTinkerX = Join-Path $root 'HtmlTinkerX\Sources\HtmlTinkerX\HtmlTinkerX.csproj'
dotnet run --project .\Tools\OfficeIMO.Html.Pdf.Workbench\OfficeIMO.Html.Pdf.Workbench.csproj --property:HtmlTinkerXProjectPath=$htmlTinkerX
```

Open `http://127.0.0.1:5105`. The host rejects non-loopback `Workbench:Url` values.

If Chromium is not installed for the HtmlTinkerX Playwright build, build once and run the generated Playwright installer for Chromium.

## Evidence contract

Each successful render exposes a PDF and a companion JSON document containing:

- schema version, engine, and renderer version;
- source and artifact SHA-256 fingerprints;
- elapsed time, byte count, and page count;
- the exact settings snapshot;
- typed diagnostics and a loss flag;
- Chromium version, reuse/retry state, blocked-request count, and stage timings for browser renders.

Artifacts are addressed by random 192-bit tokens, retained in memory only, bounded to 32 entries, and expire after 30 minutes. HTML and CSS are bounded to 2 MiB and generated PDFs to 32 MiB. The live preview applies a restrictive CSP and a sandbox; Chromium capture uses HtmlTinkerX's offline network policy.
