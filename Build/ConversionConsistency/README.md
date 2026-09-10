# Conversion consistency checks

This tool compares the native PDF and image exports of authored documents. It checks every expected page, required text, physical dimensions, diagnostics, selected text baselines and color regions. PNG, SVG, JPEG, TIFF and WebP are included. Poppler renders saved PDFs independently; Chromium renders SVG before its pixels are compared with the direct PNG export.

The source corpus covers HTML, PDF, DOCX, XLSX, PPTX, ODT, ODS, ODP, OneNote, email and EPUB documents, and Visio diagrams. A separate report suite checks a two-page action register with 40 distinct action IDs, typed numeric cells, an inline chart, and subscript/superscript text.

## Run locally

Use .NET 8 or later, PowerShell 7, Poppler's `pdftoppm` on `PATH`, and the Chromium build installed by the tool's Playwright script. Run from the repository root:

```powershell
dotnet build Build/ConversionConsistency/OfficeIMO.ConversionConsistency.Tool.csproj -c Release -f net8.0
./Build/ConversionConsistency/bin/Release/net8.0/playwright.ps1 install chromium

$tool = 'Build/ConversionConsistency/bin/Release/net8.0/OfficeIMO.ConversionConsistency.Tool.dll'
$evidence = '.artifacts/conversion-consistency'
dotnet $tool prepare --output "$evidence/fixtures"
dotnet $tool run --suite "$evidence/fixtures/suite.json" --output "$evidence/native"
dotnet $tool run --suite Build/ConversionConsistency/suite.json --output "$evidence/reports"
./Build/ConversionConsistency/Test-BundleContract.ps1 -BundlePath "$evidence/native" -OutputPath "$evidence/negative"
```

On Linux, Playwright may also need `install --with-deps chromium`. The CI workflow installs that dependency and Poppler explicitly.

Output directories must be empty when preparing or exporting. Choose a new directory for another run, or remove only the earlier output that you own. `--case native-pptx` selects a single source case. `--repository` selects the repository root; `--pdftoppm` selects the independent rasterizer executable for verification. Unknown, duplicate, and inapplicable options are rejected.

`export` creates a bundle without checking it. `verify --output <bundle-directory>` verifies an existing bundle without rerunning conversion. Exit code `0` means the selected checks passed, `1` means verification failed, and `2` means an invocation or setup error prevented completion. Missing or modified page artifacts produce a failed case in `consistency-result.json`.

## Read the evidence

Each case contains its PDF, all image pages, the independent PDF/SVG rasterizations, and pixel-difference images. `bundle.json` records source and output hashes, the source commit and tracked diff hash, untracked source-file hashes, font hash, rendering profile, native route, expected pages, and declared limitations. Task output under `.artifacts` is excluded from source provenance. `consistency-result.json` records comparisons and failures, plus the external renderer versions.

Expected labels and geometry come from the authored source contract. Adding a case requires specifying the pages and content that should survive conversion. This prevents an empty image or a shared pagination mistake from passing simply because two export routes agree.

The default pixel allowance is 2% with a mean channel error of 3. The dense report fixtures allow 3%; tight worksheet crops allow 4% and a mean error of 4.5. Those cases also assert text positions separately, because small-font rasterization differs between the managed renderer and Chromium. Changing a tolerance requires inspecting the output and retaining independent content and geometry checks.

## Coverage limits

| Source | PDF coverage |
| --- | --- |
| HTML, PDF, DOCX, ODT, PPTX, ODP | Searchable text, page geometry, and pixels compared with native images |
| XLSX, ODS | Searchable text and paper-page dimensions; worksheet images use tighter content bounds |
| OneNote | Raster PDF pages compared with native images; no searchable PDF text layer |
| Visio | Searchable semantic PDF projection on paper pages; diagram images use the native diagram canvas |
| Email | External browser PDF checks body text and dimensions; native images include email presentation chrome |
| EPUB | External browser PDF from the authored chapter HTML, compared with native EPUB images |

These limits are stored per case. Reduced PDF coverage requires an explicit explanation. A passing result means the declared checks passed; it does not establish complete format fidelity or interchangeable Word, Excel, and HTML layouts.

The workflow uploads the evidence for review. The negative checks verify that missing labels, missing or altered image files, duplicate pages, and invalid CLI options cannot produce a successful result.
