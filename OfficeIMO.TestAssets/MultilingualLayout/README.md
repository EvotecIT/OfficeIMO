# Multilingual PDF layout corpus

These fixtures measure native text reconstruction, OCR reconstruction, and searchable-PDF readback against labels authored independently of OfficeIMO. Pango/Cairo produces the PDFs and raster pages; the fixture generator also records the intended paragraph order, table rows, caption, rotation, and SHA-256 hashes in `manifest.json`.

The eight cases cover Polish/German/English and Hebrew/Arabic/English documents at 0, 90, 180, and 270 degrees. Each page contains a heading, two text columns, mixed-direction version text, a four-row table, and a figure with a caption. This is controlled regression evidence from one producer family, not an accuracy estimate for arbitrary documents or a complete script-support claim.

## Run

Native reconstruction needs the .NET SDK and the checked-in fixtures:

```sh
dotnet run --project Build/PdfQualityCorpus -c Release -f net8.0 -- layout OfficeIMO.TestAssets/MultilingualLayout artifacts/multilingual-layout native
```

To include OCR and searchable readback, omit `native`. Install Tesseract with `eng`, `pol`, `deu`, `heb`, and `ara` trained data. The runner uses 300 DPI, page segmentation mode 3, zero minimum confidence, explicit labelled quarter-turn correction, and `ReconstructLayout = true`. It does not use the expected text to correct provider output.

The output directory contains per-mode text, searchable PDFs, provider text, and `layout-quality.json`. The report separates:

- NFC/whitespace-normalized character and word error rates.
- Exact labelled segments and correctly ordered pairs of present segments.
- Exact table rows and caption classification.
- Provider-token multiset precision and recall, which ignore ordering and retain recognition mismatches.

Missing segments cannot establish a correct reading-order pair. Exact-table scoring requires a single table with the expected row count and compares cells at their labelled row and column positions. Extra tables or rows, duplicated rows, and rows collected across several tables cannot receive a perfect score. Provider errors, including text recognized inside the figure, remain in the results. Full-page scans lack separate figure objects, so caption text can be retained without caption classification. Native figure-caption geometry is tested separately.

Run the scoring and runner contracts with `dotnet run --project Build/PdfQualityCorpus -c Release -f net8.0 -- verify-runner-contracts`.

### Reproduce the historical baseline

The [recorded comparison](../../Docs/quality/multilingual-layout/2026-09-08-linux-net8.json) uses product revisions `316af053340445f7d9933b8560c7f729ff98d8af` and `92d170ad46d187f311c02a896289d826e359fd27` with this same fixture manifest. Build the historical product in a separate checkout and run the committed baseline host from the current checkout:

```sh
baseline_root="$(mktemp -d)/baseline"
git worktree add --detach "$baseline_root" 316af053340445f7d9933b8560c7f729ff98d8af
dotnet build "$baseline_root/OfficeIMO.Pdf.Ocr/OfficeIMO.Pdf.Ocr.csproj" -c Release -f net8.0
dotnet build "$baseline_root/OfficeIMO.Ocr.Tesseract/OfficeIMO.Ocr.Tesseract.csproj" -c Release -f net8.0
dotnet build Build/PdfLayoutBaseline -c Release -p:PdfLayoutBaselineRoot="$baseline_root"
dotnet Build/PdfLayoutBaseline/bin/Release/net8.0/OfficeIMO.PdfLayoutBaseline.dll layout OfficeIMO.TestAssets/MultilingualLayout artifacts/multilingual-baseline
```

Stop if any build or measurement fails. The baseline host links the same scoring source and uses the older product's default OCR layout behavior; its compile symbol omits the `ReconstructLayout` option, which did not exist in that revision. It has no product fallback logic. Both runs use the same provider settings and independently labelled input. Retain the report and source assembly hashes before removing the clean detached worktree with `git worktree remove "$baseline_root"`.

For the measured current product, build a second detached checkout at `92d170ad46d187f311c02a896289d826e359fd27` with the same two product build commands. Rebuild this host with `-p:PdfLayoutBaselineRoot="<current-checkout>" -p:PdfLayoutReconstruct=true`, then run its DLL with a different output directory. This keeps the scoring code identical while enabling the newer option. Later revisions can use the normal corpus command, but their results describe their own product code rather than the dated measurement.

## Regenerate

Run `python3 OfficeIMO.TestAssets/MultilingualLayout/generate.py` on Linux with `python3-cairo`, `python3-gi-cairo`, `gir1.2-pango-1.0`, and DejaVu Sans fonts. The recorded producer is Cairo 1.18.0 with Pango 1.52.1. Font or producer changes may change geometry and hashes; review the generated pages and labels together.

Native PDFs use Cairo's default PDF version. Scan PDFs explicitly use PDF 1.4 and classic cross-reference tables, within the searchable-stamping container contract. The runner verifies original fixture hashes and does not rewrite input containers before measurement.

The original text, layout, and generator are covered by the repository license. Embedded DejaVu font subsets carry the [included font license](FONT-LICENSE.txt). Machine-readable measurements are retained with the [PDF quality evidence](../../Docs/officeimo.real-world-corpus-evidence.md).
