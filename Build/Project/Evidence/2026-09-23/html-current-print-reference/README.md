# HTML print-reference baseline

This is the first comparison checkpoint for H10 unfamiliar-page conversion. It
reruns the existing OfficeIMO-authored H4 corpora against PeachPDF 0.9.19 and
Chromium 151.0.7922.34. It does not qualify unfamiliar pages or general browser
equivalence; that requires the independently sourced H10 corpus in the roadmap.

The source was clean at `12556b10ab4214351ee3b6f94a08b8755f859f40`, and the
loaded OfficeIMO renderer assembly identified that same commit. The run used
macOS 27 arm64, .NET 10.0.12, HtmlTinkerX 3.0.1 and Poppler `pdftoppm` 26.05.0.
The H4 advanced held-out manifest SHA-256 was
`496f78d459bfd7836987541925f3d6f4b26c87512cd03f319ca99bc4057f67a8`.
The benchmark project is opt-in and PeachPDF is not an OfficeIMO runtime
dependency.

The advanced held-out acceptance passed 8/8 cases with zero runner failures.
The separate 15-case representative run had zero runner failures but has no
visual acceptance manifest. Both runs recorded direct OfficeIMO-to-Chromium and
PeachPDF-to-Chromium print comparisons, plus OfficeIMO-to-PeachPDF differences.
OfficeIMO was given zero caller-supplied outer margins; PeachPDF and Chromium
used their print defaults. Authored `@page` rules still applied.

| Case | Corpus | OfficeIMO/Chromium MAE | PeachPDF/Chromium MAE | Print pages, OfficeIMO/PeachPDF/Chromium |
| --- | --- | ---: | ---: | --- |
| Legacy portal | Held out | 8.62 | 5.27 | 1/1/1 |
| Named-pages brochure | Held out | 1.91 | 14.53 | 3/3/3 |
| Stacking/clipping board | Held out | 2.21 | 9.94 | 1/1/1 |
| Chart report | Representative | 11.33 | 4.54 | 1/1/1 |
| Product catalog | Representative | 12.39 | 9.18 | 1/1/1 |
| Book extract | Representative | 4.04 | 3.97 | 3/3/4 |

MAE is the unweighted mean of comparable page-pixel errors in the runner's
center-cropped 96-DPI raster comparison. It is a triage signal, not a product
ranking or proof of correct text, accessibility, pagination or print geometry.
The page-count mismatch in the book extract remains open even though that
representative corpus does not have an acceptance gate.

Visual inspection identifies two first candidates for defect analysis: the
legacy portal's table-column and form-control sizing, and the chart report's
page width and SVG plot paint. Do not adjust acceptance tolerances to hide a
validated renderer defect. Reproduce each issue in the owning layout or drawing
component, then rerun the affected case and the complete held-out gate.

To reproduce, use the repository-pinned .NET SDK and put Poppler's `pdftoppm`
and `pdftotext` on `PATH`. Rebuild at the checked-out commit; `--require-clean-source`
rejects a loaded renderer assembly built from another commit.

```bash
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj \
  -c Release -f net10.0 -- html-corpus-evidence --corpus advanced-held-out \
  --output /tmp/officeimo-html-advanced --verify-acceptance --require-clean-source

dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj \
  -c Release -f net10.0 -- html-corpus-evidence --corpus representative \
  --output /tmp/officeimo-html-representative --require-clean-source
```

Choose new output directories for subsequent runs. The original full reports
and page images are task-owned temporary evidence; this compact record retains
the decisive configuration and findings without checking in those large files.
