# OfficeIMO and X-Ray comparison

An opt-in runner that measures OfficeIMO against the X-Ray .NET document packages
(`X-Ray.Content` and `X-Ray.Office`, pinned exact versions from nuget.org) on the same inputs.
It stays outside `OfficeIMO.slnx` and every solution filter, so X-Ray never enters OfficeIMO
restore, build, CI or packages. X-Ray targets .NET 10 only, so the runner does too.

## Workloads

| Lane | OfficeIMO | X-Ray |
| --- | --- | --- |
| `TextExtractionBenchmarks` | `OfficeDocumentReader.ReadDocument` with `ReaderProjection.Text`, then `GetText()` | `Extractor.Extract` with `OutputFormat.Plain` |
| `MarkdownExtractionBenchmarks` | `ReadDocument` with the full projection, `Markdown` and `Tables` | `Extractor.Extract` with `OutputFormat.Markdown`, `Content` and `Tables` |
| `OfficeToPdfBenchmarks` | `WordDocument`/`ExcelDocument`/`PowerPointDocument.Load` (read-only) and `SaveAsPdf` | `OfficeDocument.Open`, `Layout`, `PdfRenderer.Render` |

Extraction starts from the same in-memory bytes on both sides; Office to PDF starts from the same
file path, because X-Ray's renderer opens paths. Every operation consumes the full result.
OfficeIMO's Reader caps table rows (200) and chunk size (8,000 characters) by default, and its XLSX
Markdown drops what does not fit a chunk. X-Ray returns the whole document, so the runner lifts both
caps; set `XRAY_BENCH_READER_DEFAULTS=1` to measure the defaults instead.

Inputs are DOCX, XLSX, PPTX and text-layer PDF files from the repository in small, medium and large
sizes, plus two generated files written on first use: a text-heavy DOCX (240 sections, 30 tables) and
a 10,000-row XLSX with currency, percent and date formats. `cases` lists them.

## Validation

`validate` runs each workload once per library, saves every output under
`<work>/validation`, and scores it against a library-neutral reference:

- DOCX, PPTX and XLSX text comes straight from the package XML (`w:t`, `a:t`, shared strings and raw
  cell values). PDF text, page counts, fonts and page-1 rasters come from Poppler
  (`pdftotext`, `pdfinfo`, `pdffonts`, `pdftoppm`), which must be on `PATH` or in `POPPLER_BIN`.
- Word recall is the share of reference word tokens (NFKC, lower-case, letters and digits) found in
  the output; numeric tokens are reported separately because X-Ray applies cell number formats.
  Precision is the share of output word tokens present in the reference.
- Table cell recall compares reference table cells with the cells each library reports.
- XLSX outputs also count currency (`$1,234.56`), percent (`3.2%`) and ISO date (`2026-01-31`) cells.
- Office to PDF checks page count, embedded fonts, word recall of the `pdftotext` output against the
  source and writes a page-1 PNG for each side to look at.

Do not compare timings for a case whose validation shows a missing or failed output.

## Run

From the repository root, with work output outside the checkout:

```powershell
$env:XRAY_BENCH_WORK = 'D:\Scratch\xray-bench'
dotnet run -c Release --project .\Build\XRayComparison -- cases
dotnet run -c Release --project .\Build\XRayComparison -- validate
dotnet run -c Release --project .\Build\XRayComparison -- --filter '*TextExtraction*' --artifacts $env:XRAY_BENCH_WORK\bdn
dotnet run -c Release --project .\Build\XRayComparison -- --filter '*MarkdownExtraction*' --artifacts $env:XRAY_BENCH_WORK\bdn
dotnet run -c Release --project .\Build\XRayComparison -- --filter '*OfficeToPdf*' --artifacts $env:XRAY_BENCH_WORK\bdn
dotnet run -c Release --project .\Build\XRayComparison -- peaks
```

BenchmarkDotNet uses 3 warm-up and 12 measured iterations and reports the median, min, max,
standard deviation and allocated bytes per operation, with OfficeIMO as the baseline. It runs in
process, because the generated out-of-process project rebuilds the whole OfficeIMO graph and exceeds
BenchmarkDotNet's build timeout. `peaks` runs each large case in a fresh process per library and
reports the peak working set. `loop <lane> <case> oi|xr <count>` repeats one operation with per-run
time and allocation, for attaching a profiler such as
`dotnet-trace collect --profile dotnet-sampled-thread-time -- <exe> loop pdf xlsx-l oi 3`.
Timings are host-dependent: check CPU load before a run and treat a busy machine's numbers as
indicative only.
