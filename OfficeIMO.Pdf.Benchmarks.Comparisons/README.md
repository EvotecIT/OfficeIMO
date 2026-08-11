# OfficeIMO PDF library comparisons

This opt-in project measures complete, validated PDF workflows. It is deliberately outside `OfficeIMO.sln`: QuestPDF, PeachPDF, PDFsharp/MigraDoc, PdfPig, and iText are benchmark tools, not OfficeIMO runtime dependencies.

## Workload matrix

| Scale | Real-world shape | Pages | Rows per page | Narrative paragraphs per page |
| --- | --- | ---: | ---: | ---: |
| Easy | Invoice or short status report | 1 | 8 | 1 |
| Medium | Monthly operational report | 20 | 12 | 2 |
| High | Annual audit archive | 100 | 12 | 3 |

Each page contains a heading, narrative text, and a four-column account/status table. The data is deterministic and created before measurement when it is an input. Every generated result must have a PDF header, the exact expected page count, and every required narrative and table row before its lane is accepted.

The benchmark families intentionally answer different questions:

- `PdfGenerationBenchmarks`: OfficeIMO, QuestPDF, MigraDoc/PDFsharp, and iText generate the same structured report from the same logical model. The measured operation includes document construction, layout, font embedding, compression, and in-memory serialization.
- `PdfHtmlBenchmarks`: OfficeIMO.Html.Pdf and PeachPDF parse and render the exact same HTML string. The measured operation includes HTML/CSS parsing, paged layout, and in-memory PDF serialization.
- `PdfReadBenchmarks`: OfficeIMO.Pdf, PdfPig, and iText open identical bytes, enumerate every page, and extract the complete text payload. The corpus is repeated for OfficeIMO-, QuestPDF-, PeachPDF-, MigraDoc-, and iText-produced PDFs to avoid a single-producer result.
- `PdfSplitBenchmarks`: OfficeIMO, iText, and PDFsharp split the same OfficeIMO- and iText-produced documents into single pages and fixed-size bundles.
- `PdfMergeBenchmarks`: all three engines merge the same ordered source set and preserve the exact page-marker sequence.
- `PdfPageSelectionBenchmarks`: all three engines extract the same reversed, non-contiguous page selection and preserve that order.
- `PdfCorpusReadBenchmarks`: OfficeIMO, PdfPig, and iText extract four prepared large documents: a 500-page OfficeIMO PDF, the 492-page NIST SP 800-53 Rev. 5, an 85-page Type3-font fixture, and a 258-page 12.9 MB PDF/A fixture.

Every timed producer/read combination must first pass page-count and complete deterministic-content validation. Split, merge, and selection outputs are opened with PdfPig and checked for output count, page count, order, and every required narrative and table row. A failed or mutation-blocked workflow is reported as compatibility evidence and is not published as a performance result.

## Interoperability corpus

`Corpus/pdf-corpus.json` combines repository fixtures, generated documents, and pinned public files. Downloaded files are opt-in, written under an ignored output directory, and accepted only when their SHA-256 matches the manifest. The corpus currently covers:

- native Microsoft Word, Excel, and PowerPoint exports;
- a 25-page OfficeIMO.Word source with tables, chart, SmartArt, image, links, lists, headers, and footers, plus its OfficeIMO PDF and structured conversion diagnostics;
- a 500-page OfficeIMO.Pdf document;
- NIST SP 800-53 Rev. 5 and IRS Form W-9;
- a W3C standards document;
- CC0 veraPDF Type0/ToUnicode, Type3-font, and large PDF/A fixtures.

The read oracle is PdfPig, with iText fallback when PdfPig cannot read a file. OfficeIMO is compared by duplicate-aware, per-page token recall. Corpus reading opts into `PdfReadOptions.IncludeArtifactText` so headers, footers, and chart decorations are included in the same visual-text contract as the comparison readers.

The OfficeIMO.Word source deliberately contains SmartArt. OfficeIMO.Word.Pdf currently reports `NativeBodySmartArtUnsupported`, so the generated OfficeIMO PDF is not labeled as containing SmartArt; its conversion JSON preserves that product gap. The Windows Word COM lane opens the same DOCX, adds a genuine Word-native SmartArt object through `Shapes.AddSmartArt`, and exports it through Microsoft Word. Validation requires the resulting 26-page PDF and independently checks that OfficeIMO recovers the interoperability-page heading plus every SmartArt node label. New, unrecognized conversion-loss diagnostics fail corpus preparation.

After a read pass, the corpus selects the last, middle, and first pages, splits that result into single-page documents, merges them again, and independently verifies order and token retention. OfficeIMO intentionally blocks unsafe full rewrites of documents whose forms, signatures, tagged content, active content, outlines, xref streams, or object streams cannot yet be preserved. The JSON report records these as `Blocked` with machine-readable mutation blocker codes, separately from failed output validation. Those blockers identify manipulation work to implement; the runner does not bypass them.

## Measurements

BenchmarkDotNet reports mean/median timing, statistical error, rank, GC collections, and managed allocated bytes per operation. `Allocated` is managed allocation volume, not peak memory or total resident memory. This distinction matters for QuestPDF because its Skia work uses native memory.

The existing `OfficeIMO.Pdf.Benchmarks` budget runner remains the OfficeIMO-only source for sampled peak managed heap and retained writer-buffer evidence. A future process-isolated lane is required before comparing total peak resident memory across managed and native engines; do not relabel BenchmarkDotNet managed allocations as total memory.

Setup and validation are outside measured operations. Output byte length is observed but is not treated as a correctness substitute: compression and font subsetting legitimately produce different file sizes.

## Run

Run one quick correctness/performance smoke through the shared PowerForge evidence path:

```powershell
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfgenerate -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtml -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfread -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfsplit -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfmerge -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfselect -RunMode quick -Framework net10.0
```

Use `-RunMode full` for publication-quality BenchmarkDotNet statistics. `-Publish` is valid only with a full run and updates the shared benchmark evidence catalog. Raw BenchmarkDotNet artifacts stay under the ignored output root.

For a local short engineering run without catalog updates:

```powershell
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfGenerationBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfHtmlBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfReadBenchmarks*" --job Short
```

Prepare and validate the real-document corpus before running its BenchmarkDotNet lane:

```powershell
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- corpus --download --output Ignore/Benchmarks/PdfComparisons/corpus

pwsh Build/Run-LibraryComparisonBenchmarks.ps1 `
    -Workload pdfcorpusread `
    -RunMode quick `
    -Framework net10.0 `
    -PdfCorpusRoot Ignore/Benchmarks/PdfComparisons/corpus/files
```

The corpus workload is excluded from `-Workload all` because it requires opt-in downloads. Use `--only id-1,id-2` with the corpus command to rerun selected entries.

On Windows with Microsoft Word installed, this command generates the rich DOCX, injects a Word-native SmartArt fixture, exports it through Word COM, and validates OfficeIMO readback of the rendered node labels. COM remains a fixture producer and is not an OfficeIMO dependency:

```powershell
pwsh Build/Run-PdfWordComCorpus.ps1 -Framework net10.0
```

## Benchmark-only libraries

Versions are pinned in the benchmark project so evidence is reproducible. PDFsharp/MigraDoc is MIT, PeachPDF is BSD-3-Clause, and PdfPig is Apache-2.0. QuestPDF uses its Community license for this open-source benchmark project. iText is AGPL/commercial and remains isolated here for benchmark use; it is not linked by or distributed with an OfficeIMO runtime package.

The benchmark uses the maintained cross-platform PDFsharp 6.x package. PdfSharpCore 1.3.67 is not included: restoring it currently reports NuGet vulnerability advisories through its ImageSharp 1.0.4 dependency, while PDFsharp 6.x already covers the equivalent cross-platform split, merge, selection, and MigraDoc generation workflows.
