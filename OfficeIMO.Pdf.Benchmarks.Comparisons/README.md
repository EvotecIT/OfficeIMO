# OfficeIMO PDF library comparisons

This opt-in project measures complete, validated PDF workflows. It is deliberately outside `OfficeIMO.sln`: QuestPDF, PeachPDF, PDFsharp/MigraDoc, PdfPig, iText, pdfHTML, HtmlTinkerX, and Chromium are benchmark tools, not OfficeIMO runtime dependencies.

## Workload matrix

| Scale | Real-world shape | Pages | Rows per page | Narrative paragraphs per page |
| --- | --- | ---: | ---: | ---: |
| Easy | Invoice or short status report | 1 | 8 | 1 |
| Medium | Monthly operational report | 20 | 12 | 2 |
| High | Annual audit archive | 100 | 12 | 3 |

Each page contains a heading, narrative text, and a four-column account/status table. The data is deterministic and created before measurement when it is an input. Every generated result must have a PDF header, the exact expected page count, and every required narrative and table row before its lane is accepted. Table validation treats numerically equivalent cells such as `1074.50` and `1074.5` as the same value while continuing to require exact text in non-numeric cells.

The benchmark families intentionally answer different questions:

- `PdfGenerationBenchmarks`: OfficeIMO, QuestPDF, MigraDoc/PDFsharp, and iText generate the same structured report from the same logical model. The measured operation includes document construction, layout, font embedding, compression, and in-memory serialization.
- `PdfHtmlBenchmarks`: OfficeIMO.Html.Pdf, PeachPDF, iText pdfHTML, and Chromium through HtmlTinkerX parse and render the exact same HTML string. Every engine emits tagged PDF bytes and must preserve the exact page count, narrative, and table content before its measurements are accepted. The managed engines include HTML/CSS parsing, paged layout, and in-memory serialization. Chromium reuses one HtmlTinkerX-owned browser session per benchmark case; each measured operation still replaces and reparses the complete page before printing, so warmed browser throughput is not mislabeled as process startup.
- `PdfHtmlPayloadBenchmarks`: OfficeIMO.Html.Pdf and PeachPDF render exact 21 KiB plain-text, table-heavy, and multilingual HTML payloads as tagged PDFs. The multilingual lane makes the same bundled Carlito font the primary CSS family for both engines and requires every measured Latin, Greek, and Cyrillic sample plus the embedded font in the resulting artifact. It therefore runs portably without host-font dependencies or role mismatches. The quick runner uses BenchmarkDotNet's process-isolated `Dry` job for cold-start evidence; full runs measure warmed throughput. Cleanup reopens each result, checks page count, first/last content, the unique terminal marker, all multilingual samples, and reports HTML bytes, PDF bytes, pages, and extracted-text length.
- `PdfFormatConversionBenchmarks` and `PdfExtendedFormatConversionBenchmarks`: all fourteen advertised OfficeIMO source routes parse deterministic source bytes and produce a PDF in one measured operation. DOCX, XLSX, PPTX, HTML, Markdown, RTF, AsciiDoc, LaTeX, MHTML, OneNote, ODT, ODS, ODP, and Visio outputs are reopened independently; every lane requires all four semantic fields for each of 120 records and reports source bytes, PDF bytes, pages, and extracted-text length. This is an OfficeIMO route-health benchmark, not a third-party comparison: adapters with materially different format contracts are not forced into artificial parity. The shared runner can execute this local health lane, but never writes it to the library-comparison evidence catalog, including when `all` or `-Publish` selects it.
- `PdfReadBenchmarks`: OfficeIMO.Pdf, PdfPig, and iText open identical bytes, enumerate every page, and extract the complete text payload. The corpus is repeated for OfficeIMO-, QuestPDF-, PeachPDF-, MigraDoc-, and iText-produced PDFs to avoid a single-producer result.
- `PdfStructuredReadFastBenchmarks` and `PdfStructuredReadCompleteBenchmarks`: separate OfficeIMO-only route-health suites for the one canonical `PdfDocument.Load(...).Read(...)` contract. Both routes include source snapshotting, parsing, glyph recovery, word/line grouping, recursive XY-cut reading order, semantic classification, logical projection, and table extraction; `Structured` additionally applies document-wide evidence. Keeping the profiles in separate BenchmarkDotNet classes prevents shared ranks or baselines between unequal work. The runner excludes both suites from comparison publication, both gate their page/table invariants, and `Structured` additionally gates the labelled document-wide semantic contract it promises. Its labelled Easy fixture uses two pages, rather than the general one-page Easy scenario, because running-header/footer recovery requires repeated document evidence. The deterministic benchmark and scorecard raise only their document-wide work ceiling as a function of fixture page count so the 100-page case measures the complete route; production read defaults remain unchanged.
- `PdfSplitBenchmarks`: OfficeIMO, iText, and PDFsharp split the same OfficeIMO- and iText-produced documents into single pages and fixed-size bundles, then reopen every output with the producing engine and verify its page count inside the timed operation.
- `PdfMergeBenchmarks`: all three engines merge the same ordered source set, reopen the serialized output with the producing engine inside the timed operation, and preserve the exact page-marker sequence.
- `PdfPageSelectionBenchmarks`: all three engines extract the same reversed, non-contiguous page selection, reopen the serialized output with the producing engine inside the timed operation, and preserve that order.
- `PdfReverseConversionBenchmarks`: five isolated PDF producers feed the same DOCX, HTML, XLSX-table, editable-PPTX, ODT, ODS, editable-ODP, and PNG reconstruction routes. Global setup validates the complete generated source, reopens every target artifact, checks page scope and editable structures, and requires deterministic narrative or table-row retention before timings are accepted; timed methods include PDF parsing, projection, and target serialization. Run this lane through `Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfreverse`.
- `PdfCorpusReadBenchmarks`: OfficeIMO, PdfPig, and iText extract four prepared large documents: a 500-page OfficeIMO PDF, the 492-page NIST SP 800-53 Rev. 5, an 85-page Type3-font fixture, and a 258-page 12.9 MB PDF/A fixture.

Every timed producer/read combination must first pass page-count and complete deterministic-content validation. Split, merge, and selection include one producer-native post-save reopen and expected-page-count check in every measured lane. Their setup additionally opens outputs with PdfPig and checks output count, page count, order, and every required narrative and table row. A failed or mutation-blocked workflow is reported as compatibility evidence and is not published as a performance result.

## Interoperability corpus

`Corpus/pdf-corpus.json` combines repository fixtures, generated documents, and pinned public files. Downloaded files are opt-in, written under an ignored output directory, and accepted only when their SHA-256 matches the manifest. The corpus currently covers:

- native Microsoft Word, Excel, and PowerPoint exports, with an opt-in Windows COM lane that regenerates all three producers;
- a 25-page OfficeIMO.Word source with tables, chart, SmartArt, image, links, lists, headers, and footers, plus its OfficeIMO PDF and structured conversion diagnostics;
- a 500-page OfficeIMO.Pdf document;
- NIST SP 800-53 Rev. 5 and IRS Form W-9;
- a W3C standards document;
- CC0 veraPDF Type0/ToUnicode, Type3-font, and large PDF/A fixtures.

The read oracle is PdfPig, with iText fallback when PdfPig cannot read a file. OfficeIMO's canonical structured read is compared by duplicate-aware, per-page token recall. Labelled `expectedText` values independently gate OfficeIMO text instead of inheriting an oracle defect. A `pageExpectations` entry can require exact table, image, image-region, and figure counts plus a minimum vector-primitive count. Each declared table has its own exact row and column shape and exact required cells, so text from unrelated tables cannot satisfy the contract and extra false-positive tables fail validation. Feature labels remain reporting dimensions rather than implicit correctness claims. Corpus reading opts into `PdfLoadOptions.IncludeArtifactText` so headers, footers, and chart decorations are included in the same visual-text contract as the comparison readers. The schema-3 corpus report records the observed semantic counts and also aggregates read failures, elapsed time, and managed allocations by tier and feature. These single-pass corpus observations diagnose document classes; BenchmarkDotNet remains the statistical performance source.

The OfficeIMO.Word source deliberately contains SmartArt. OfficeIMO.Word.Pdf currently reports `NativeBodySmartArtUnsupported`, so the generated OfficeIMO PDF is not labeled as containing SmartArt; its conversion JSON preserves that product gap. The Windows Office COM lane opens the same DOCX, adds a genuine Word-native SmartArt object through `Shapes.AddSmartArt`, and exports it through Microsoft Word. It also creates an Excel workbook with a table, chart, tightly positioned identifiers, and multilingual cells, plus a PowerPoint deck with a table, multilingual text, and a process diagram. Word, Excel, and PowerPoint export their own PDFs; OfficeIMO is only the parser under test. Validation checks producer-specific page counts, independently labelled multilingual text, and producer-specific semantic structures before comparing duplicate-aware token recall with the independent oracle. New, unrecognized Word conversion-loss diagnostics fail corpus preparation.

After a read pass, the corpus selects the last, middle, and first pages, splits that result into single-page documents, merges them again, and independently verifies order and token retention. OfficeIMO intentionally blocks unsafe full rewrites of documents whose forms, signatures, tagged content, active content, outlines, xref streams, or object streams cannot yet be preserved. The JSON report records these as `Blocked` with machine-readable mutation blocker codes, separately from failed output validation. Those blockers identify manipulation work to implement; the runner does not bypass them.

## Measurements

BenchmarkDotNet reports mean/median timing, statistical error, rank, GC collections, and managed allocated bytes per operation. `Allocated` is managed allocation volume, not peak memory or total resident memory. This distinction matters for QuestPDF because its Skia work uses native memory and for Chromium because the browser process is outside the benchmark host's managed heap.

The artifact evidence runner complements BenchmarkDotNet with a fresh worker process per engine and iteration. It samples the complete worker process tree, including Chromium and its descendants, from process start through renderer shutdown. Those sampled peak working-set values have an equivalent process boundary and are marked comparable; they remain sampled resident-memory observations rather than exact allocation accounting. The existing `OfficeIMO.Pdf.Benchmarks` budget runner remains the OfficeIMO-only source for sampled peak managed heap and retained writer-buffer evidence.

Deep deterministic-content validation remains outside measured operations. The manipulation benchmarks deliberately include the equivalent producer-native post-save reopen described above. Output byte length is observed but is not treated as a correctness substitute: compression and font subsetting legitimately produce different file sizes.

## Run

Generate a reviewable HTML-to-PDF evidence bundle before interpreting benchmark timings. This command renders the same deterministic HTML two or more times with OfficeIMO, PeachPDF, iText pdfHTML, and Chromium through HtmlTinkerX. It writes the source HTML, every PDF, first-page PNG previews, and `html-pdf-evidence.json`. The report records exact-byte, semantic, and visual repeatability; page and content checks; tagged-PDF structure; output size; cancellation capability; managed allocation volume; and sampled peak process-tree working set.

Every conversion iteration runs in a fresh worker. The coordinator validates and renders previews only after the worker exits, so its own PDF inspection and rasterization memory is excluded. The report records the sampler identity, sample count, observed process-count range, and peak working set for each worker tree. When Poppler's `pdftoppm` is on `PATH`, the runner also creates independent external previews. Use `--require-external-rasterizer` for a visual gate that must fail when Poppler is unavailable.

```powershell
$repoRoot = if ($env:EVOTEC_GITHUB_ROOT) { $env:EVOTEC_GITHUB_ROOT } else { 'C:\Support\GitHub' }
$env:HTMLTINKERX_PROJECT_PATH = Join-Path $repoRoot 'HtmlTinkerX/Sources/HtmlTinkerX/HtmlTinkerX.csproj'
$output = Join-Path 'Ignore/Benchmarks/HtmlPdfEvidence' (Get-Date -Format 'yyyyMMdd-HHmmss')
try {
    dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
        -c Release `
        -f net10.0 `
        -- html-evidence `
        --output $output `
        --scale Easy `
        --iterations 3 `
        --require-external-rasterizer `
        --require-clean-source
} finally {
    Remove-Item Env:HTMLTINKERX_PROJECT_PATH -ErrorAction SilentlyContinue
}
```

The evidence runner validates artifacts but is not a statistical performance runner. Continue to use BenchmarkDotNet through the shared script for performance results.

After a reviewed Windows or Linux High-scale run, validate every referenced PDF and preview against the report and write the compact committed summary. Raw PDFs and PNGs remain temporary; the summary retains their paths, sizes, SHA-256 hashes, aggregate manifest hash, exact source commits, contracts, and measurements:

```powershell
pwsh Build/Export-HtmlPdfArtifactEvidence.ps1 `
    -EvidencePath $output `
    -Platform windows `
    -OutputPath Docs/benchmarks/html-pdf-artifact-evidence/html-pdf-artifact-evidence-net10.0-windows-high.json

pwsh Build/Test-HtmlPdfArtifactEvidence.ps1
```

The release gate requires matching Windows and Linux summaries from the same clean OfficeIMO and HtmlTinkerX source commits. Any production renderer or evidence-runner change makes them stale. A package-pin-only HtmlTinkerX change is instead proven by the packed browser consumer gate because these runs compile the recorded HtmlTinkerX source checkout directly.

Run one quick correctness/performance smoke through the shared PowerForge evidence path:

```powershell
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfgenerate -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtml -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtmlpayload -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfformats -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfread -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfstructuredread -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfsplit -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfmerge -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfselect -RunMode quick -Framework net10.0
```

To validate a specific HtmlTinkerX source checkout, pass either its repository root or project path. The runner records the exact clean HtmlTinkerX commit and carries the project reference into BenchmarkDotNet child builds. Without this option, the benchmark uses the pinned package version:

```powershell
$repoRoot = if ($env:EVOTEC_GITHUB_ROOT) { $env:EVOTEC_GITHUB_ROOT } else { 'C:\Support\GitHub' }
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 `
    -Workload pdfhtml `
    -RunMode quick `
    -Framework net10.0 `
    -HtmlTinkerXRoot (Join-Path $repoRoot 'HtmlTinkerX')
```

Use `-RunMode full` for publication-quality BenchmarkDotNet statistics. `-Publish` is valid only with a full run and updates the shared benchmark evidence catalog. Raw BenchmarkDotNet artifacts stay under the ignored output root.

For a local short engineering run without catalog updates:

```powershell
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfGenerationBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfHtmlBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfHtmlPayloadBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*Pdf*FormatConversionBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfReadBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfStructuredRead*Benchmarks*" --job Short
```

Generate the deterministic semantic accuracy report separately from timings:

```powershell
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- semantic-evidence --scale Medium --output Ignore/Benchmarks/PdfComparisons/semantic-accuracy.json
```

The report records labelled-region character error rate, pairwise reading-order accuracy, Kendall tau, per-kind precision/recall/F1, heading detection and exact-level F1, logical-table detection F1, cell-adjacency structure F1, and cross-page continuation-pair F1. Its generated corpus is a regression gate, not an independent estimate of real-world accuracy. Cell adjacency is a labelled structure score, not a claim of full TEDS equivalence. Whole-document CER, full tree-edit-distance TEDS, and independent-producer generalization remain explicitly unmeasured until suitable corpus annotations exist.

### External structured-parser validation

Use the structured suites only for parsers that perform the same semantic work.
The raw `PdfReadBenchmarks` and a save/rewrite benchmark are not comparable to
layout reconstruction. A useful external result should record the exact source
commit, runtime, machine, corpus hash, selected profile, failures, elapsed time,
and managed allocations. Run the statistical suite and semantic scorecard
separately:

```powershell
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 `
    -Workload pdfstructuredread `
    -RunMode full `
    -Framework net10.0

dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
    -c Release `
    -f net10.0 `
    -- semantic-evidence `
    --scale High `
    --output Ignore/Benchmarks/PdfComparisons/semantic-accuracy-high.json
```

The generated fixture is suitable for repeatability and regression checks, but
not for a general competitive accuracy claim. For that, run every parser on the
same independently labelled, redistributable documents and require the same
output contract: characters, pairwise reading order, heading kind and level,
header/footer classification, table detection and cell adjacency or TEDS,
cross-page continuation pairs, elapsed time, allocations, and failure rate by
document class. Publish the annotation schema and corpus hashes with the result.

Prepare and validate the real-document corpus before running its BenchmarkDotNet lane:

```powershell
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- corpus --download --output Ignore/Benchmarks/PdfComparisons/corpus

pwsh Build/Run-LibraryComparisonBenchmarks.ps1 `
    -Workload pdfcorpusread `
    -RunMode quick `
    -Framework net10.0 `
    -PdfCorpusRoot Ignore/Benchmarks/PdfComparisons/corpus/files
```

The corpus workload is excluded from `-Workload all` because it requires opt-in downloads. Use `--only id-1,id-2` with the corpus command to rerun selected entries. `--additional-manifest path.json` appends one schema-1 manifest of generated or machine-local entries. Both the base and additional manifests are strict: unknown fields, invalid source combinations, missing local or downloaded SHA-256 hashes, out-of-range recall thresholds, and malformed semantic expectations fail before any document is read.

On Windows with Microsoft Word, Excel, and PowerPoint installed, this command regenerates all three Office-produced PDFs and validates OfficeIMO readback. COM remains an opt-in fixture producer and is not an OfficeIMO runtime dependency:

```powershell
pwsh Build/Run-PdfOfficeComCorpus.ps1 -Framework net10.0
```

## Benchmark-only libraries

Versions are pinned in the benchmark project so evidence is reproducible. PDFsharp/MigraDoc and HtmlTinkerX are MIT, PeachPDF is BSD-3-Clause, and PdfPig is Apache-2.0. QuestPDF uses its Community license for this open-source benchmark project. iText Core and pdfHTML are AGPL/commercial and remain isolated here for benchmark use; they are not linked by or distributed with an OfficeIMO runtime package.

The benchmark uses the maintained cross-platform PDFsharp 6.x package. PdfSharpCore 1.3.67 is not included: restoring it currently reports NuGet vulnerability advisories through its ImageSharp 1.0.4 dependency, while PDFsharp 6.x already covers the equivalent cross-platform split, merge, selection, and MigraDoc generation workflows.
