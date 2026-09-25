# OfficeIMO PDF library comparisons

This opt-in project measures complete, validated PDF workflows. It is deliberately outside `OfficeIMO.sln`: QuestPDF, PeachPDF, PDFsharp/MigraDoc, PdfPig, iText, pdfHTML, HtmlTinkerX, and Chromium are benchmark tools, not OfficeIMO runtime dependencies. Public QuestPDF evidence is pinned to **QuestPDF 2026.5.0**, the last release with the embedded Community MIT License; that exact version is recorded in every result and displayed beside QuestPDF on the benchmark page.

## Workload matrix

| Scale | Real-world shape | Pages | Rows per page | Narrative paragraphs per page |
| --- | --- | ---: | ---: | ---: |
| Easy | Invoice or short status report | 1 | 8 | 1 |
| Medium | Monthly operational report | 20 | 12 | 2 |
| High | Annual audit archive | 100 | 12 | 3 |

Each page contains a heading, narrative text, and a four-column account/status table. The data is deterministic and created before measurement when it is an input. Every generated result must have a PDF header, the exact expected page count, and every required narrative and table row before its lane is accepted. Table validation treats numerically equivalent cells such as `1074.50` and `1074.5` as the same value while continuing to require exact text in non-numeric cells.

The benchmark families intentionally answer different questions:

- `PdfNativeOperationsBenchmarks`: OfficeIMO reads, selects, merges, and splits independently generated iText and MigraDoc inputs. Each source has 100 pages with four table rows and one narrative paragraph per page. Selection preserves a descending 25-page selection; merge combines 25 four-page documents. `Split` and `SplitSelections` both emit 100 single-page PDFs, with the latter exercising the compound selection API. Setup independently validates all input and output pages before timing. This lane compares OfficeIMO revisions and does not time the input producers.
- `PdfPreflightScalingBenchmarks`: OfficeIMO preflights 5-, 6-, 50-, and 500-page documents, both plain and with page labels, XMP metadata, and viewer preferences. Setup verifies that each input is readable, rewrite-safe, and has the expected page count and catalog features. This is an OfficeIMO revision-to-revision scaling check, not a third-party comparison.
- `PdfCatalogPageOperationScalingBenchmarks`: OfficeIMO splits every page and extracts a reversed, non-contiguous selection from 5-, 6-, 50-, and 500-page documents with no catalog features, one page-label range, a distinct range for every page, or a bookmark on every page. Setup reopens the outputs and checks page order, content markers, and applicable page labels or named destinations. This OfficeIMO-only lane exposes page-count and catalog-feature costs that the standard/dense comparison matrix does not exercise.
- `PdfMergeApiScalingBenchmarks`: OfficeIMO merges the same two validated sources through `MergeBytes`, `MergeWith(byte[])`, and `MergeWith(PdfDocument)` at 5, 6, 50, and 500 pages per source. Setup checks every output page and marker. This lane compares OfficeIMO API paths with each other; the three-engine `PdfMergeBenchmarks` remains the library comparison.
- `PdfPageImportScalingBenchmarks` and `PdfPageImportSourceScalingBenchmarks`: public page import inserts a one-page source at the start, middle, or end of 5-, 6-, 50-, and 500-page targets, then appends all pages from 5-, 6-, 50-, and 500-page sources to a two-page target. Setup validates every output page and its order. These lanes compare OfficeIMO revisions and expose target-size and source-size scaling separately.
- `PdfGenerationBenchmarks`: OfficeIMO, QuestPDF 2026.5.0, MigraDoc/PDFsharp, and iText generate the same structured report from the same logical model. The measured operation includes document construction, layout, font embedding, compression, and in-memory serialization.
- `PdfInvoiceGenerationBenchmarks`: OfficeIMO.Pdf, QuestPDF 2026.5.0, and iText directly compose the same two-page branded invoice from one prepared model, with the same logo, fonts, and visible content. Setup reopens every PDF and requires the page count, parties, seller VAT identifier, dates, every line description/quantity/unit price/VAT/net amount, intermediate totals, purchase order, payment data, note, terms, payable amount, and approval names before timing begins. This is a PDF composition-engine comparison; it deliberately does not use OfficeIMO's typed invoice renderer.
- `PdfTypedInvoiceWorkflowBenchmarks`: OfficeIMO alone captures the typed invoice, calculates and serializes its EN 16931 CII representation, renders the branded visible document, and embeds the exact electronic invoice bytes. Setup validates both visible content and the attachment. It is an absolute workflow-cost measurement, not a competitive ranking against general-purpose PDF libraries.
- `PdfHtmlBenchmarks`: OfficeIMO.Html.Pdf, PeachPDF, iText pdfHTML, and Chromium through HtmlTinkerX parse and render the exact same HTML string. Every engine emits tagged PDF bytes and must preserve the exact page count, narrative, and table content before its measurements are accepted. The managed engines include HTML/CSS parsing, paged layout, and in-memory serialization. Chromium reuses one HtmlTinkerX-owned browser session per benchmark case; each measured operation still replaces and reparses the complete page before printing, so warmed browser throughput is not mislabeled as process startup.
- `PdfHtmlPayloadBenchmarks`: OfficeIMO.Html.Pdf and PeachPDF render exact 21 KiB plain-text, table-heavy, and multilingual HTML payloads as tagged PDFs. The multilingual lane makes the same bundled Carlito font the primary CSS family for both engines and requires every measured Latin, Greek, and Cyrillic sample plus the embedded font in the resulting artifact. It therefore runs portably without host-font dependencies or role mismatches. The quick runner uses BenchmarkDotNet's process-isolated `Dry` job for cold-start evidence; full runs measure warmed throughput. Cleanup reopens each result, checks page count, first/last content, the unique terminal marker, all multilingual samples, and reports HTML bytes, PDF bytes, pages, and extracted-text length.
- `PdfFormatConversionBenchmarks` and `PdfExtendedFormatConversionBenchmarks`: all fourteen advertised OfficeIMO source routes parse deterministic source bytes and produce a PDF in one measured operation. DOCX, XLSX, PPTX, HTML, Markdown, RTF, AsciiDoc, LaTeX, MHTML, OneNote, ODT, ODS, ODP, and Visio outputs are reopened independently; every lane requires all four semantic fields for each of 120 records and reports source bytes, PDF bytes, pages, and extracted-text length. This is an OfficeIMO route-health benchmark, not a third-party comparison: adapters with materially different format contracts are not forced into artificial parity. The shared runner can execute this local health lane, but never writes it to the library-comparison evidence catalog, including when `all` or `-Publish` selects it.
- `PdfTableLayoutScalingBenchmarks`: OfficeIMO.Pdf composes a worksheet-shaped seven-column table with shrink-to-fit enabled at 5, 6, 30, and 120 data rows. Separate plain-cell and explicit-rich-font cases cover both the lightweight and repeated-probe sizing paths. Each measured operation includes document construction, layout, pagination, and serialization. Cleanup reopens the artifact and checks every row marker, so the suite exposes both the 5-to-6 boundary and sustained allocation growth without accepting missing middle rows.
- `PdfWordStructureScalingBenchmarks`: OfficeIMO.Word loads DOCX bytes and converts multi-level lists, multi-paragraph tables, nested tables, and a single cell containing 5, 6, 30, or 120 paragraphs. Each operation includes package loading, structural projection, PDF layout, pagination, and serialization. Cleanup reopens the artifact, checks every generated paragraph marker, and requires each list item to remain a distinct rendered marker line across three indentation levels.
- `PdfStructuredFormatScalingBenchmarks`: OfficeIMO.Markdown, OfficeIMO.Excel, OfficeIMO.Word, and OfficeIMO.Html parse deterministic structured inputs containing headings, lists, worksheet-style tables, and nested-table content at 5, 6, 30, and 120 records, then render them to PDF. Each format is a separate logical benchmark group because its source uses native format structures; ranks compare record-count scaling only within that format and are not cross-format rankings. Each operation includes source parsing, document projection, PDF layout, pagination, and serialization. Cleanup reopens every PDF and validates all four semantic fields for every generated record so scaling evidence cannot pass after silently dropping content.
- `PdfReadBenchmarks`: OfficeIMO.Pdf, PdfPig, and iText open identical bytes, enumerate every page, and extract the complete text payload. The corpus is repeated for OfficeIMO-, QuestPDF 2026.5.0-, PeachPDF-, MigraDoc-, and iText-produced PDFs to avoid a single-producer result.
- `PdfStructuredReadFastBenchmarks` and `PdfStructuredReadCompleteBenchmarks`: separate OfficeIMO-only route-health suites for the one canonical `PdfDocument.Load(...).Read(...)` contract. Both routes include source snapshotting, parsing, glyph recovery, word/line grouping, recursive XY-cut reading order, semantic classification, logical projection, and table extraction; `Structured` additionally applies document-wide evidence. Keeping the profiles in separate BenchmarkDotNet classes prevents shared ranks or baselines between unequal work. The runner excludes both suites from comparison publication, both gate their page/table invariants, and `Structured` additionally gates the labelled document-wide semantic contract it promises. Its labelled Easy fixture uses two pages, rather than the general one-page Easy scenario, because running-header/footer recovery requires repeated document evidence. The deterministic benchmark and scorecard raise only their document-wide work ceiling as a function of fixture page count so the 100-page case measures the complete route; production read defaults remain unchanged.
- `PdfSplitBenchmarks`: OfficeIMO, iText, and PDFsharp split the same OfficeIMO- and iText-produced documents into single pages and fixed-size bundles, then reopen every output with the producing engine and verify its page count inside the timed operation.
- `PdfMergeBenchmarks`: all three engines merge the same ordered source set, reopen the serialized output with the producing engine inside the timed operation, and preserve the exact page-marker sequence.
- `PdfPageSelectionBenchmarks`: all three engines extract the same reversed, non-contiguous page selection, reopen the serialized output with the producing engine inside the timed operation, and preserve that order.

- `PdfReverseConversionBenchmarks`: five isolated PDF producers feed the same DOCX, HTML, XLSX-table, editable-PPTX, ODT, ODS, editable-ODP, and PNG reconstruction routes. Global setup validates the complete generated source, reopens every target artifact, checks page scope and editable structures, and requires deterministic narrative or table-row retention before timings are accepted; timed methods include PDF parsing, projection, and target serialization. Run this lane through `Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfreverse`.
- `PdfCorpusReadBenchmarks`: OfficeIMO, PdfPig, and iText extract four prepared large documents: a 500-page OfficeIMO PDF, the 492-page NIST SP 800-53 Rev. 5, an 85-page Type3-font fixture, and a 258-page 12.9 MB PDF/A fixture.

The split, merge, and page-selection suites use 5-, 20-, 100-, and 500-page source scales. Each scale runs both a standard page (one narrative paragraph and four table rows) and a dense page (three paragraphs and twelve rows). At 500 pages, selection extracts 125 pages in reverse order, merge combines 50 ten-page documents, and split emits either 500 single-page files or 25-page bundles. Every scale and density keeps the same source and output validation for each engine.

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

Qualify the file-backed H4 corpus across the three HTML output intents before
interpreting individual screenshots or PDF files:

```powershell
$output = Join-Path $env:TEMP ("OfficeIMO-H4-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
    -c Release `
    -f net10.0 `
    -- html-corpus-evidence `
    --corpus advanced-held-out `
    --output $output `
    --verify-acceptance `
    --require-clean-source
```

The command renders every selected H4 source as an OfficeIMO screen PNG, print PDF,
screen-to-PDF artifact, and per-page scene PNG/SVG; a PeachPDF print PDF; and a
Chromium screen PNG and print PDF. It records text-marker preservation, normalized
text overlap, selected-element geometry, page counts, pixel differences, versions,
hashes, timings, allocations, diagnostics, and individual provider failures in
the schema-4 `html-corpus-evidence.json`. The report compares both managed print PDFs
directly with Chromium print as well as with each other, so a difference between
the managed engines cannot be mistaken for browser accuracy. The comparison
uses A4 pages and zero caller-supplied outer margins for OfficeIMO; authored
`@page` rules still apply. Screen markers require an exact normalized
phrase. PDF extraction uses ordered normalized tokens so table reading order may
interleave cells without allowing marker words to pass out of order; the policy and
result for every marker are recorded. With `--verify-acceptance`, the runner also writes the
human-readable `html-corpus-acceptance.md` report and exits unsuccessfully when a
required criterion, case, or capability selection fails.
With `--require-clean-source`, the runner also requires the loaded OfficeIMO renderer
assembly to identify the same Git commit as the clean checkout; rebuild after a commit
instead of reusing a `--no-build` result from an older head.

Poppler's `pdftoppm` and `pdftotext` must be on `PATH`. External PDF page images,
page counts, and text are observed through Poppler so OfficeIMO.Pdf is not its own
comparison oracle. Reference engines may omit text that OfficeIMO intentionally
preserves, including form values, image alternatives, and SVG labels; those remain
visible in each reference's `missingMarkers` field. The H4/advanced-held-out gate
requires OfficeIMO to preserve every declared source marker and applies the exact
per-case screen, print, and screen-to-page criteria in the checked-in
`acceptance.json`. Every case and applicable capability selection must pass. Chromium
provides the screen and print reference; it is not a universal oracle, so intentional
sheet, extraction, font, clipping, and managed-layout differences are admitted only
through a named classification and bounded criteria. Screen-to-page qualification
also requires sequential pages, a fixed 640 x 900 pixel canvas, the expected page
count, uniform page dimensions, and exact final-page padding before pixel tolerances
are considered. Raw artifacts remain in the
caller-selected output directory for visual review at normal reading size.

H4/advanced-held-out is the independently authored held-out selection. Its original
source bytes and hashes are preserved in the bundle, including legacy encodings.
Pass `--corpus representative` to use the established qualification corpus.
`--case <id>` is available for focused inspection and cannot be combined with
`--verify-acceptance`, because a partial run cannot qualify the corpus.
`--require-clean-source` fails closed unless Git proves both an exact commit and an
empty tracked and untracked status.

For an unfamiliar public page, use the opt-in MHTML lane to freeze the loaded
document and its captured resources, then replay those exact bytes with network
access disabled in Chromium. OfficeIMO and PeachPDF read the same archive through
their MHTML resource loaders:

```powershell
$capture = Join-Path ([System.IO.Path]::GetTempPath()) ('OfficeIMO-html-capture-' + [guid]::NewGuid())
$replay = Join-Path ([System.IO.Path]::GetTempPath()) ('OfficeIMO-html-replay-' + [guid]::NewGuid())
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
    -c Release -f net10.0 -- html-mhtml-evidence `
    --url https://www.w3.org/WAI/tutorials/tables/ --output $capture

dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
    -c Release -f net10.0 -- html-mhtml-evidence `
    --mhtml (Join-Path $capture 'source.mhtml') --output $replay `
    --replay-browser --require-clean-source
```

Choose a page you are permitted to capture and retain, and use new output
directories. The live capture waits for `DOMContentLoaded` plus one second; it
does not establish readiness for every scripted application. The replay is the
comparable input: Chromium opens the frozen MHTML offline, while OfficeIMO emits
print, screen-media-paged and screen-snapshot-paged PDFs and PeachPDF emits a
print PDF. The print diagnostics also include `officeimo-print-zero-margin`,
which removes OfficeIMO's default page margins, and
`officeimo-print-zero-margin-local-fonts`, which additionally opts in to
embedding document-selected fonts installed on the host. The
`officeimo-print-fit-1200-local-fonts` lane applies the same font policy to
the explicit 1200px print layout width. The local-font lanes
can differ across machines and do not change the library default. Compare
those lanes with their explicit settings; equal page counts do not establish
equal layout. Browser print and screen-media print are separate references;
page-count differences across intents are not failures by themselves. The JSON
records the archive hash, versions, source commit and dirty state, page counts,
MIME diagnostics, operation failures, and each OfficeIMO PDF intent's conversion
warnings and loss status. Explicit screen intents use the MHTML-aware resource
path, so embedded stylesheets and images remain available under the same archive
policy as print. `--require-clean-source` also checks that the loaded
Core, Email, HTML Core, HTML AngleSharp, HTML, HTML PDF, MHTML, MHTML PDF,
PDF and evidence-runner assemblies identify the clean
commit. Inspect PDF text and rendered
pages before making a compatibility claim. This lane is a diagnostic first pass,
not the H10 corpus acceptance gate or a sandbox for arbitrary hostile sites.

For a frozen archive that exceeds the untrusted CSS rule limit, append
`--max-css-rules 20000` to run an explicit bounded qualification profile. The
report records the effective rule and selector limits. This does not change
the library's default 10,000-rule untrusted profile; measure time and memory
before deciding whether a higher limit is suitable for ordinary use.

Measure and enforce the H4/advanced-held-out OfficeIMO static-rendering budgets separately from
the reference-engine comparison:

```powershell
$output = Join-Path $env:TEMP ("OfficeIMO-H4-budget-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
dotnet run --project Build/HtmlStaticBudget/OfficeIMO.Html.StaticBudget.csproj `
    -c Release `
    -- `
    --iterations 3 `
    --require-clean-source `
    --output $output
```

The command uses isolated cold and warmed worker processes. Each measured iteration
renders all eight held-out cases through screen PNG/SVG, print PDF/PNG/SVG, and
screen-to-page PDF/PNG/SVG. The report records elapsed time, managed allocations,
process-tree peak working set, output bytes, deterministic fingerprints, source and
corpus hashes, and asynchronous cancellation latency. The gate compares the cold
fingerprint with every warmed iteration as well as checking warm repeatability. Ceilings are absolute,
platform-specific regression limits in the frozen `budgets.json`; they are not a
general throughput claim. Use `--measure-only` when calibrating a new platform or
runtime, then review the raw report before changing a ceiling.

Run the provider-neutral runtime conformance suite against the comparison-only
Chromium/Playwright adapter:

```sh
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons -c Release -f net10.0 -- html-runtime-conformance
```

The adapter stays outside the normal solution and package graph. It proves that
the public host/context/page contract can serve an external browser without
exposing Playwright handles through common OfficeIMO APIs.

Compare one isolated public-page result with Chromium using the exact bytes
retained by `HtmlPublicPilot`:

```powershell
$pilot = 'Ignore/HtmlPublicPilot/wpt-svg'
$reference = Join-Path $pilot 'browser-reference'
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
    -c Release `
    -f net10.0 `
    -- html-public-browser-evidence `
    --acquisition (Join-Path $pilot 'acquisition.json') `
    --officeimo-screen (Join-Path $pilot 'screen.png') `
    --document-url https://wpt.live/css/compositing/line-with-svg-background-ref.html `
    --output $reference
```

The acquisition must have been produced with `--retain-input` and contain only
direct static GET responses. Dynamic requests and redirect chains are rejected
because their complete exchange bytes are not retained by this evidence format.
Chromium runs with service workers blocked and receives those retained resources
through context-wide request interception; an unrecorded or non-GET request
aborts the run. The new output directory contains both screen images, a pixel
difference image, hashes, dimensions, browser/tool versions, and comparison
metrics. This is an opt-in qualification oracle: HtmlTinkerX, Playwright, and
Chromium remain benchmark dependencies and do not enter the OfficeIMO runtime
package graph.

Generate the reviewable direct-PDF invoice bundle before interpreting invoice benchmark timings. This route renders the same prepared two-page invoice through OfficeIMO.Pdf, QuestPDF 2026.5.0, and iText, validates the complete text and numeric contract, and writes three PDFs plus two PNG page previews per engine. `invoice-evidence.json` hashes every artifact and records the exact OfficeIMO commit/tree, source cleanliness, target framework, runtime, operating system, process architecture, and QuestPDF package and assembly versions.

```powershell
$output = Join-Path 'Ignore/Benchmarks/PdfInvoiceEvidence' (Get-Date -Format 'yyyyMMdd-HHmmss')
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj `
    -c Release `
    -f net10.0 `
    -- invoice-evidence `
    --output $output `
    --require-clean-source
```

The comparison dependencies remain local benchmark tools. Review their current license terms before publishing the PDFs, previews, source adaptations, or benchmark claims outside this repository.

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
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfinvoice -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfinvoiceworkflow -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtml -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtmlpayload -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfformats -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfread -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfstructuredread -RunMode quick -Framework net10.0
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfnative -RunMode full -Framework net8.0
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

The real-world invoice lane is deliberately local-only and is not written to the
shared evidence catalog. Its normal build uses QuestPDF 2026.5.0 under that release's
Community MIT License. QuestPDF 2026.6.0 and later use different terms that restrict
use by competing PDF products; the pinned public evidence therefore must not be
described as a current-version QuestPDF result. iText remains AGPL/commercial. Review
the controlling license before publishing comparative artifacts or distributing
benchmark binaries. These restrictions do not apply to the OfficeIMO-only showcase.

Maintainers who have independently confirmed that they are authorized to exercise a
different QuestPDF release can use the isolated internal runner. The explicit switch
is a safeguard and provenance record, not a license grant. Internal runs reject
website publication, never update the shared catalog, and delete their temporary
artifacts unless `-KeepArtifacts` is supplied:

```powershell
pwsh Build/Run-InternalQuestPdfBenchmarks.ps1 `
    -QuestPdfPackageVersion 2026.9.0 `
    -Workload pdfgenerate `
    -RunMode quick `
    -QuestPdfLicenseType Community `
    -ConfirmQuestPdfAuthorization
```

For a local short engineering run without catalog updates:

```powershell
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfGenerationBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfInvoiceGenerationBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfTypedInvoiceWorkflowBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfHtmlBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfHtmlPayloadBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*Pdf*FormatConversionBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfTableLayoutScalingBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfWordStructureScalingBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfStructuredFormatScalingBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfReadBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfStructuredRead*Benchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfPreflightScalingBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfCatalogPageOperationScalingBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfMergeApiScalingBenchmarks*" --job Short
dotnet run --project OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -- --filter "*PdfPageImport*ScalingBenchmarks*" --job Short
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

Versions are pinned in the benchmark project so evidence is reproducible. Public QuestPDF measurements use QuestPDF 2026.5.0 under its embedded Community MIT License and display that version in the evidence. Newer QuestPDF versions are not silently substituted. PDFsharp/MigraDoc and HtmlTinkerX are MIT, PeachPDF is BSD-3-Clause, and PdfPig is Apache-2.0. iText Core and pdfHTML are AGPL/commercial and remain isolated here for benchmark use; none of these libraries is linked by or distributed with an OfficeIMO runtime package.

The benchmark uses the maintained cross-platform PDFsharp 6.x package. PdfSharpCore 1.3.67 is not included: restoring it currently reports NuGet vulnerability advisories through its ImageSharp 1.0.4 dependency, while PDFsharp 6.x already covers the equivalent cross-platform split, merge, selection, and MigraDoc generation workflows.
