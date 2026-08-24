# Benchmark artifacts

This folder stores small, committed benchmark summaries and artifacts. Raw BenchmarkDotNet output, traces, and other machine-specific bulk evidence stay local.

## Current non-Excel, non-PDF, non-image posture

| Owner | Time and allocations | Peak memory and output size | Equivalent comparison | Current evidence gap |
| --- | --- | --- | --- | --- |
| CSV | BenchmarkDotNet read/write suites | Validated SQL-shaped output is 0.991-0.993x Sylvan's UTF-8 size; sequential and parallel OfficeIMO output is byte-identical | CsvHelper, Sep, Sylvan, Dataplat.Dbatools.Csv, and LumenWorks | Extend size evidence to other equivalent file-producing shapes and add isolated peak-memory evidence |
| Word | Validated BenchmarkDotNet create, read, report, and replace suites | DOCX payloads are validated; size is not exported by the shared runner | DocX, NPOI, and Open XML SDK | Add environment-qualified output-size evidence without publishing license-restricted numbers |
| PowerPoint | Repeated isolated package workflows are within the 2× contender ceiling for both time and allocation on Windows and Linux | Sampled managed-heap growth, process peak, and output bytes are recorded and budgeted | ShapeCrawler for validated create/save and open/edit/save | Large open/edit/save remains close to the ceiling at 1.70-1.82× on Windows and should be improved further |
| Reader | BenchmarkDotNet extraction, detection, transport, and chunking suites | External processes record peak working set; creation size is not applicable | Optional direct-process runners for equivalent extraction | Add representative application corpora and release baselines |
| Markdown | BenchmarkDotNet parse, HTML render, transform, and HTML-to-Markdown suites | The isolated parser runner records allocation, retained heap, managed-heap peak, process peak, input bytes, and semantic output bytes; parsing has no file-size metric | Markdig and ReverseMarkdown after semantic equivalence checks | HTML-to-Markdown and all seven CommonMark semantic parse corpora are within 2x for time and allocation; improve the narrowest margins and measure source-backed parsing separately |
| HTML | BenchmarkDotNet stages plus validated isolated non-PDF layout evidence; the optimized corpora allocate 18.6-49.8% less than the clean baseline | Retained heap, sampled managed peak, and absolute process peak are recorded and budgeted; file size is not applicable to the in-memory layout lane | No general renderer is equivalent across the complete paged vector and semantic contract | The 2,500-row table still allocates about 645 MiB and the 1,000-page lane about 393 MiB; continue owner-level optimization and add non-Windows evidence |
| RTF | Full validated RTF-to-HTML evidence shows OfficeIMO faster and lower-allocation in every corpus | Output is smaller on the three equivalent generated corpora; the higher-fidelity producer output remains 3.14x larger | RtfPipe for validated RTF-to-HTML | Reduce fidelity-preserving producer HTML overhead further and add Linux/macOS evidence |
| OpenDocument | BenchmarkDotNet open, sparse-write, formula, and validated ODS create/read comparisons | Evidence runner records peak working set and package input/output bytes; the ODS comparison exports validated size sidecars | OpenStandardLibrary for equivalent dense-string ODS create and read workloads | Add ODT/ODP comparisons only when another library can perform the same contract; capture non-Windows evidence |
| OneNote | BenchmarkDotNet native read, write, and Markdown projection plus isolated create/write, read, and read/edit/write workflows | Managed allocations, sampled managed-heap growth, process peak, input bytes, and output bytes are recorded | No like-for-like offline semantic `.one` competitor has been identified | Native scale evidence and regression budgets are in place; add non-Windows evidence and a competitor only when the same offline semantic contract exists |
| Email | Full validated EML read/write evidence is within 2x of MimeKit; the 2,000-item PST scale contract fell from 68.19s to 4.03s | EML retained heap, managed peak, and process peak are all within 2x; validated output is 1.07-1.08x MimeKit; the 4.08 MB PST retained 0.70 MB managed memory | MimeKit 4.17.0 for equivalent complete MIME workflows | Improve the 1.52x normal-write allocation margin and add non-Windows evidence; add a PST competitor only if the same managed creation contract is available |
| ZIP | Full validated safe-traversal evidence is 1.03-1.16x raw platform traversal by mean time and 1.00-1.01x by allocation | Isolated peak managed-heap growth matches the platform lane; output size is not applicable because both lanes consume the same ZIP | `System.IO.Compression` metadata projection, explicitly without OfficeIMO's safety policy | Add Linux/macOS evidence |
| EPUB | Full validated open/read evidence uses 0.26-0.37x VersOne's mean time and 0.24-0.28x its managed allocation | Retained result size is 0.94-0.99x; managed peak is 0.28-0.44x; both lanes consume the same input package | VersOne.Epub plus HtmlAgilityPack for equivalent metadata, raw-XHTML, visible-text, and spine-order extraction | Add Linux/macOS evidence; add creation evidence only if OfficeIMO gains an EPUB writer |
| LaTeX | Validated BenchmarkDotNet lossless parse and parse-plus-preserve-write curves plus isolated evidence; the optimized large lane is about 69% faster and allocates about 62% less than the initial baseline | Allocation, retained heap, sampled managed peak, process peak, input bytes, and byte-identical preserve output are recorded and budgeted | No general .NET parser performs the same public lossless syntax plus semantic workflow | The 701 KiB source still retains about 57 MiB across its 256,036-token public model; reduce graph overhead and add non-Windows evidence |
| Security, Provenance, AsciiDoc, Visio, and Markup | Reader coverage reaches some inputs, but there is no complete owner-level performance suite | Incomplete | Add a competitor only where the same public workflow exists | These are the next inventory gaps under the repository roadmap's performance-evidence item |

This table describes measurement coverage, not a library ranking. A returned
byte count is not yet durable size evidence unless the runner records it with
the source commit and environment. Open work belongs in
[`Docs/ROADMAP.md`](../ROADMAP.md), rather than a second benchmark backlog.

For equivalent competitor lanes, the working classification is: at most 2× in
both elapsed time and allocation is contender-level but may still warrant
improvement; above 2× through 5× is a material remediation gap; more than 5× is
unacceptable unless the contracts differ. A 40× ratio is an incident threshold,
never a success boundary.

## PowerPoint package workflows

`OfficeIMO.PowerPoint.Benchmarks` measures deterministic create/save and
open/edit/save workflows in fresh processes. The ShapeCrawler comparison uses
the same semantic validator and, for editing, the exact same source package.
The checked-in runner records allocation, sampled managed-heap growth, process
peak, input/output size, source commit, dirty-tree state, and environment.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.PowerPoint.Benchmarks -- --verify-budgets
dotnet run -c Release -f net8.0 --project .\OfficeIMO.PowerPoint.Benchmarks -- --operation OpenEditSave --repeat 5 --corpus-dir .benchmark-artifacts\powerpoint\corpus --json .benchmark-artifacts\powerpoint\officeimo.json
dotnet run -c Release -f net8.0 --project .\OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler -- --operation OpenEditSave --repeat 5 --corpus-dir .benchmark-artifacts\powerpoint\corpus --json .benchmark-artifacts\powerpoint\shapecrawler.json
```

The Windows and Ubuntu 24.04 evidence is recorded in
[`OfficeIMO.PowerPoint.Benchmarks/BASELINE.md`](../../OfficeIMO.PowerPoint.Benchmarks/BASELINE.md).
Every package lane is within the 2× contender ceiling for time and allocation,
but the large edit lane remains an explicit optimization target.

## CSV output size

The validated SQL-shaped DataReader writer now records UTF-8 output size and
hashes for OfficeIMO sequential, OfficeIMO parallel, and Sylvan.Data.Csv:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks -- --datareader-write-size-evidence --rows 25000,100000 --json .benchmark-artifacts\csv\datareader-write-size.json
```

The Windows evidence is recorded in
[`officeimo.csv-datareader-size-2026-08-24.md`](officeimo.csv-datareader-size-2026-08-24.md).
OfficeIMO output is 0.991-0.993x Sylvan's UTF-8 size across mixed, quoted, and
multiline shapes at 25,000 and 100,000 rows. Sequential and parallel OfficeIMO
outputs are byte-identical.

## Reader baselines

- `officeimo.reader.foundation-2026-07-10.md`: first Reader-wide extraction, detection, transport, and parser/chunker baseline after the P0 foundation work.

Reader benchmark code lives in `OfficeIMO.Reader.Benchmarks`.

## Markdown comparisons

`OfficeIMO.Markdown.Benchmarks` compares equivalent CommonMark parsing and HTML
rendering with Markdig, and equivalent HTML-to-Markdown conversion with
ReverseMarkdown. Setup rejects a corpus when the compared outputs do not have
the same normalized semantic HTML.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Markdown.Benchmarks -- --validate-equivalence
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload markdownparse -RunMode full -Framework net8.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload markdownhtml -RunMode full -Framework net8.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload htmltomarkdown -RunMode full -Framework net8.0
```

The benchmark classes do not pin a runtime job. `-f net8.0` or `-f net10.0`
selects the runtime, and `--job Dry` remains an execution check instead of
silently adding a full second job. The project is intentionally outside
`OfficeIMO.sln`, so Markdig and ReverseMarkdown remain opt-in comparison
dependencies rather than normal solution restore inputs.

The full Windows HTML-to-Markdown run is recorded in
[`officeimo.markdown-html-to-markdown-2026-08-24.md`](officeimo.markdown-html-to-markdown-2026-08-24.md).
Its three validated corpora are within 2x of ReverseMarkdown for both mean time
and managed allocation. This does not close the source-backed parsing lane,
which retains additional source and syntax ownership and remains outside the
contender boundary.

The full Windows CommonMark semantic parse run is recorded in
[`officeimo.markdown-commonmark-parse-2026-08-24.md`](officeimo.markdown-commonmark-parse-2026-08-24.md).
All seven validated corpora are within 2x for both mean time and managed
allocation. The narrowest allocation margins are Transcript at 1.91x, Large
table at 1.84x, Portable README at 1.83x, and Rich AST at 1.80x, so contender
status does not end optimization work. The isolated runner also confirms every
retained-heap, sampled managed-heap peak, and absolute process-peak ratio within
2x across the matrix.

## HTML non-PDF layout

`OfficeIMO.Html.Benchmarks` includes process-isolated complete layout workloads
for a report, paged purchase tables, forced-page long documents, and the strict
static standards surface. Every result validates text and page contracts before
it is accepted.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Html.Benchmarks -- --layout-verify-budgets --repeat 3 --json .benchmark-artifacts\html\layout-evidence.json
```

The clean before/after Windows evidence is recorded in
[`officeimo.html-layout-2026-08-24.md`](officeimo.html-layout-2026-08-24.md).
Allocations fell by 18.6-49.8% across all six workloads. The 2,500-row paged
table fell from 1,216.19 MiB to 645.31 MiB allocated and from 1.45 s to 0.91 s,
but remains an explicit optimization target. No competitor claim is made
without an equivalent paged vector and semantic render contract.

## RTF comparisons

`OfficeIMO.Rtf.Benchmarks.Comparisons` measures complete RTF-to-HTML parsing
and rendering through OfficeIMO and RtfPipe. It uses a shared-feature corpus at
12, 250, and 2,000 records plus a producer fixture. Validation checks required
text, complete record counts, tables, and cells before timing and records UTF-8
input/output sizes separately.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\rtf\validation.json
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload rtfhtml -RunMode full -Framework net8.0
```

The comparison project is outside `OfficeIMO.sln`; RtfPipe remains an opt-in
benchmark dependency and does not enter normal restore, build, or packages.
The shared runner captures the validated input/output-size sidecar with the
same source commit, framework, platform, and environment provenance as timing
and allocation evidence.

The full Windows run is recorded in
[`officeimo.rtf-html-2026-08-24.md`](officeimo.rtf-html-2026-08-24.md).
OfficeIMO is faster and lower-allocation in every corpus and produces less HTML
for all three generated corpora. The producer fixture's remaining 3.14x size
ratio is explicitly fidelity-qualified and remains an optimization signal.

## Email MIME comparisons

`OfficeIMO.Email.Benchmarks.Comparisons` measures complete EML parsing with
decoded attachment consumption and complete in-memory EML serialization through
OfficeIMO.Email and MimeKit. Both outputs are cross-read and checked for equal
envelope fields, normalized bodies, ordered attachment metadata, decoded lengths,
and payload hashes before size evidence is accepted.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\email\validation.json
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload emailmimeread -RunMode full -Framework net10.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload emailmimewrite -RunMode full -Framework net10.0
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- --evidence --repeat 3 --json .benchmark-artifacts\email\memory-evidence.json
```

The full Windows run is recorded in
[`officeimo.email-mime-2026-08-24.md`](officeimo.email-mime-2026-08-24.md).
All read and write lanes are within 2x for both mean time and managed allocation;
validated OfficeIMO output is 1.07-1.08x the MimeKit byte size. The comparison
project stays outside `OfficeIMO.sln`, so MimeKit remains benchmark-only. The
same evidence note records the separate 16.9x PST scale-test improvement from
removing disk-journal setup for single-block data trees. The isolated runner
also places retained heap, managed-heap peak, and absolute process peak within
2x for every read and write scale; normal-write allocation at 1.52x remains the
next optimization margin.

## OpenDocument comparisons

`OfficeIMO.OpenDocument.Benchmarks.Comparisons` measures complete in-memory ODS
creation and package open/read traversal through OfficeIMO and
OpenStandardLibrary. Both writers receive the same dense string-cell corpus.
Both readers receive the same OfficeIMO-generated package and enumerate every
populated cell. Validation reopens both outputs through OfficeIMO and checks
sheet, row, cell, content-length, and boundary-marker contracts before output
sizes are accepted.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\opendocument\validation.json
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload odscreate -RunMode full -Framework net8.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload odsread -RunMode full -Framework net8.0
```

The comparison project is outside `OfficeIMO.sln`; OpenStandardLibrary remains
an opt-in benchmark dependency. It is not used as an ODT or ODP comparator
because it does not expose equivalent document models for those formats.

## OneNote native workflows

`OfficeIMO.OneNote.Benchmarks` measures default validated writing separately from
serialization-only writing. Its isolated evidence runner covers deterministic
1-page, 25-page, and 100-page create/write, read, and read/edit/write workflows.
File-producing lanes are reopened after measurement and must match the exact
ordered semantic fingerprint.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.OneNote.Benchmarks -- --verify-budgets
dotnet run -c Release -f net10.0 --project .\OfficeIMO.OneNote.Benchmarks -- --filter '*WriteDesktopSection*'
```

The Windows native baseline and writer optimization evidence are recorded in
[`officeimo.onenote-native-2026-08-24.md`](officeimo.onenote-native-2026-08-24.md).
No contender ratio is claimed: no other managed library was found that provides
the same offline semantic `.one` read/write contract. The checked-in budgets
guard allocation, managed-heap growth, process peak, and output size; timing is
reviewed with BenchmarkDotNet and a same-machine healthy commit.

## ZIP traversal comparison

`OfficeIMO.Zip.Benchmarks` measures safe, deterministic OfficeIMO traversal
against direct `System.IO.Compression` metadata projection. Both lanes return
the same validated ordered descriptor fields; only OfficeIMO applies path and
expansion safety limits, so the comparison reports the cost of that policy.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- validate
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload ziptraverse -RunMode full -Framework net10.0
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- --evidence --repeat 5 --json .benchmark-artifacts\zip\evidence.json
```

The full Windows run is recorded in
[`officeimo.zip-traversal-2026-08-24.md`](officeimo.zip-traversal-2026-08-24.md).
All scales are within 1.16x of the raw platform lane for mean time, within 1.01x
for managed allocation, and show effectively identical peak managed-heap
growth. ZIP creation size is not compared because the package owns traversal
policy rather than archive writing.

## EPUB open/read comparison

`OfficeIMO.Epub.Benchmarks.Comparisons` compares complete stream-based EPUB 3
loads through OfficeIMO and the VersOne.Epub plus HtmlAgilityPack plain-text
extraction workflow documented by VersOne. Both lanes receive the same package,
load metadata and content, extract normalized visible chapter text, and
enumerate every chapter in spine order. The preflight requires identical title,
creator, language, ordered chapter paths, raw-XHTML and visible-text lengths and
hashes, and path hashes before measurements are accepted.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\epub\validation.json
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload epubread -RunMode full -Framework net8.0
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --evidence --repeat 3 --json .benchmark-artifacts\epub\comparison-evidence.json
```

The comparison project is outside `OfficeIMO.sln`; VersOne.Epub and
HtmlAgilityPack remain opt-in benchmark dependencies. The validated sidecar
records shared input size. There is no creation-size comparison because
OfficeIMO.Epub currently exposes a read/extraction API rather than an EPUB
package writer.

The full Windows run is recorded in
[`officeimo.epub-read-2026-08-24.md`](officeimo.epub-read-2026-08-24.md).
OfficeIMO uses 0.26-0.37x the comparator's mean time and 0.24-0.28x its managed
allocation. The isolated runner places retained heap, managed-heap peak, and
absolute process peak below 1x for both scales. EPUB therefore clears the 2x
contender ceiling without a runtime optimization in this slice.

## PDF comparisons

The opt-in `OfficeIMO.Pdf.Benchmarks.Comparisons` project measures validated
PDF workflows at easy, medium, high, and real-document scale:

- structured PDF generation with OfficeIMO, QuestPDF, MigraDoc/PDFsharp, and iText;
- identical-HTML rendering with OfficeIMO.Html.Pdf and PeachPDF;
- full-document text extraction with OfficeIMO.Pdf, PdfPig, and iText over a
  five-producer synthetic corpus;
- split, bundled split, merge, and reversed non-contiguous page selection with
  OfficeIMO, iText, and PDFsharp;
- prepared large-document extraction over Office exports, generated rich and
  500-page OfficeIMO documents, government publications, and pinned CC0 PDF/A,
  Type0, and Type3 fixtures.

The comparison project is deliberately outside `OfficeIMO.sln`, so its
third-party packages remain benchmark-only dependencies. See its README for the
equivalence checks, corpus provenance, Word COM workflow, mutation-blocker
interpretation, memory limits, and PowerForge runner commands.

## Word comparisons

`OfficeIMO.Word.Benchmarks` contains validated BenchmarkDotNet comparisons for
plain DOCX creation, structured reports, full paragraph traversal, and
replace-and-save workflows. It compares OfficeIMO.Word with DocX and the Open
XML SDK, plus opt-in NPOI 2.8.0, only where each implementation performs
equivalent work. The workload table and validators define the measured feature
set and the output contract each library must satisfy.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- validate
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload word -RunMode full -Framework net8.0 -AcceptNPOIOSMFLicense
```

The public repository keeps the benchmark code, inputs, validators, workload
contracts, and reproduction instructions. Raw BenchmarkDotNet and PowerForge
evidence stays in ignored or temporary output roots. Numerical Word comparison
results are local-only because the Xceed Community License requires advance
permission before publishing DocX benchmark or performance comparisons; the
shared runner enforces that boundary. The isolated benchmark project records
the acknowledgement required by NPOI's binary EULA; the EULA may also require
a maintenance fee for revenue-generating use, but it does not prohibit
benchmark publication and is not the reason results stay local.

## Excel artifacts

## Artifact types

- `officeimo.excel.snapshot-YYYY-MM-DD.json`: lightweight scenario snapshot for write, read, and round-trip flows.
- `officeimo.excel.write-profile-YYYY-MM-DD.json`: write-stage breakdown for optimization work.
- `officeimo.excel.read-profile-YYYY-MM-DD.json`: read-stage comparison for automatic, forced sequential, and forced parallel range conversion.
- `officeimo.excel.library-comparison.json`: local opt-in comparison across matching library surfaces.
- `officeimo.excel.npoi-comparison-current.json`: the current checked opt-in NPOI verification artifact for equivalent `.xlsx` row/cell writes and `.xls` read lanes, including scalar values, formulas, metadata, conditional formatting, AutoFilter range, style signals, and embedded pictures. The runner also supports XLS write comparisons through the paired command below. NPOI stays outside normal solution restore/build.
- `officeimo.excel.npoi-verification-notes.md`: benchmark-only scope notes for the opt-in NPOI runner.
- `officeimo.excel.datareader-table-2026-08-10.md`: dual-CPU evidence for package-native streaming `IDataReader` table writes, including the equivalent-contract validation policy.
- `comparison-current\officeimo.excel.comparison-suite-manifest.json`: release-style suite manifest.
- `comparison-current\officeimo.excel.comparison-summary.md|csv|json`: one-table decision summary with speed, allocation, and package-size ratios.
- `readme-current\officeimo.csv.comparison.json|officeimo.excel.comparison.json`: compact, PSPublishModule-compatible selections that generate the benchmark README tables.
- `officeimo.excel.comparison-report.md`: generated website/blog-oriented report distilled from comparison data.
- `Website\data\benchmarks-excel.json|benchmarks-excel-summary.json|benchmarks-excel-index.json`: generated website-facing benchmark data.

## Generate artifacts

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --snapshot .\Docs\benchmarks\officeimo.excel.snapshot-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-write .\Docs\benchmarks\officeimo.excel.write-profile-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --profile-read .\Docs\benchmarks\officeimo.excel.read-profile-YYYY-MM-DD.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --skip-legacy-epplus --warmup 20 --iterations 9
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --rows 2500 --warmup 1 --iterations 3 --out .\Docs\benchmarks\officeimo.excel.npoi-comparison-current.json
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks.NPOI\OfficeIMO.Excel.Benchmarks.NPOI.csproj -- --paired-xls-write --rows 25000 --warmup 12 --iterations 40 --affinity 0xFFFF --priority High
```

After a suite run, generate the website/blog data layer:

```powershell
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```

Refresh the marker-delimited CSV and Excel benchmark README tables through
PSPublishModule:

```powershell
.\Build\Benchmarks\Update-BenchmarkReadmes.ps1 -Run All
```

Use `-Run Csv` or `-Run Excel` when only one snapshot needs refreshing. With no
`-Run` value, the script simply regenerates the tables from the committed
compact JSON. The focused benchmarks run locally and are not scheduled in CI.
PSPublishModule owns generic Markdown replacement;
the repository script owns only benchmark invocation and OfficeIMO-specific row
selection. Generated comparison JSON is committed while raw output is ignored.
