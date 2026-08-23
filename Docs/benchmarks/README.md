# Benchmark artifacts

This folder stores small, committed benchmark summaries and artifacts. Raw BenchmarkDotNet output, traces, and other machine-specific bulk evidence stay local.

## Current non-Excel, non-PDF, non-image posture

| Owner | Time and allocations | Peak memory and output size | Equivalent comparison | Current evidence gap |
| --- | --- | --- | --- | --- |
| CSV | BenchmarkDotNet read/write suites | Output is validated; size is not a shared comparison metric | CsvHelper, Sep, Sylvan, Dataplat.Dbatools.Csv, and LumenWorks | Extend size evidence for file-producing lanes |
| Word | Validated BenchmarkDotNet create, read, report, and replace suites | DOCX payloads are validated; size is not exported by the shared runner | DocX, NPOI, and Open XML SDK | Add environment-qualified output-size evidence without publishing license-restricted numbers |
| PowerPoint | Isolated workflow runner records elapsed time and allocations | Peak working set and output bytes are recorded | ShapeCrawler for create/save and open/edit/save | Refresh Windows evidence and add a non-Windows baseline before setting budgets |
| Reader | BenchmarkDotNet extraction, detection, transport, and chunking suites | External processes record peak working set; creation size is not applicable | Optional direct-process runners for equivalent extraction | Add representative application corpora and release baselines |
| Markdown | BenchmarkDotNet parse, HTML render, transform, and HTML-to-Markdown suites | Managed allocations are recorded; text output bytes are not yet a shared metric | Markdig and ReverseMarkdown after semantic equivalence checks | Optimize the measured allocation gaps, then add output-size sidecars where size affects storage or transport |
| HTML | BenchmarkDotNet stage, pagination, Drawing, and PDF projection suites | Managed allocations are recorded; several output methods return byte counts | No general renderer is equivalent across the complete OfficeIMO contract | Add bounded peak-memory evidence for large non-PDF rendering workflows |
| RTF | BenchmarkDotNet plus regression budgets for parse, rewrite, and adapters | Budget runner records peak working set and output bytes | RtfPipe for validated RTF-to-HTML | Add Linux/macOS evidence and tune only after repeatable full runs |
| OpenDocument | BenchmarkDotNet open, sparse-write, and formula suites | Sparse write returns output length but no durable size report exists | No current managed .NET authoring library covers the same ODT/ODS/ODP contract | Add create/open/edit scorecards with output size and peak memory |
| OneNote | BenchmarkDotNet native read, write, and Markdown projection | Write returns bytes but size and peak memory are not recorded as evidence | No like-for-like offline semantic `.one` competitor | Add validated artifact-size and isolated peak-memory evidence |
| Email | Allocation/time regression tests plus bounded real-store evidence | Retained memory and source-read ceilings are covered | MimeKit is the credible MIME comparison target | Keep comparison work isolated from the runtime package and validate identical MIME semantics |
| ZIP, Security, Provenance, AsciiDoc, LaTeX, EPUB, Visio, and Markup | Reader coverage reaches some inputs, but there is no complete owner-level performance suite | Incomplete | Add a competitor only where the same public workflow exists | These are the next inventory gaps under the repository roadmap's performance-evidence item |

This table describes measurement coverage, not a library ranking. A returned
byte count is not yet durable size evidence unless the runner records it with
the source commit and environment. Open work belongs in
[`Docs/ROADMAP.md`](../ROADMAP.md), rather than a second benchmark backlog.

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
