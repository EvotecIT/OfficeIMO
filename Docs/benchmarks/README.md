# Benchmark artifacts

This folder stores small, committed benchmark summaries and artifacts. Raw BenchmarkDotNet output, traces, and other machine-specific bulk evidence stay local.

## Current non-Excel, non-PDF, non-image posture

| Owner | Time and allocations | Peak memory and output size | Equivalent comparison | Current evidence gap |
| --- | --- | --- | --- | --- |
| CSV | BenchmarkDotNet read/write suites | Validated SQL-shaped output is 0.991-0.993x Sylvan's UTF-8 size; sequential and parallel OfficeIMO output is byte-identical | CsvHelper, Sep, Sylvan, Dataplat.Dbatools.Csv, and LumenWorks | Extend size evidence to other equivalent file-producing shapes and add isolated peak-memory evidence |
| Word | Every validated create, read, and rich replace lane is within 2x of Open XML SDK; the former 100-paragraph breach is now 1.86x time and 1.41x allocation | Every isolated managed/process peak is within 1.63x; equivalent DOCX output sizes differ by less than 0.2% | Open XML SDK for public equivalent evidence; DocX and NPOI remain local license-restricted lanes | Move the remaining 1.54-1.87x margins toward parity and add Linux/macOS evidence |
| PowerPoint | Repeated isolated package workflows are within the 2× contender ceiling for both time and allocation on Windows and Linux; the worst current margins are 1.36× elapsed and 1.38× allocation | Sampled managed-heap growth, process peak, and output bytes are recorded and budgeted; unsigned saves no longer clone the full package for signature inspection | ShapeCrawler for validated create/save and open/edit/save | Move the remaining large open/edit/save margin toward parity without weakening signature mutation policy |
| Reader | BenchmarkDotNet extraction, detection, transport, and chunking suites; table-heavy Markdown is now 96.4% faster / 95.7% lower-allocation than the initial baseline, 10,000 repeated XML siblings are 10.3x faster, and equivalent YAML is within 1.85x time / 1.88x allocation of JSON | Isolated retained heap, sampled managed peak, process working-set growth, input bytes, and normalized output bytes are recorded per format; the direct Markdown output is hash-identical before and after optimization | JSON is the equivalent internal structured-output floor; optional direct-process runners cover equivalent extraction subsets; no accepted general competitor for the complete rich result contract | The 20-entry Markdown ZIP lane still allocates 46.26 MB; reduce the remaining aggregate graph, move YAML further inside the contender margin, add application corpora and cross-platform evidence, and accept a competitor only when it materializes equivalent rich results |
| Markdown | BenchmarkDotNet parse, HTML render, transform, and HTML-to-Markdown suites; all seven equivalent CommonMark semantic parse corpora are within 2x of Markdig | The isolated parser runner records allocation, retained heap, managed-heap peak, process peak, input bytes, and semantic output bytes; the compact source-backed pass reduced allocation another 3.5-6.4% and retained heap 7.2-15.8%; parsing has no file-size metric | Markdig and ReverseMarkdown after semantic equivalence checks; Markdig is only a diagnostic floor for the richer source-backed contract | The source-backed nested-list lane is still 4.09x semantic time / 3.12x semantic allocation and 4.79x / 4.32x the Markdig diagnostic floor; reduce source/trivia/syntax graph cost toward 2x without weakening snapshot semantics |
| HTML | BenchmarkDotNet stages plus validated isolated non-PDF layout evidence; the optimized corpora allocate 18.6-49.8% less than the clean baseline | Retained heap, sampled managed peak, and absolute process peak are recorded and budgeted; file size is not applicable to the in-memory layout lane | No general renderer is equivalent across the complete paged vector and semantic contract | The 2,500-row table still allocates about 645 MiB and the 1,000-page lane about 393 MiB; continue owner-level optimization and add non-Windows evidence |
| RTF | Full validated RTF-to-HTML evidence shows OfficeIMO faster and lower-allocation in every corpus | Output is smaller on the three equivalent generated corpora; the higher-fidelity producer output remains 3.14x larger | RtfPipe for validated RTF-to-HTML | Reduce fidelity-preserving producer HTML overhead further and add Linux/macOS evidence |
| ADF | Typed parse is 0.98-1.18x platform-model time and 1.24-1.27x allocation; semantic round trips are 0.99-1.02x time and 1.52-1.65x allocation | Retained heap and managed/process peaks are within 1.51x; round-trip output is byte-identical in size | Benchmark-only typed `System.Text.Json` model preserving nodes, marks, attributes, and extension data; it does not perform ADF validation | Add Linux/macOS evidence and a package competitor only when it can perform the same licensed parse/preserve/validate/write contract |
| Confluence | Managed-section replacement is 12.8-24.6% faster and allocates 63.4-65.6% less than the initial local baseline | Isolated allocation, retained heap, managed/process peak, input/output bytes, and exact SHA-256 hashes are recorded | No accepted competitor; a benchmark-only duplicate of OfficeIMO marker validation, exact replacement, and two hashes would be synthetic | Add Linux/macOS evidence and reduce compatible `netstandard2.0` / .NET Framework body-construction copies |
| Google Workspace | Declared-length buffering is 4.77x faster than the initial 4 MiB path; unknown-length buffering moved from 3.75x time / 2.97x allocation versus declared to 1.30x / 1.00x | Isolated allocation, retained heap, managed/process peak, exact input/output bytes, and SHA-256 are recorded for both response modes | Raw `HttpClient` is a non-equivalent floor because it omits retry, timeout, safety, diagnostics, response-limit, and mutation-outcome contracts | Add Linux/macOS evidence; measure resumable Drive upload/download and checkpoint overhead separately |
| OpenDocument | ODS creation is 0.08-0.18x OpenStandardLibrary time with lower allocation; read is 0.03-0.07x time with 1.25-1.26x allocation | Every isolated allocation/managed/process-peak ratio is within 1.63x; equivalent OfficeIMO output is 35.2-41.3% smaller | OpenStandardLibrary for equivalent dense-string ODS create and read workloads | Add ODT/ODP comparisons only when another library can perform the same contract; capture non-Windows evidence |
| OneNote | BenchmarkDotNet native read, write, and Markdown projection plus isolated create/write, read, and read/edit/write workflows | Managed allocations, sampled managed-heap growth, process peak, input bytes, and output bytes are recorded | No like-for-like offline semantic `.one` competitor has been identified | Native scale evidence and regression budgets are in place; add non-Windows evidence and a competitor only when the same offline semantic contract exists |
| Email | Full validated EML read/write evidence is within 2x of MimeKit; the 2,000-item PST scale contract fell from 68.19s to 4.03s | EML retained heap, managed peak, and process peak are all within 2x; validated output is 1.07-1.08x MimeKit; the 4.08 MB PST retained 0.70 MB managed memory | MimeKit 4.17.0 for equivalent complete MIME workflows | Improve the 1.52x normal-write allocation margin and add non-Windows evidence; add a PST competitor only if the same managed creation contract is available |
| MHTML | Complete validated reads are 0.88-1.08x MimeKit plus AngleSharp time and 0.62-0.92x allocation; writes are 0.67-0.73x MimeKit time and 0.70-1.40x allocation | Retained heap, managed peak, process peak, decoded bytes, and output bytes are recorded; OfficeIMO output is 0.864-0.973x MimeKit size | MimeKit 4.17.0 plus AngleSharp 1.7.1 for equivalent MIME, HTML DOM, and decoded-resource workflows | Reduce the 1.40x large-write allocation and approximately 1.46x managed-peak margin and add Linux/macOS evidence |
| ZIP | Full validated safe-traversal evidence is 1.03-1.16x raw platform traversal by mean time and 1.00-1.01x by allocation | Isolated peak managed-heap growth matches the platform lane; output size is not applicable because both lanes consume the same ZIP | `System.IO.Compression` metadata projection, explicitly without OfficeIMO's safety policy | Add Linux/macOS evidence |
| EPUB | Full validated open/read evidence uses 0.26-0.37x VersOne's mean time and 0.24-0.28x its managed allocation | Retained result size is 0.94-0.99x; managed peak is 0.28-0.44x; both lanes consume the same input package | VersOne.Epub plus HtmlAgilityPack for equivalent metadata, raw-XHTML, visible-text, and spine-order extraction | Add Linux/macOS evidence; add creation evidence only if OfficeIMO gains an EPUB writer |
| LaTeX | Validated BenchmarkDotNet lossless parse and parse-plus-preserve-write curves plus isolated evidence; the optimized large lane is about 79% faster and allocates about 69% less than the initial baseline | Allocation, retained heap, sampled managed peak, process peak, input bytes, and byte-identical preserve output are recorded and budgeted; the latest compact-graph pass cut large retained memory another 22.1% | No general .NET parser performs the same public lossless syntax plus semantic workflow | The 701 KiB source still retains about 44.4 MiB across its 256,036-token public model; reduce graph overhead and add non-Windows evidence |
| AsciiDoc | Validated BenchmarkDotNet lossless parse and parse-plus-preserve-write curves plus isolated evidence; source-backed text and compact syntax spans reduce managed allocation by about 39% at every scale | Allocation, retained heap, sampled managed peak, process peak, input bytes, and byte-identical preserve output are recorded and budgeted; the compact graph pass reduced large retained memory another 13.2% | No accepted comparison: AsciiDocNet documents unsupported tables and list continuations required by this corpus | The 558 KiB large source still allocates 29.34 MiB and retains 19.50 MiB; reduce collection and semantic graph overhead, add non-Windows evidence, and compare only equivalent public workflows |
| Markup | Validated CommonMark-plus-tables semantic parse is 1.15-1.31x Markdig's mean time and 1.13-1.26x its allocation | Isolated allocation, retained heap, managed peak, process peak, input bytes, event count, and digest are recorded and budgeted; file size is not applicable to parsing | Markdig 1.3.2 after exact semantic-event validation | Move the remaining 1.25-1.31x margins toward parity and add Linux/macOS evidence |
| Security | Detached RSA CMS signing is 0.87-0.94x platform time and 0.025-1.03x allocation; verification is 0.68-1.87x time and 0.036-1.79x allocation | Every isolated time/allocation/retained/managed/process ratio is below 2x; OfficeIMO CMS output is 1.105x platform size because it includes algorithm protection | .NET `SignedCms` for equivalent signature-only detached CMS workflows, with Bouncy Castle retained for richer structures and older targets | Reduce the remaining 1.68-1.87x platform-signature time margins and 1.69-1.79x OfficeIMO-signature small-allocation margins; add Linux/macOS evidence |
| Provenance | Validated structural inspect/remove curves cover PNG, TIFF, SVG, ZIP, and text; the 1 MiB PNG lane is 16.0x faster to inspect and 13.7x faster to remove than the initial local baseline | Isolated retained heap, managed batch peak, process peak, input bytes, and exact removal-output bytes are recorded | No accepted managed .NET comparator exposes the same bounded inspect-and-selective-remove contract | Reduce large SVG, text, and ZIP allocation; profile the remaining PNG CRC/JUMBF cost; add Linux/macOS evidence |
| Visio | Validated BenchmarkDotNet creation/save and load/structural-inspection curves; creation allocation is 16.0-64.4% below the initial baseline and large load allocation is 31.0% lower | Isolated retained heap, managed peak, process peak, package bytes, and structure counts are recorded | No accepted comparison; a commercial lane requires a valid license and the same complete VSDX contract | Large creation still allocates about 46.9 MiB for a 3.35 MB package and large loading about 63.8 MiB; stabilize elapsed evidence, reduce XML/preservation/package overhead, and add non-Windows evidence |

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
with worst observed margins of 1.36× elapsed on Windows and 1.38× allocation on
Linux. Three Windows lanes allocate less than ShapeCrawler; large create/save
uses 0.62× its elapsed time and 0.25× its managed allocation.

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

## Word Open XML SDK evidence

`OfficeIMO.Word.Benchmarks` has a publishable lane containing only OfficeIMO and
the MIT-licensed Open XML SDK. It validates equivalent rich DOCX creation,
structured reports, complete reads, and rich load-replace-save workflows before
timing them. The clean Windows evidence is recorded in
[`officeimo.word-openxml-2026-08-24.md`](officeimo.word-openxml-2026-08-24.md).
All controlled and isolated time, allocation, managed-peak, and process-peak
ratios are within 2x. The former 100-paragraph breach is now 1.86x time and
1.41x allocation in the controlled run; these contender margins remain
improvement targets rather than an optimality claim.

## Reader baselines

- `officeimo.reader.foundation-2026-07-10.md`: first Reader-wide extraction, detection, transport, and parser/chunker baseline after the P0 foundation work.
- [`officeimo.reader-extraction-2026-08-24.md`](officeimo.reader-extraction-2026-08-24.md): current routing, Markdown projection, XML sibling scaling, and YAML single-parse improvements plus provenance-bound retained, peak-memory, and normalized-output-size evidence.

Reader benchmark code lives in `OfficeIMO.Reader.Benchmarks`.

## Google Workspace transport

`OfficeIMO.GoogleWorkspace.Benchmarks` measures the public shared transport with deterministic in-memory responses, excluding network latency but retaining OfficeIMO policy and response processing. The clean Windows evidence is recorded in [`officeimo.googleworkspace-transport-2026-08-24.md`](officeimo.googleworkspace-transport-2026-08-24.md). The 4 MiB declared-length lane fell from 14.00 MiB to 4.00 MiB allocated, while the formerly deficient unknown-length lane is now 1.30x declared time and allocation-equivalent with exact bytes and bounded-response behavior.

## Confluence managed sections

`OfficeIMO.Confluence.Benchmarks` measures deterministic managed-section replacement without network or authentication noise. The current Windows evidence is recorded in [`officeimo.confluence-managed-section-2026-08-24.md`](officeimo.confluence-managed-section-2026-08-24.md). The 1 MiB lane fell from 6.19 MiB to 2.13 MiB allocated while preserving the exact updated body and both UTF-8 SHA-256 hashes.

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

## Markup semantic parsing

`OfficeIMO.Markup.Benchmarks` compares the CommonMark-compatible semantic
projection with Markdig using CommonMark plus pipe tables in both lanes. The
preflight requires identical semantic event counts and SHA-256 digests before
either parser is timed.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- --filter '*OfficeMarkupParseBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- --evidence --verify-budgets --repeat 3 --json .benchmark-artifacts\markup\evidence.json
```

The clean Windows evidence is recorded in
[`officeimo.markup-parse-2026-08-24.md`](officeimo.markup-parse-2026-08-24.md).
All three scales are within 1.31x for BenchmarkDotNet mean time and within 1.26x
for managed allocation. The isolated retained-heap and peak-memory ratios are
also within 1.31x. Markdig stays in the opt-in benchmark project and does not
enter normal restore, runtime, or packaging paths.

## Security detached CMS

`OfficeIMO.Security.Benchmarks` compares RSA-2048 SHA-256 detached CMS signing
and signature-only verification with .NET `SignedCms`. Both lanes use the same
content and certificate and materialize equivalent signer metadata, standard
signed-attribute inspection, certificate bytes, and key-usage policy. Preflight
cross-verifies both signatures and requires tamper rejection before timing.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- --filter '*SecurityCms*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\security\evidence.json
```

The clean Windows evidence is recorded in
[`officeimo.security-cms-2026-08-24.md`](officeimo.security-cms-2026-08-24.md).
The former 24.6-171.5x verification-time incident is removed. Every elapsed
and allocation lane is now below the 2x contender ceiling, including retained
heap and peak-memory evidence from isolated processes. The narrowest remaining
margins are 1.87x controlled time for the minimal platform-produced CMS and
1.79x controlled allocation for the minimal OfficeIMO-produced CMS. On .NET 8
and later `System.Security.Cryptography.Pkcs` owns the narrow attribute-free
platform fast path; Bouncy Castle remains the complete fallback and the only
path on `netstandard2.0` and .NET Framework.

## Provenance structural carriers

`OfficeIMO.Provenance.Benchmarks` measures bounded structural carrier
inspection and selective removal across deterministic PNG, TIFF, SVG, ZIP, and
structured-text fixtures. Preflight requires exactly one structurally valid
C2PA carrier, one removal, no carrier afterward, and the exact expected output
size.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- --filter '*ProvenanceBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\provenance\evidence.json
```

The clean Windows evidence is recorded in
[`officeimo.provenance-2026-08-24.md`](officeimo.provenance-2026-08-24.md).
The 1 MiB PNG lane improved from 99.78 ms to 6.22 ms for inspection and from
178.94 ms to 13.05 ms for removal; removal allocation fell from 1,046.05 KiB
to 21.54 KiB. No contender ratio is claimed because no accepted managed .NET
implementation exposes the same bounded structural contract. Large SVG, text,
and ZIP allocation remains explicit remediation rather than being described as
settled.

## Visio package workflows

`OfficeIMO.Visio.Benchmarks` measures complete in-memory VSDX creation/save and
load/structural-inspection over deterministic multi-page shape and connector
graphs. Preflight reopens every generated package and checks pages, shapes,
connectors, Shape Data, boundary text, and output size.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- --filter '*VisioBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\visio\evidence.json
```

The clean Windows evidence is recorded in
[`officeimo.visio-2026-08-24.md`](officeimo.visio-2026-08-24.md). Page-local ID
indexing, allocation-free tree traversal, and direct package streaming reduced
creation allocation by 16.0% to 64.4% across the three scales; large
load/inspection allocation is 31.0% lower. The timing sample is deliberately
not presented as a general speedup because the short runs remain noisy. No
contender ratio is claimed without a licensed, contract-equivalent managed
implementation.

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

## MHTML archive comparisons

`OfficeIMO.Mhtml.Benchmarks.Comparisons` measures complete MHTML loading and
serialization. The comparison lane uses MimeKit for multipart/related MIME and
AngleSharp for the HTML DOM, then retains every decoded resource. Both outputs
must pass both readers with equal root metadata, HTML, element count, ordered
resource metadata, decoded lengths, and payload hashes.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- --filter '*Mhtml*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- evidence --repeat 3 --json .benchmark-artifacts\mhtml\evidence.json
```

The clean Windows evidence is recorded in
[`officeimo.mhtml-2026-08-24.md`](officeimo.mhtml-2026-08-24.md). All time,
allocation, retained-heap, and peak-memory lanes are within 2x. Large-read
allocation fell 40.2% from the initial source, and OfficeIMO output is
0.864-0.973x MimeKit's size. Large-write allocation and managed peak remain the
weakest contender margins at 1.40x and approximately 1.46x, respectively.

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
dotnet run -c Release -f net10.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- evidence --repeat 3 --json .benchmark-artifacts\opendocument\ods-comparison-evidence.json
```

The comparison project is outside `OfficeIMO.sln`; OpenStandardLibrary remains
an opt-in benchmark dependency. It is not used as an ODT or ODP comparator
because it does not expose equivalent document models for those formats.
The clean Windows time, allocation, peak-memory, and size record is documented
in [`officeimo.opendocument-ods-2026-08-24.md`](officeimo.opendocument-ods-2026-08-24.md).

## Atlas Document Format JSON

`OfficeIMO.Adf.Benchmarks` measures complete typed ADF JSON parsing and
semantic parse/write round trips. The primary platform floor materializes a
typed document, node, mark, attribute, and extension-data graph; a second raw
JSON-tree lane remains visible but is not treated as equivalent work.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Adf.Benchmarks -- --filter '*' --job short --artifacts .benchmark-artifacts\adf\bdn
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Adf.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\adf\isolated.json
```

The clean Windows record is documented in
[`officeimo.adf-2026-08-24.md`](officeimo.adf-2026-08-24.md). Every primary
elapsed, allocation, retained-heap, managed-peak, and process-peak ratio is
within the 2x contender margin, and both round-trip scales preserve the exact
input byte count.

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

## ZIP traversal policy diagnostic

`OfficeIMO.Zip.Benchmarks` measures safe, deterministic OfficeIMO traversal
against direct `System.IO.Compression` metadata projection. Both lanes return
the same validated ordered descriptor fields; only OfficeIMO applies path and
expansion safety limits, so this non-parity diagnostic reports the cost of that
policy and is excluded from the published library-comparison catalog.

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

`html-pdf-artifact-evidence/` contains the compact Windows and Linux High-scale
artifact summaries for the identical-HTML lane. Each summary binds the four
engines, exact source commits, input bytes, correctness and accessibility
contracts, cancellation support, process-tree memory, output sizes, and
managed plus external visual hashes to a validated artifact manifest. The raw
PDF and PNG bundles remain temporary because their hashes and measurements are
the reproducible evidence contract.

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
