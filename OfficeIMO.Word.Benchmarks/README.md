# OfficeIMO.Word benchmarks

This opt-in project measures common Word document work with BenchmarkDotNet. It
compares the current `OfficeIMO.Word` source with:

- [DocX 5.2.0](https://www.nuget.org/packages/DocX/5.2.0), a high-level Word
  library available for open-source and non-commercial use under the
  [Xceed Community License](https://github.com/xceedsoftware/DocX/blob/master/CommunityLicense.txt).
- [NPOI 2.8.1](https://www.nuget.org/packages/NPOI/2.8.1), a high-level Word API
  whose source is Apache-2.0 licensed and whose official binary package is
  subject to the [NPOI binary EULA](https://github.com/nissl-lab/npoi/blob/master/OSMFEULA.txt).
- [Open XML SDK 3.5.1](https://github.com/dotnet/Open-XML-SDK), Microsoft's MIT-
  licensed low-level API for the document format.

- [RecursiveAI.OneDocument.Documents 1.0.0](https://www.nuget.org/packages/RecursiveAI.OneDocument.Documents/1.0.0),
  a managed Word library licensed `MIT OR Apache-2.0`. It targets `net10.0`
  only, so its lanes exist only when the runner is built for `net10.0`.

These dependencies remain inside the benchmark project. They are not runtime
dependencies of `OfficeIMO.Word`.

### OneDocument lanes

On `net10.0` the read, replace and structured-report comparisons gain a
`OneDocument` lane, and read gains `OneDocumentExtractText`:

| Lane | Measured operation | Correctness proof |
| --- | --- | --- |
| Read paragraphs, `OneDocument` | `Document.Open` and the text of every `Document.Paragraphs` entry | Same paragraph count, character count and checksum as every other reader |
| Read paragraphs, `OfficeIMOReaderText` (all targets) | `OfficeIMO.Reader` with `ReaderProjection.TextAndMetadata`, then `GetText()`, split into lines | Same observation; this is OfficeIMO's plain-text route, so it pairs with OneDocument's extractor |
| Read paragraphs, `OfficeIMOReaderStreamedText` (all targets) | The same Reader read with `ReaderWordOptions.StreamTextWhenPossible`, which streams simple bodies without building the Word model | Same observation |
| Read paragraphs, `OneDocumentExtractText` | `Document.Open` and `ExtractText()`, OneDocument's streaming extractor, split into lines | Same observation; it skips the document model, so it is the cheaper of the two OneDocument routes |
| Replace and save | `Document.Open`, `Document.Replace` for `{{Status}}`, `SaveAs` to a DOCX stream | Every output paragraph and Open XML SDK validation |
| Create structured report | `Document.Create`, header and footer stories, bold 18 pt title, summary, and a two-column table built from a cell array | The same structure, formatting, header/footer and Open XML SDK checks as the other lanes |

OneDocument has no paragraph-level API on header and footer stories, so the
report lane writes their text through the story's typed `w:hdr`/`w:ftr` view.
Its table uses OneDocument's default 4680-twip grid columns rather than the
2400-twip widths the Open XML SDK lane writes; the shared validator checks
content and formatting, not column widths.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*WordReadComparisonBenchmarks*' '*WordReplaceComparisonBenchmarks*' '*WordCreateReportComparisonBenchmarks*' --job Short
```

## Comparison workloads

| Workload | Measured operation | Correctness proof |
| --- | --- | --- |
| Create paragraphs | Create a DOCX with 100 or 1,000 deterministic paragraphs and return its bytes | Exact paragraph count and text |
| Create structured report | Create a bold title, summary, header, footer, and a two-column table with 100 or 1,000 data rows | Exact structure, formatting signal, and every cell value |
| Read paragraphs | Load the same styled DOCX and traverse every body paragraph through each library's public API | Paragraph count, character count, and deterministic checksum |
| Replace and save | Load the same DOCX, replace every `{{Status}}` token, and serialize the result | Exact replacement count implied by every output paragraph and no remaining token |
| Create with images | Create one caption and one inline PNG per item (10 or 100 distinct images) and return the bytes | Every caption, one embedded picture per item, and the exact source bytes of each image |
| Fill a template | Load a single-run template, replace ten distinct `{{Field}}` tokens, and save | Every filled paragraph |
| Read tables | Load a 100- or 1,000-row report and read every table cell through the library's table model | Cell count, character count, and deterministic checksum |

Paragraph creation and paragraph reading also run at 10,000 items as the
large-document case. Image payloads are distinct (each carries its own PNG text
chunk) so no library can deduplicate them. Libraries without a document-level
replace API, NPOI and the Open XML SDK, fill templates in one pass over runs or
text nodes; OfficeIMO and DocX call their replace API once per field. The table
fixture is the Open XML SDK report output with a minimal style catalog, which
DocX requires before it enumerates content.

Corpus construction and validation run in global setup, outside measured
operations. Creation and replacement include serialization because the useful
result is a DOCX payload. Read includes package load because that is the common
one-shot application workflow. The Open XML SDK lanes are format-native
baselines, not claims that low-level code has the same maintenance cost or
feature surface as a document library.

Every OfficeIMO, NPOI, and Open XML SDK output passes Open XML SDK validation.
The DocX 5.2.0 output is checked independently at ZIP/XML level for the
complete semantic payload. Its generated `/word/webSettings.xml` currently uses
`application/vnd.openxmlformats-officedocument.wordprocessingml.websettings+xml`;
Open XML SDK 3.5.1 expects `...webSettings+xml` and rejects the package before
schema validation. The benchmark keeps that observed compatibility caveat
visible instead of modifying DocX output inside the timed operation.

## Validate before measuring

Run every implementation once at both sizes and reject unequal output:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- validate
```

Run a dry BenchmarkDotNet smoke test:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*Word*ComparisonBenchmarks*' --job Dry
```

`Dry` proves execution only. Its single cold-start observation is not a
performance result. Run the normal job on an otherwise idle machine before
drawing conclusions:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*Word*ComparisonBenchmarks*'
```

### Shared custom-data snapshots

The opt-in PowerForge suite measures load, public legacy-write assessment and disposal with one
64 KiB payload shared by the package root, main document and 0, 32 or 128
additional parents. Setup and validation stay outside timing. Each measured sample
checks the assessment result; setup checks the payload bytes and shared relationships. Allocation
is recorded without a timing or memory threshold in correctness CI.

```powershell
dotnet build OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj -c Release -f net8.0
Import-Module PSPublishModule
Invoke-BenchmarkSuite -Path ./Build/Benchmarks/word-shared-clone.benchmark.ps1 -OutputRoot ./.benchmark-artifacts/word-shared-clone
```

### Publishable Open XML SDK evidence

The public evidence classes isolate OfficeIMO.Word and the MIT-licensed Open
XML SDK from the license-restricted combined comparison. Both implementations
start from the same rich package shell, perform equivalent work, and must pass
the structural validators before measurement:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- validate-openxml
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*Word*OpenXmlEvidenceBenchmarks*'
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\word\openxml-evidence.json
```

The isolated runner records source commit and dirty-tree state together with
elapsed time, allocation, retained heap, sampled managed peak, absolute process
peak, input bytes, and output bytes. The current Windows result is documented
in [`Docs/benchmarks/officeimo.word-openxml-2026-08-24.md`](../Docs/benchmarks/officeimo.word-openxml-2026-08-24.md).
The [2026-09-22 structured-report cell-writing run](../Docs/benchmarks/officeimo.word-table-report-2026-09-22.md)
records a three-platform allocation reduction and its remaining timing limits.
The [100-item format-depth budgets](../Docs/benchmarks/officeimo.word-format-depth-budgets-2026-09-22.md)
set absolute elapsed, allocation, managed-peak, process-peak, and DOCX-size
limits for paragraph creation, structured reports, and rich replacement. Run
`evidence --repeat 3 --budget .\OfficeIMO.Word.Benchmarks\word-format-depth-budgets.json`
after `validate-openxml` to check them. The process-peak reading is unavailable
on the current macOS host.

The shared repository runner adds provenance capture and PowerForge-normalized
JSON/CSV/Markdown evidence. Word workloads are excluded from the default `all`
selection so ordinary comparison runs do not restore NPOI or require its license
acknowledgement. `word` expands to the four comparison workloads:

```powershell
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload word -RunMode full -Framework net8.0 -AcceptNPOIOSMFLicense
```

Use `-RunMode quick` only for a clean-source smoke run. The isolated benchmark
project records `AcceptNPOIOSMFLicense=true`, as required by the NPOI package;
it is nonpackable and outside the normal solution. The interoperability CI job
restores and builds this project to run the deterministic `validate-workflows`
checks; comparison measurements remain opt-in. The shared runner also requires
`-AcceptNPOIOSMFLicense` before starting a Word comparison. Review the EULA and
decide whether its maintenance-fee conditions apply to your use.
The benchmark project pins `System.Security.Cryptography.Xml` 10.0.13 because
NPOI's transitive 10.0.10 minimum is affected by a
[published high-severity advisory](https://github.com/advisories/GHSA-7hqp-7c5m-4vq6).

## Public harness, local numerical results

The benchmark source, deterministic inputs, validators, and run instructions
stay in this public repository. That makes each measured workload auditable and
lets community members reproduce or improve it.

Numerical Word comparison artifacts stay local because the Xceed Community
License says Community Licensees may not publish DocX benchmark or performance
comparison results without Xceed's advance permission. The shared runner rejects
`-Publish` for every Word workload and never writes Word quick-run evidence into
the website catalog because each combined result contains a DocX lane. NPOI's
binary EULA adds fee and acknowledgement conditions but does not add the same
benchmark-publication restriction; NPOI is not the reason the combined results
remain local. If Xceed grants written permission, the publication gate can be
changed in a reviewed update. Raw BenchmarkDotNet artifacts belong in ignored
or temporary output folders.

## OfficeIMO-only workflow diagnostics

`WordWorkflowBenchmarks` remains the focused OfficeIMO suite for field refresh,
mail merge, structured comparison, VBA signature inspection, package load, and
Word-to-HTML over loaded and one-shot documents:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*WordWorkflowBenchmarks*'
```

Those workflows have no directly equivalent public API in every comparison
library, so they stay outside the cross-library rankings.
