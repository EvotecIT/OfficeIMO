# OfficeIMO.Word benchmarks

This opt-in project measures common Word document work with BenchmarkDotNet. It
compares the current `OfficeIMO.Word` source with:

- [DocX 5.2.0](https://www.nuget.org/packages/DocX/5.2.0), a high-level Word
  library available for open-source and non-commercial use under the
  [Xceed Community License](https://github.com/xceedsoftware/DocX/blob/master/CommunityLicense.txt).
- [NPOI 2.8.0](https://www.nuget.org/packages/NPOI/2.8.0), a high-level Word API
  whose source is Apache-2.0 licensed and whose official binary package is
  subject to the [NPOI binary EULA](https://github.com/nissl-lab/npoi/blob/master/OSMFEULA.txt).
- [Open XML SDK 3.5.1](https://github.com/dotnet/Open-XML-SDK), Microsoft's MIT-
  licensed low-level API for the document format.

These dependencies remain inside the benchmark project. They are not runtime
dependencies of `OfficeIMO.Word`.

## Comparison workloads

| Workload | Measured operation | Correctness proof |
| --- | --- | --- |
| Create paragraphs | Create a DOCX with 100 or 1,000 deterministic paragraphs and return its bytes | Exact paragraph count and text |
| Create structured report | Create a bold title, summary, header, footer, and a two-column table with 100 or 1,000 data rows | Exact structure, formatting signal, and every cell value |
| Read paragraphs | Load the same styled DOCX and traverse every body paragraph through each library's public API | Paragraph count, character count, and deterministic checksum |
| Replace and save | Load the same DOCX, replace every `{{Status}}` token, and serialize the result | Exact replacement count implied by every output paragraph and no remaining token |

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

The shared repository runner adds provenance capture and PowerForge-normalized
JSON/CSV/Markdown evidence. `word` expands to the four comparison workloads:

```powershell
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload word -RunMode full -Framework net8.0 -AcceptNPOIOSMFLicense
```

Use `-RunMode quick` only for a clean-source smoke run. The isolated benchmark
project records `AcceptNPOIOSMFLicense=true`, as required by the NPOI package;
it is outside the normal solution and CI restore/build path. The shared runner
also requires `-AcceptNPOIOSMFLicense` before starting a Word comparison. Review
the EULA and decide whether its maintenance-fee conditions apply to your use.
The benchmark project pins `System.Security.Cryptography.Xml` 8.0.4 because
NPOI's transitive 8.0.2 minimum is affected by a
[published high-severity advisory](https://github.com/advisories/GHSA-23rf-6693-g89p).

## Public harness, local numerical results

The benchmark source, deterministic inputs, validators, feature catalog, and
run instructions stay in this public repository. That makes the comparison
auditable and lets community members reproduce or improve it. Correctness
validation and the feature catalog are not numerical performance claims.

Numerical Word comparison artifacts stay local because the Xceed Community
License says Community Licensees may not publish DocX benchmark or performance
comparison results without Xceed's advance permission. The shared runner rejects
`-Publish` for every Word workload because each combined result contains a DocX
lane. NPOI's binary EULA adds fee and acknowledgement conditions but does not add
the same benchmark-publication restriction; NPOI is not the reason the combined
results remain local. If Xceed grants written permission, the publication gate
can be changed in a reviewed update. Raw BenchmarkDotNet artifacts belong in
ignored or temporary output folders.

## Feature applicability

The machine-readable [capability catalog](word-library-capabilities.json)
separates high-level support, low-level format access, adjacent conversion or
security packages, partial support, and features with no built-in workflow.
Render its current matrix with:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- features
```

OfficeIMO, DocX, and NPOI provide high-level DOCX authoring APIs. OfficeIMO also
has a diagnosed legacy DOC subset and adjacent HTML, Markdown, PDF, and security
packages, plus structured comparison/redline workflows. NPOI has a partial
legacy DOC surface but no equivalent built-in workflow for several higher-level
features. Open XML SDK exposes the underlying format rather than equivalent
productivity APIs, so the matrix labels those capabilities `low-level` instead
of treating format access as a high-level feature.

## OfficeIMO-only workflow diagnostics

`WordWorkflowBenchmarks` remains the focused OfficeIMO suite for field refresh,
mail merge, structured comparison, VBA signature inspection, package load, and
Word-to-HTML over loaded and one-shot documents:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*WordWorkflowBenchmarks*'
```

Those workflows have no directly equivalent public API in every comparison
library, so they stay outside the cross-library rankings.
