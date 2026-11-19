# Library comparison evidence

OfficeIMO measures its libraries against other .NET libraries with one runner,
records the results in one evidence catalog, and shows them on the website and
in the README result blocks from that catalog only. The catalog follows the
[PowerForge benchmark standard](https://github.com/EvotecIT/PSPublishModule/blob/main/Docs/PowerForge.BenchmarkStandard.md)
(catalog schema 5).

| File | Role |
| --- | --- |
| `Build/LibraryComparisonDefinitions.ps1` | Workloads: project, filter, libraries and the cases each workload runs. Further workloads are declared per area in `Build/LibraryComparisonDefinitions/<area>.ps1`. |
| `Build/benchmark-libraries.json` | Library policy: package, pinned version, licence and restriction class of every compared library. |
| `Build/benchmark-declarations.json` | Generated declarations the catalog checks every result against. Never edit it. |
| `Build/Run-LibraryComparisonBenchmarks.ps1` | Runs workloads and records each comparison in a catalog as soon as it is measured. |
| `Build/Merge-LibraryComparisonBenchmarks.ps1` | Merges catalogs from several machines into the website catalog. |
| `Website/static/data/benchmarks/evidence/index.json` | The website catalog. Its results are published beside it. |

## Declarations

`./Build/Export-BenchmarkDeclarations.ps1` writes `Build/benchmark-declarations.json`
from the workload definitions and the library policy; `-Check` fails when the
file is out of date, and the runner contract test runs it. Each comparison
declares:

- its libraries, with OfficeIMO as the subject;
- its cases and axes, taken from the workload's expected cases. A method name
  starts with its library (`MagickNETEncode` is Magick.NET's `Encode` case;
  `Sylvan` is Sylvan's default case). Methods that do not follow that pattern
  are listed in the workload's `MethodIdentity` or in `Build/BenchmarkDeclarations.ps1`. A method that names no
  declared library fails, so no result is ever attributed by guessing;
- its frameworks: only those on which every library runs its declared version
  with the same cases. Comparisons that include ExcelReader.NET are declared for
  net10.0 only, because the net8.0 build uses an older ExcelReader.NET;
- its eligibility. Workloads not meant as public comparisons, and comparisons
  of OfficeIMO alone, are recorded but never published.

Library ids are the library names readers see, such as `ClosedXML` or
`Open XML SDK`. When a policy library lists several packages or versions,
its `declare` entry names the package and pinned version the comparisons
measure.

## Licences

Every library has a restriction class in `Build/benchmark-libraries.json`:

| Class | Runs | Published |
| --- | --- | --- |
| `open` | Yes | Yes |
| `publish-restricted` | Locally, after `-IncludeLibrary <id>` | No |
| `acceptance-required` | Locally, after `-AcceptLibraryLicense <id>` and the library's own setting, such as `EPPlusLicense` | No |
| `excluded` | Never | Never |

NPOI, the Recursive AI One libraries, EPPlus 4.5.3.3 (LGPL) and QuestPDF
2026.5.0 are open. Later QuestPDF releases are publish-restricted. DocX and
EPPlus 8 need an explicit licence acceptance and are never published. The
runner never accepts a licence for you, and enabling any restricted library
keeps the whole run local, so `-Publish` rejects it. A published comparison
leaves out restricted libraries' cells and publishes the rest; the website shows
them as "not published (licence)". `-ExcludeLibrary <id>` skips comparisons that
need that library.

## Profiles and publishing

`-Profile` selects how much each workload measures:

| Profile | BenchmarkDotNet job | Published |
| --- | --- | --- |
| `diagnostic` | One dry pass | Never |
| `standard` | 1 launch, 5 warmups, 15 measured iterations, outliers kept | When the run meets the publish policy |
| `full` | The workload's own job | When the run meets the publish policy |

The catalog carries the standard publish policy: at least 3 warmups and 10
measured samples per cell, a 95% confidence interval of the median, passed
output validation, a clean source commit, the operating system, CPU, runtime and
SDK, no machine or user name, and the version and SHA-256 of every compared
package. The runner restores each project first and records the restored
version and hash of every competitor package before measuring.

Benchmarks record their output validation by appending one JSON object per
library and case (`{ "library": "ClosedXML", "case": "Default", "status": "passed", "detail": "rows=65000" }`)
to the file named by `OFFICEIMO_BENCHMARK_VALIDATION_PATH`, through
`Benchmarks/BenchmarkOutputValidationLog.cs`. Each lane's output is checked
before timing by reading it back and comparing its content with the expected
data; a cell without a passed validation is recorded but not published.

Benchmark methods name their library and case with
`[BenchmarkCategory("library:<id>", "case:<name>")]`, using the declared library
id. The importer takes a row's identity from these categories first, so
`./Build/Test-BenchmarkCategories.ps1` (run by the runner contract test) fails
when a category disagrees with what the declarations derive. A `library:`
category of a restricted library also keeps its cases out of runs that do not
enable it.

Results that drift from the declarations (a case, library or axis value that is
not declared, a declared combination that did not run, or a combination outside
the comparison's declared cells) are recorded unpublished and the runner reports
the drift. When not every library defines every case and axis value, the
generated declaration lists the comparison's cells explicitly. A crashed or failed case is recorded as a failed
cell; the rest of the comparison is still published when it meets the policy.

## First run

- Use PowerShell 7.4 or later, Git, the repository's .NET 10 SDK, and the .NET 8
  SDK used by the pinned PowerForge module. The runner reads
  `Build/benchmark-powerforge.json`, caches the PowerForge source under
  `EVOTEC_SCRATCH_ROOT` when it is set (otherwise local application data),
  builds the module and imports it. `-PowerForgeRoot` selects an existing clean
  checkout at that exact commit.
- Use a clean OfficeIMO checkout. A dirty tree is rejected for cataloged runs.
- Clear `OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS` before collecting the CommonMark
  matrix, and pass `-PdfCorpusRoot` to include `pdfcorpusread`.
- Keep equivalent engines on the same processor domain. Leave `-AffinityMask 0`
  on macOS. A nonzero mask on Windows or Linux records a separate lane of the
  same comparison with the mask as a lane parameter.

Run these commands from the repository root. Add `-PlanOnly` or `-DryRun` to
inspect the plan without restoring, building or measuring.

| Machine | Publishable collection |
| --- | --- |
| Windows | `pwsh -NoLogo -NoProfile -File ./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -Profile standard -Publish -EvidenceRoot ./artifacts/benchmark-bundles/windows` |
| Linux | `pwsh -NoLogo -NoProfile -File ./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -Profile standard -Publish -EvidenceRoot ./artifacts/benchmark-bundles/linux` |
| macOS | `pwsh -NoLogo -NoProfile -File ./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -Profile standard -Publish -EvidenceRoot ./artifacts/benchmark-bundles/macos` |

Use `-Profile full` for complete BenchmarkDotNet runs. The runner uses the host
platform and `net10.0` by default; `-Framework net8.0` records separate lanes for
the comparisons declared for it. `-Platform` is a planning override only.

## Select a subset

`-Workload` takes workload ids, `all`, and area names such as `word`, `powerpoint` or `excel` (a workload id wins over an area of the same name, as with `csv`):

```powershell
./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -Profile full -PlanOnly |
    Select-Object Workload, Area, ComparisonId, Libraries, DeclaredFrameworks, CanRun, BlockedReasons

./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload csv,xlsx,pptxcreate -Profile standard -Publish -EvidenceRoot ./artifacts/benchmark-bundles/subset -DryRun
```

`-MeasurementProfile WarmedV1` is limited to `-Workload pdfhtmlpayload -Profile full`
and records a separate lane (`measurement = warmedv1`) of the same comparison.
Internal QuestPDF experiments use `Build/Run-InternalQuestPdfBenchmarks.ps1`;
their results never enter a shared catalog.

## Merge bundles

Keep collection output under the ignored `artifacts/benchmark-bundles` folder or
outside the checkout. Copy each complete evidence root, including `index.json`
and every result beside it, to the publishing checkout and merge:

```powershell
./Build/Merge-LibraryComparisonBenchmarks.ps1 -SourcePath ./artifacts/benchmark-bundles/windows/index.json,./artifacts/benchmark-bundles/macos/index.json
```

The destination is always an input, so a later subset keeps the other lanes. A
newer run of a lane replaces it and the earlier run stays in the catalog's
superseded history with its date and hash; result files are content-addressed,
so none is overwritten. PowerForge verifies every result against its hash and
re-evaluates every lane against the declarations and the publish policy. With
no sources and no catalog, the merge writes a catalog of declarations only.
Catalogs from earlier schemas are rejected; record fresh runs instead.

After merging, refresh the CSV and Excel README blocks from the catalog with
`./Build/Benchmarks/Update-BenchmarkReadmes.ps1`.

## Website

`Website/site.json` reads the catalog through PowerForge.Web's `BenchmarkEvidence`
entry. The build writes one JSON file per area under `/data/benchmarks/` and one
rendered results fragment per area, and the pages use the PowerForge benchmark
components: `/benchmarks/` lists every area and loads an area's results on
demand, and `/benchmarks/<area>/` shows every comparison of that area. Only
published runs are shown; a comparison without one says it has not been
measured or published yet. Area titles and descriptions come from
`Website/data/benchmark_areas.json`. `Website/scripts/Test-BenchmarkPage.ps1`
checks the built pages against the catalog and runs the browser checks.
