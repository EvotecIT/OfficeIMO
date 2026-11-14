# Library comparison evidence

Run the comparison runner once on each machine, copy its portable bundle, then
merge the bundle indexes. PowerForge validates result artifacts and keeps each
platform, run mode, source revision and measurement profile separate. A subset
merge retains the other lanes already in the destination catalog.

## First run

- Use PowerShell 7.4 or later, Git, the repository's .NET 10 SDK, and the .NET 8
  SDK/runtime used by the pinned PowerForge module. Package restore and the first
  PowerForge source checkout need network access. The initializer reads
  `Build/benchmark-powerforge.json`, caches source under `EVOTEC_SCRATCH_ROOT` when configured (otherwise local application data), builds the unsigned net8.0 module, and imports
  it. `-PowerForgeRoot` selects an existing clean checkout at that exact commit;
  it does not change the pin or update a shared checkout.
- Clear `OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS` before collecting the fixed seven-corpus CommonMark matrix.
- Use a clean OfficeIMO checkout for cataloged measurements. If selecting a
  local HtmlTinkerX source through `-HtmlTinkerXRoot`, keep it clean too. Record
  the actual source, dependencies and environment; copying a bundle does not
  turn one machine's results into another platform's evidence.
- Reserve the machine through the normal resource procedure before measurements.
  Keep equivalent engines on the same processor domain and settings. Leave
  `-AffinityMask 0` on macOS. A nonzero mask on Windows or Linux identifies a
  separate comparison profile; select it from the host's actual topology.
- Review the exact dependency licenses before running restricted lanes. Public
  QuestPDF comparisons pin 2026.5.0, whose embedded license permits the public
  lane; this is a historical version. Word's DocX comparison cannot be published
  without Xceed's advance permission. Word also requires explicit
  `-AcceptNPOIOSMFLicense` acknowledgement. `all -Publish` reports restricted
  optional lanes as skipped and leaves their declarations visible.
- Prepare the independent PDF corpus and pass `-PdfCorpusRoot` to include
  `pdfcorpusread`. Without it, `all` reports that lane as skipped. An explicit
  selection with a missing prerequisite fails. Review `BlockedReasons` in the
  plan before a first full run.

Run these commands from the repository root. They perform full
measurements; add `-DryRun` to inspect the plan without setup, restore, builds,
measurement, catalog writes or output-directory creation. `-PlanOnly` also
returns the plan, including declarations for restricted workloads.

| Machine | Full public-eligible collection command |
| --- | --- |
| Windows | `pwsh -NoLogo -NoProfile -File ./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -RunMode full -Publish -EvidenceRoot ./artifacts/benchmark-bundles/windows` |
| Linux | `pwsh -NoLogo -NoProfile -File ./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -RunMode full -Publish -EvidenceRoot ./artifacts/benchmark-bundles/linux` |
| macOS | `pwsh -NoLogo -NoProfile -File ./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -RunMode full -Publish -EvidenceRoot ./artifacts/benchmark-bundles/macos` |

These commands use the actual host platform and default to `net10.0`.
`-Framework net8.0` collects a separate cohort for comparisons supporting that framework. `-Platform windows`,
`linux` or `macos` is a planning override only, used with `-PlanOnly` or
`-DryRun`; it cannot relabel a real collection.

## Select a subset

Run subset and merge examples below inside PowerShell so comma-separated values
bind as arrays. `-Workload` accepts one or more workload IDs, `all`, and the grouped selectors
`word` and `powerpoint`. The plan is the source of supported IDs, areas, titles,
libraries, prerequisites and publication restrictions:

```powershell
./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload all -RunMode full -PlanOnly |
    Select-Object Workload, Area, ComparisonId, Libraries, CanRun, BlockedReasons

./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload csv,xlsx,pptxcreate -RunMode full -Publish -EvidenceRoot ./artifacts/benchmark-bundles/subset -DryRun
```

Remove `-DryRun` when ready to collect. `quick` runs are diagnostics and cannot
be marked publishable. `-MeasurementProfile WarmedV1` is limited to
`-Workload pdfhtmlpayload -RunMode full` and produces its own comparison ID.
Internal QuestPDF experiments require the runner's explicit authorization and
version flags; their artifacts cannot enter public bundles or website catalogs.

## Copy and merge bundles

Keep collection output under the ignored `artifacts/benchmark-bundles` directory
or outside the checkout. An unignored bundle dirties the source between
provenance captures. Copy each complete `EvidenceRoot`, including `index.json` and every referenced
normalized JSON artifact. Keep relative directories intact. Do not copy an index
alone or move the declaration array into a measured provenance record.

After putting the bundles under `./artifacts/benchmark-bundles` on the publishing checkout:

```powershell
./Build/Merge-LibraryComparisonBenchmarks.ps1 -SourcePath ./artifacts/benchmark-bundles/windows/index.json,./artifacts/benchmark-bundles/linux/index.json,./artifacts/benchmark-bundles/macos/index.json
```

Add `-DryRun` to inspect inputs, declarations, output paths and the PowerForge pin
without writes or setup. Merge runs no measurements. The destination defaults to
`Website/static/data/benchmarks/library-comparisons/index.json` and is always
included as an input when it exists. For a later subset, supply only that bundle's
index: other destination lanes remain. PowerForge selects lanes and verifies
hashes, provenance and compatibility; it does not average unlike environments.
The wrapper sets `-ResultPathPrefix /data/benchmarks/library-comparisons`, so
portable bundle filenames become website result links without changing payloads.

The wrapper obtains declarations from `all -RunMode full -PlanOnly` for both
`net8.0` and `net10.0`, omitting unsupported framework combinations. It maps historical affinity and warmed variants to their
explicit base declarations, preserving separate IDs. An unknown comparison ID
fails rather than guessing an area. With no sources and no destination it creates
a declared-only catalog, so never-run comparisons still have platform coverage.
Repeated merges of the same inputs produce the same catalog and partials.

## Generated website output

The merge regenerates these partials from the resulting catalog:

- `Website/themes/officeimo/partials/generated/benchmarks-catalog.html`: all
  declared comparisons, grouped by area, with libraries and per-platform evidence.
- `Website/themes/officeimo/partials/generated/benchmarks-overview.html`: current
  catalog coverage, separating full evidence from diagnostic lanes.

Regenerate them without merging with
`./Build/Generate-LibraryComparisonWebsiteData.ps1`. Use `-CatalogPath` and
`-OutputDirectory` for an isolated preview, or `-DryRun` to validate and show the
output plan. The website pipeline's `Generate-BenchmarkWebsiteOverview.ps1`
entry point delegates to this same generator. Edit the generator, never the
generated partials. The manual page
shell chooses which partials to include. The legacy Excel matrix and
`Website/data/benchmarks-excel*.json` are separate archived evidence and are not
rewritten by this pipeline.

Full published evidence, diagnostic evidence, incompatible full lanes, not-run
platforms and platforms outside a declaration's target set remain distinguishable.
Missing historical environment values stay missing; the generator labels them
as unrecorded rather than substituting the publishing host's environment.

## Catalog contract and checks

`Build/library-comparisons.schema.json` uses JSON Schema draft 7, supported by
PowerShell's `Test-Json`. The exact v4 top-level fields are `schemaVersion`,
`expectedPlatforms`, `areas` (string IDs), `comparisons`, `entries` and
`availability`. Each comparison declares `comparisonId`, `area`, `title`,
`libraries` (strings) and `expectedPlatforms` (strings). Each evidence entry keeps
the v3 lane, artifact, hash, environment and compatibility fields and adds `area`.
Availability retains `comparisonId`, `runMode`, `platform`, `available`, `runModes`
and nullable `latestUtc`. Additional artifact fields are allowed.

```powershell
Get-Content ./Website/static/data/benchmarks/library-comparisons/index.json -Raw |
    Test-Json -SchemaFile ./Build/library-comparisons.schema.json
```

The generator validates v4 before writing. PowerForge accepts existing v3 input
bundles for migration and emits v4 through the merge. Schema validation checks
types and required fields; PowerForge owns artifact integrity, lane selection
and compatibility. Plans and existing-artifact merges verify automation without
executing benchmarks. They do not establish new performance results or complete
the first real cross-platform collection.
