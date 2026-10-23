# Pull-request validation

`select.py` follows tracked `ProjectReference` edges in reverse, from changed
projects to their consumers. `pr-suites.json` assigns executable contracts to
jobs and Ubuntu partitions. Conditional references are included for every host;
PowerShell's repository-root references resolve to the same graph. Shared source,
fixtures, build configuration and unknown paths select all managed contracts.
Linked source and resource inputs also select their consumers. A directly
changed test project without a declared executable route fails selection so a
new project cannot silently receive a passing aggregate without running tests.

Ordinary PRs run affected Windows .NET Framework contracts and the newest Linux
TFM. The smaller platform matrices retain Windows, Linux and macOS on the newest
TFM. The macOS smoke retains its full-run TFM and uses the newest TFM on PRs; its PR build includes only the projects it executes. Full runs retain the original macOS solution build. Website, JavaScript,
PowerShell and renderer comparisons use the same selection; their own assets
also select their owning suite. Grouped Ubuntu partitions restore, build and test
only affected projects inside the selected partition.

The ordinary library build uses maintained solution membership; opt-in
verification and comparison executables outside that solution keep their own
validation routes and dependencies.

Master pushes, nightly runs, manual dispatch and the `full-ci` PR label retain
the full managed matrices and managed/native AOT qualification. Adding or
removing that label reevaluates PR jobs. Full runs retain all native scenario
groups. The conversion comparison also runs on master and nightly.

Run the selector and its focused contracts locally:

```sh
python -m unittest discover -s Build/Ci -p 'test_*.py'
python Build/Ci/select.py --base origin/master
python Build/Ci/select.py --files OfficeIMO.Word/WordDocument.cs
python Build/Ci/select.py --full
```

The `PR validation` job always reports, accepts intentional skips, and fails on
failed or cancelled dependencies or failed selection. It aggregates the .NET
workflow; the separate website, PowerShell and conversion workflows retain their
own results. Configure required checks explicitly when adopting that aggregate;
the workflow does not alter repository protection or reviewer configuration.

Font data and ordinary NuGet restores use local composite actions. The NuGet
action resolves the configured global package directory; package-consumer
qualification keeps its isolated caches. Complete default-branch builds populate
the immutable NuGet key; small suites restore it without saving partial graphs.
PowerShell tests cache exact module versions. `Save-PowerShellTools.ps1` retrieves
the three dependency-free tools listed in `powershell-tools.json` directly from
the Gallery CDN and verifies their SHA-256 and module manifests. Both PowerShell
hosts and website preparation use that bootstrap, avoiding Gallery discovery
requests and NuGet provider setup. Recheck package dependencies when updating pins.
Chromium installation belongs to HtmlTinkerX;
the comparison tool delegates installation to that owner and caches its output.

Test declarations and format fixtures are retained. The redundant complete
LegacyImport invocation in the Excel partition is owned by `legacy-imports`;
distinct Excel legacy-reader tests remain in their partition. An obsolete
absence audit for the retired aggregate test project is removed. New routing
contracts cover dependency propagation, conditional references, shared inputs,
unmapped projects, docs/website/JavaScript changes and the full-matrix override.

Website PR selection uses `website_inputs` rather than the managed dependency
closure. Website assets, browser sources, PowerShell content, documentation,
site generators and build/dependency configuration select it. Ordinary library
implementation and unrelated CI edits retain their own correctness lanes. Full
site/API generation remains on master publication, scheduled publication,
manual dispatch and `full-ci` requests.

PSWriteOffice content sync reuses the checked-in snapshot only when its producer
commit has identical module, engine, shared build and snapshot-processing inputs. It also checks the
API component hashes and command inventory, documentation/catalog/navigation
fingerprints, and current working inputs. Shallow checkouts may fetch the
producer's tree. Missing provenance, changed inputs or damaged outputs use the
normal unsigned fresh export. Reused source links retain their actual producer
revision; they are not relabeled as a new build.

CodeQL retains its full manual production build and query suite.
`Invoke-CodeQLBuild.ps1` selects net10 for multi-targeted projects through
`OfficeIMOCodeQLBuild`; single-targeted projects keep their own framework
identity. This avoids duplicate instances writing to the same outputs. The
standalone generator retains netstandard2.0. Compiler analyzers, source generators
and binary logs remain enabled. `-OriginalFrameworkSelection` supplies the old
solution-wide framework selection for paired diagnostics. Compare both finalized databases with `CodeQL/ExtractedFiles.ql`
(CSV output), then run the same complete query suite to SARIF and compare:

```sh
python Build/Ci/compare_codeql.py before-files.csv after-files.csv before.sarif after.sarif expected-query-ids.json --query-metadata expected-query-metadata.json --before-query-results before.zip --after-query-results after.zip --output comparison.json
```

The comparison requires identical extracted source paths/line counts, query IDs
and alert locations/supporting locations/messages/flows. Keep product source, CLI/query versions,
output paths, cache policy and query options the same for both runs. Changed
project scans, cached compilation outputs and no-build extraction are not used.
Resolve `expected-query-ids.json` and query metadata (`id` and absolute `query`
path per entry) from the unchanged full C# query suite. Analyze each clean
database with that suite, then retain `codeql database bundle --include-results`
ZIPs. Each resolved query must have its BQRS result in both bundles; SARIF
descriptors alone do not prove execution. Empty, unsuccessful or equally
incomplete analysis evidence cannot pass comparison. SARIF invocations are
optional, but any recorded failure or error notification rejects the evidence.
