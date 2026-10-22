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
PowerShell tests cache exact module
versions and use PSResourceGet on PowerShell 7, with the supported PowerShellGet
route on Windows PowerShell 5.1. Chromium installation belongs to HtmlTinkerX;
the comparison tool delegates installation to that owner and caches its output.

Test declarations and format fixtures are retained. The redundant complete
LegacyImport invocation in the Excel partition is owned by `legacy-imports`;
distinct Excel legacy-reader tests remain in their partition. An obsolete
absence audit for the retired aggregate test project is removed. New routing
contracts cover dependency propagation, conditional references, shared inputs,
unmapped projects, docs/website/JavaScript changes and the full-matrix override.
