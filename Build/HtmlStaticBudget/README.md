# HTML static-rendering budgets

This build tool measures the frozen H4/advanced-held-out corpus with OfficeIMO's owned static
renderer. It has no comparison-engine dependencies, so the same command can run on
Windows, Linux, and macOS.

```powershell
$output = Join-Path $env:TEMP ("OfficeIMO-H4-budget-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
dotnet run --project Build/HtmlStaticBudget/OfficeIMO.Html.StaticBudget.csproj `
    -c Release `
    -- `
    --iterations 3 `
    --require-clean-source `
    --output $output
```

Each run creates isolated cold and warmed workers. Every measured iteration renders
all eight cases through screen PNG/SVG, print PDF/PNG/SVG, and screen-to-page
PDF/PNG/SVG. The JSON report records elapsed time, managed allocations, process-tree
peak working set, output bytes, deterministic fingerprints, source and corpus hashes,
and cancellation latency. The enforced gate requires every warmed fingerprint to
match the independently started cold worker, in addition to repeatability within the
warmed worker. Corpus documents enter OfficeIMO through their frozen source bytes so
legacy-encoding work remains part of the measured path.

Use `--measure-only` to collect a new calibration result without enforcing the
checked-in platform ceiling. Evidence runs should use clean, commit-addressable source
and a new output directory.

The manual **Verify optional qualification suites** workflow provides the same
budget command on GitHub-hosted Windows, Linux and macOS runners. Select the
`html-static` suite and a workflow revision containing that lane. `source_ref`
selects the source to check out; an empty value uses the workflow revision.
Use a commit SHA when comparing a control with a candidate.

This suite also runs static HTML correctness on .NET 8 and .NET 10, the installed
mathematical-font resource contract, and the canonical packed-consumer smoke script.
The Windows package lane executes the .NET Framework 4.7.2 consumers. Independent
package and measurement steps still run after a correctness failure, while the
failed job remains failed. Reports are retained as workflow artifacts for seven days.

The Linux lane installs Noto CJK and Arabic-capable fonts for the multilingual
fixture. These are validation-host assets. Output fingerprints from different
platforms or font environments need separate interpretation; compare candidate
and control within the same declared environment. Timing and memory budgets remain
opt-in evidence, outside ordinary pull-request correctness gates.
