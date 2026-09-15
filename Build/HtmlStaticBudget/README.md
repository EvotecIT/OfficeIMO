# HTML static-rendering budgets

This build tool measures the frozen H4/v2 corpus with OfficeIMO's owned static
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
