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

Each iteration also records the case, profile, encoder, page, byte length and SHA-256
of every output. Use these entries to locate an output-size failure and distinguish
content or layout changes from encoding changes. PDF entries describe the whole
document and include its page count; PNG and SVG entries identify the rendered page.
Their byte lengths sum to the iteration's output total. The report retains metadata,
without keeping another copy of the encoded output.

Use `--measure-only` to collect a new calibration result without enforcing the
checked-in platform ceiling. Evidence runs should use clean, commit-addressable source
and a new output directory.

The manual **Verify optional qualification suites** workflow provides the same
budget command on GitHub-hosted Windows, Linux and macOS runners. Select the
`html-static` suite and a workflow revision containing that lane. `source_ref`
selects the source to check out; an empty value uses the workflow revision.
Use a full 40-character commit SHA when comparing a control with a candidate.
The checkout action treats an abbreviated SHA as a branch or tag name.

Set `control_source_ref` to a full prior-source commit to run its unchanged budget
command after the candidate on the same runner. Both revisions enforce their
checked-in ceilings and retain separate reports, including failed measurements.
Use a control with the same corpus, worker and ceilings; compare the reports' source,
environment and output metadata before attributing a timing difference to code.
One ordered pair is diagnostic evidence, not a portable performance ranking.

This suite also runs static HTML and font-selection correctness on .NET 8 and .NET 10,
PDF correctness on .NET 10, the installed mathematical-font resource contract,
and the canonical packed-consumer smoke script.
The Windows package lane executes net472-targeted consumers on its installed .NET Framework runtime. Independent
package and measurement steps still run after a correctness failure, while the
failed job remains failed. Reports are retained as workflow artifacts for seven days.

Select `html-static-package-probe` for the same packed consumers on all three
platforms, including native .NET Framework execution on Windows, without repeating
correctness suites or budget measurements. The mathematical consumer uses a pinned
synthetic MATH font to exercise glyph construction and placement without relying on
an installed system font. This lane does not qualify timing or memory budgets.

The Linux lane installs Noto CJK, Arabic-capable Noto fonts and IPA Gothic for the
multilingual fixture. IPA Gothic supplies Japanese TrueType outlines accepted by
the installed multilingual PDF fallback; the Noto CJK package contains CFF
collections. These are validation-host assets. Output fingerprints from different
platforms or font environments need separate interpretation; compare candidate
and control within the same declared environment. Timing and memory budgets remain
opt-in evidence, outside ordinary pull-request correctness gates.
