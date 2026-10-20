# H10 unfamiliar-page budget runner

This opt-in runner measures the PDF intents and editable targets declared for one frozen page in `OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/html-h10-page-selection.json`. It starts a fresh process for each operation. The frozen PDF worker uses the declared print, optional zero-margin print, screen-media and screen-snapshot profiles; the editable worker saves and reopens each target and checks its declared text markers. The report records conversion time and managed allocations inside the worker, plus peak process-tree working set across startup and conversion. A measured run is baseline evidence, not a fidelity pass or an accepted platform budget. Review the browser/reference appearance and every loss diagnostic separately.

Build all three tools in Release, then run the budget tool from the repository root against a frozen archive whose hash is known:

```sh
dotnet build OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj -c Release -f net10.0 -m:3
dotnet build Build/HtmlEditableEvidence/OfficeIMO.Html.EditableEvidence.csproj -c Release -m:3
dotnet build Build/HtmlH10Budget/OfficeIMO.Html.H10Budget.csproj -c Release -m:3
dotnet Build/HtmlH10Budget/bin/Release/net10.0/OfficeIMO.Html.H10Budget.dll \
  --case noaa-nos-chesapeake-bathymetry \
  --mhtml Ignore/HtmlUnknownPageQualification/h10-noaa-chesapeake-capture-800cf477a/source.mhtml \
  --output Ignore/HtmlUnknownPageQualification/noaa-chesapeake-budget \
  --expected-sha256 ecdac81c8d6253b79cde1648bc50026839997b40a036a8df0c3bfc7901d1292f \
  --require-clean-source
```

The PDF worker can also replay one intent directly from the same frozen archive:

```sh
dotnet OfficeIMO.Pdf.Benchmarks.Comparisons/bin/Release/net10.0/OfficeIMO.Pdf.Benchmarks.Comparisons.dll \
  html-mhtml-evidence --mhtml <frozen-source.mhtml> --output <new-directory> \
  --isolated-officeimo-intent print --require-clean-source
```

The other intent values are `print-zero-margin`, `screen-media` and `screen-snapshot`. This worker accepts frozen MHTML only. Native Edge reference capture stays with the HtmlTinkerX qualification owner and runs outside the measured worker. The worker records one schema-3 operation, the exact archive hash, conversion warnings and the saved PDF's page inventory. Its timer and allocation counter include PDF conversion and serialization after MHTML loading; they exclude archive parsing, file writing, saved-PDF inspection and native comparison. Source coverage and appearance need separate all-page acceptance.

Use a new output directory for each run. `h10-budget.json` records the source commit, manifest and archive hashes, environment, per-operation metrics and failures. The child reports and converted artifacts remain below that directory for inspection. A clean-source run also makes each child verify that its loaded OfficeIMO assemblies were built from the exact Git head.

After representative baselines have been collected on each supported platform, pass `--ceilings <json>` to enforce a case-specific budget. The JSON contract has `schemaVersion: 2`, `sourceSha256`, `caseSha256`, and `platforms`; each platform entry has `osFamily`, `caseId`, and an `operations` array. `caseSha256` hashes the selected page's normalized JSON declaration, so adding or changing a different corpus page does not invalidate this case's ceiling. The report still records the complete manifest hash for provenance. Every declared `pdf-print`, optional `pdf-print-zero-margin`, `pdf-screen-media`, `pdf-screen-snapshot`, and `editable-<target>` operation needs positive `conversionMilliseconds`, `managedAllocatedBytes`, and `peakWorkingSetBytes` ceilings. Missing operations, a source or selected-case mismatch, conversion failures, or exceeded ceilings fail the run. Keep baseline measurements and accepted ceilings distinct: do not set ceilings from a single quiet-machine sample or claim visual acceptance because this tool passes.
