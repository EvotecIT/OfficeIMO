# HTML editable evidence runner

This opt-in runner converts one frozen MHTML page from the [predeclared H10 corpus](../../OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/html-h10-page-selection.json) into its selected editable targets. It saves and reopens each artifact, checks the page's declared visible text markers, and records conversion reports, resource counts, elapsed time and allocations. A passing marker check does **not** establish visual or structural fidelity. Inspect each target against the `editableTargets.inspect` criteria and review every loss diagnostic before accepting it.

Each target report also records `NativeTables` from the reopened format model, separately from its HTML export. Word, PowerPoint, OneNote, RTF and Markdown entries contain native table rows and cells; Excel entries contain worksheet grids and do not imply named Excel tables. Compare these cells with the frozen source, including column ownership and fragmented tables. The observations do not change the marker-based `Passed` field into structural acceptance. Table snapshots and Excel reads are bounded to 100,000 cells per table or worksheet.

Build the runner from the repository's supported .NET SDK, then run its DLL from the repository root:

```sh
dotnet build Build/HtmlEditableEvidence/OfficeIMO.Html.EditableEvidence.csproj -c Release
dotnet Build/HtmlEditableEvidence/bin/Release/net10.0/OfficeIMO.Html.EditableEvidence.dll \
  --case epa-drinking-water-contaminants \
  --mhtml Ignore/HtmlUnknownPageQualification/h10-epa-ch-final-711d4e73c/source.mhtml \
  --output Ignore/HtmlUnknownPageQualification/epa-editable-evidence \
  --target all \
  --expected-sha256 12ac44c94520a28c5797c1c764df47ecb833db932e2772be868a25e9b7062dc3 \
  --require-clean-source
```

Use a new output directory for each run. The runner retains converted documents only in ignored local evidence; it writes `summary.json` and one `report.json` plus a reopened text projection per target. For inspection, reopened OneNote images use bounded files in that target's `assets/` folder, while RTF images are embedded in its local HTML projection; both have a 16 MiB per-image limit. This does not change their default public HTML export profiles. The summary includes the case role, archive hash, Git head, worktree cleanliness and loaded owner versions. `--require-clean-source` also verifies that every loaded owner assembly was rebuilt from that head. `--max-css-rules` can raise the untrusted profile's default 10,000-rule cap to at most 40,000 for a separately documented exploratory replay. Reported time and allocation are observations, not accepted cross-platform budgets.
