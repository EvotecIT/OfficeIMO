# Synthetic HTML report fixtures

This opt-in producer uses HtmlForgeX typed components to create two complete, script-free reports in portrait and landscape. The reports cover unequal parallel table flows in flex and grid regions, real table headers and footers, monitoring tables, alert queues, following content, internal links, and separate closed/open native `details` cases.

The predeclared [contract](synthetic-report-contract.json) owns the exact content inventories. These new synthetic reports do not establish acceptance of the unavailable historical report corpus. The producer stays outside the normal solution and performs no rendering or qualification.

Use the .NET SDK pinned by the repository and the HtmlForgeX source revision in [source-binding.json](producer/source-binding.json). Set `EVOTEC_GITHUB_ROOT` and `EVOTEC_SCRATCH_ROOT` to your existing repository and mounted scratch roots. An explicit `HtmlForgeXSourceRoot` MSBuild property overrides the project default of `EVOTEC_GITHUB_ROOT/HtmlForgeX`.

Supply the unmodified Carlito regular and bold fonts from `ChartForgeX.Examples/Fixtures/Fonts/Carlito` in the ChartForgeX checkout. Copy them as shown below; their exact hashes, upstream source, and SIL Open Font License 1.1 binding are in the manifest. Fonts and generated reports remain in scratch storage.

Run from the OfficeIMO repository root:

```sh
: "${EVOTEC_GITHUB_ROOT:?Set the existing repository root}"
: "${EVOTEC_SCRATCH_ROOT:?Set the mounted scratch root}"
fixtureSource="$PWD/Build/HtmlReportFixtures"
fixtureScratch="$EVOTEC_SCRATCH_ROOT/html-report-fixtures"
fontSource="$EVOTEC_GITHUB_ROOT/ChartForgeX/ChartForgeX.Examples/Fixtures/Fonts/Carlito"
hfxSource="$EVOTEC_GITHUB_ROOT/HtmlForgeX"
mkdir -p "$fixtureScratch/inputs/report-fonts"
cp "$fontSource/Carlito-Regular.ttf" "$fixtureScratch/inputs/report-fonts/regular.ttf"
cp "$fontSource/Carlito-Bold.ttf" "$fixtureScratch/inputs/report-fonts/bold.ttf"
cp "$fontSource/OFL.txt" "$fixtureScratch/inputs/report-fonts/OFL.txt"

DOTNET_PROCESSOR_COUNT=3 dotnet build "$fixtureSource/producer/SyntheticReportProducer.csproj" -c Release -f net10.0 -m:3 --artifacts-path "$fixtureScratch/build" -p:HtmlForgeXSourceRoot="$hfxSource" -p:HfxMinifyAssets=false
DOTNET_PROCESSOR_COUNT=3 dotnet run --project "$fixtureSource/producer/SyntheticReportProducer.csproj" -c Release -f net10.0 --no-build --no-restore --artifacts-path "$fixtureScratch/build" -p:HtmlForgeXSourceRoot="$hfxSource" -- --contract "$fixtureSource/synthetic-report-contract.json" --inputs "$fixtureScratch/inputs" --output "$fixtureScratch/reports"
```

The output names are `parallel-tables-portrait.html`, `parallel-tables-landscape.html`, `grid-queues-portrait.html`, `grid-queues-landscape.html`, and `producer-receipt.json`. The producer verifies the supplied font hashes and records the input and output hashes. The manifest preserves the expected HTML hashes, the original qualification contract byte hash, and a separate canonical content hash for the compact portable contract.

Only the core `Table` components emit repeating `tfoot` groups. Monitoring table summaries and queue summaries are terminal content. Queue rows have unique visible labels because the public queue API has no row-ID hook. Use the declared inventories and an independent renderer to assess the generated reports; producer execution alone does not qualify their layout.
