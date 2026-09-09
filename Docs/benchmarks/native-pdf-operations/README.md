# Native PDF operations: 8 September 2026

[The measured data](2026-09-08-windows-net8.json) compares product commits `d3aa130ac52beb39cf382cef7987cc5b29d5bd2e` and `9e4001e83179d39636edb3b976cd6befd381e40a` on .NET 8.0.30. BenchmarkDotNet 0.15.8 ran its full default job in separate baseline/current processes on two CPU-affinity domains of one Windows workstation. The JSON retains iteration values, statistical summaries, managed allocation, host information, and hashes of the original reports.

## Equivalent work

Each independent producer creates a 100-page PDF with four table rows and one narrative paragraph per page. The benchmark includes OfficeIMO source parsing and, for page operations, serialization of every output. Setup reopens inputs and outputs with PdfPig and requires every expected page and narrative/table value. Selection emits 25 pages in descending source order. Merge combines 25 four-page documents. Both split methods emit the same 100 single-page documents.

## Observations

Timing ranges below span the two affinity runs. Allocation values are managed MiB per complete operation; they are not peak heap or process memory.

| Operation and input | Baseline mean | Current mean | Baseline allocation | Current allocation |
| --- | ---: | ---: | ---: | ---: |
| Explicit split selections, iText | 518–742 ms | 34–35 ms | 592.57 MiB | 68.05 MiB |
| Explicit split selections, MigraDoc | 995–1,117 ms | 48–51 ms | 1,032 MiB | 85.79 MiB |
| Page selection, iText | 11.6–15.0 ms | 9.2–10.9 ms | 12.20 MiB | 8.77 MiB |
| Page selection, MigraDoc | 21.1–23.1 ms | 9.8–10.5 ms | 18.57 MiB | 14.08 MiB |
| Whole-document split, iText | 37.3–37.5 ms | 34.0–35.2 ms | 70.56 MiB | 68.03 MiB |
| Whole-document split, MigraDoc | 53.7–59.5 ms | 55.9–59.0 ms | 89.83 MiB | 85.77 MiB |
| Plain-text reading, iText | 40.6–45.6 ms | 39.2–41.3 ms | 63.43 MiB | 61.36 MiB |
| Plain-text reading, MigraDoc | 51.4–57.5 ms | 45.8–60.0 ms | 72.03 MiB | 69.00 MiB |
| Structured reading, iText | 181.6–187.4 ms | 177.7–179.4 ms | 165.94 MiB | 161.11 MiB |
| Structured reading, MigraDoc | 215.7–226.6 ms | 203.3–229.0 ms | 215.39 MiB | 208.31 MiB |
| Merge, iText | 31.0–31.2 ms | 31.1–31.3 ms | 37.93 MiB | 38.17 MiB |
| Merge, MigraDoc | 34.0–35.8 ms | 33.0–38.5 ms | 52.84 MiB | 52.95 MiB |

Reusing one extraction session removes repeated whole-source parsing from compound split selections. That gain is large in both runs. Page selection also consistently allocates less and completes sooner. The reader's borrowed operand buffer reduces allocation by roughly 3–4%; timings for reading and ordinary splitting are less conclusive. Merge already parsed each source once, and now takes a fresh metadata snapshot to prevent edits to a returned read model from changing the source document's page-operation metadata.

Several short builds overlapped the first run. The second used a different affinity domain without concurrent full builds, but neither run represents a dedicated laboratory host. Small changes, overlapping distributions, and differences that change sign across runs do not establish a general speedup or regression. In particular, this evidence does not claim faster merge or uniformly faster structured reading.

## Reproduce

For a current-checkout measurement, run the committed benchmark through the shared evidence runner:

```powershell
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfnative -RunMode full -Framework net8.0 -AffinityMask 65535
```

To reconstruct both dated jobs, keep the benchmark source from this checkout and build the two product revisions separately. The [revision host](../../../Build/NativePdfRevisionBenchmarks/OfficeIMO.NativePdfRevisionBenchmarks.csproj) links the same benchmark class, producer generators, font, and validation helpers; neither historical product revision needs to contain the benchmark. It loads each revision's Core/Pdf assemblies in a separate BenchmarkDotNet process.

Run from the repository root with PowerShell 7 and the .NET 8 SDK/runtime available:

```powershell
$comparisonRoot = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-native-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $comparisonRoot | Out-Null
git worktree add --detach (Join-Path $comparisonRoot 'baseline') d3aa130ac52beb39cf382cef7987cc5b29d5bd2e
git worktree add --detach (Join-Path $comparisonRoot 'current') 9e4001e83179d39636edb3b976cd6befd381e40a
foreach ($revision in 'baseline', 'current') {
    dotnet build (Join-Path $comparisonRoot "$revision/OfficeIMO.Pdf/OfficeIMO.Pdf.csproj") -c Release -f net8.0
    if ($LASTEXITCODE -ne 0) { throw "Product build failed: $revision" }
}
$env:NativePdfBaselineDirectory = Join-Path $comparisonRoot 'baseline/OfficeIMO.Pdf/bin/Release/net8.0'
$env:NativePdfCurrentDirectory = Join-Path $comparisonRoot 'current/OfficeIMO.Pdf/bin/Release/net8.0'
$env:PDF_BENCHMARK_RUN = 'full'
foreach ($mask in '65535', '4294901760') {
    $env:PDF_BENCHMARK_AFFINITY = $mask
    dotnet run --project Build/NativePdfRevisionBenchmarks -c Release -- --filter '*' --artifacts (Join-Path $comparisonRoot "results-$mask")
    if ($LASTEXITCODE -ne 0) { throw "Benchmark host failed: $mask" }
}
```

These affinity masks describe the recorded Windows workstation. Choose valid masks for the host being measured and compare both revisions on the same mask. `PDF_BENCHMARK_RUN=dry` executes all 24 producer/operation/revision combinations with setup validation for a smoke check; it does not produce publishable timing evidence. Full runs use `Job.Default` and `MemoryDiagnoser`. Inspect the BenchmarkDotNet reports for failed cases as well as the process exit code.

The original host referenced the current Pdf project directly. The committed revision host references its release-built assemblies instead, so the current checkout cannot silently replace the measured revision. Preserve the results and assembly hashes before removing the two clean detached worktrees with `git worktree remove`. Do not compare these dated results with different page counts, excluded serialization, or unvalidated output.
