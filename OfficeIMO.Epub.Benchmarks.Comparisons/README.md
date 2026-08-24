# OfficeIMO EPUB comparisons

This opt-in BenchmarkDotNet project compares complete EPUB open/read workflows
through OfficeIMO and VersOne.Epub plus the HtmlAgilityPack extraction workflow
recommended by VersOne's documentation. Both lanes receive the same
deterministic EPUB 3 package from a caller-owned memory stream, load the book
metadata and content, extract normalized visible chapter text, and enumerate
every file in spine reading order.

VersOne.Epub and HtmlAgilityPack remain benchmark-only dependencies. This
project is intentionally not part of `OfficeIMO.sln`, so normal restore, build,
test, and package operations do not acquire them. OfficeIMO also extracts
structured-content flags, diagnostics, navigation, and manifest resources as
part of its normal full load; the validator compares only the contracts both
workflows expose.

## Validate equivalent output

Run the semantic preflight without collecting timings:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\epub\comparison-validation.json
```

The validator checks title, creator, language, chapter count and order, exact
raw XHTML and normalized visible-text lengths and hashes, path hash, and
boundary paths. The JSON also records the shared input-package size. EPUB
creation is not compared because
OfficeIMO.Epub exposes a read/extraction API rather than a package writer.

## Measure time and allocations

Start with a dry execution check:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --filter '*EpubReadComparisonBenchmarks*' --job Dry --noOverwrite
```

Use the short job while changing a workload, then the default job for recorded
evidence:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --filter '*EpubReadComparisonBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --filter '*EpubReadComparisonBenchmarks*' --noOverwrite
```

BenchmarkDotNet artifacts and validation JSON are environment-specific. Keep
them under `.benchmark-artifacts` or another ignored output root.

## Capture isolated memory evidence

Use the process-isolated evidence runner to measure median elapsed time and
managed allocations together with retained managed heap, sampled managed-heap
peak, and process working-set peak. The retention phase keeps equivalent
metadata, raw XHTML, normalized text, and chapter-order projections alive for
both readers. Each child process loads only the selected reader; the parent
validates equivalent output before measurement and rejects fingerprint or
input-size differences across probes.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --evidence --repeat 3 --json .benchmark-artifacts\epub\comparison-evidence.json
```

The runner reports OfficeIMO divided by VersOne.Epub ratios. A ratio at or
below `2.00x` is the contender boundary for every measured dimension; lower is
better. Because both workflows read the same package and this library does not
write EPUB files, input bytes are the applicable size evidence and there is no
separate output-file-size ratio.
