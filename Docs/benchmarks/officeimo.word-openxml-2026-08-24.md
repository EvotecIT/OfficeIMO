# OfficeIMO.Word Open XML comparison evidence (2026-08-24)

## Result

The representative Word create, read, and rich replace workflows are within
contender margins against direct Open XML SDK code except for one narrow edge:
creating 100 paragraphs is 2.06x slower in the controlled BenchmarkDotNet run.
Its allocation is 1.81x, and the same workload at 1,000 paragraphs is 1.39x
slower with 1.65x allocation. This small fixed-cost gap remains open work; it is
not rounded into a pass.

Structured-report creation is 1.31-1.65x by mean time and 1.42-1.64x by
allocation. Rich load-replace-save is 1.34-1.52x by mean time and 1.48-1.75x by
allocation. OfficeIMO read is faster and lower-allocation at both measured
scales.

The work removed repeated style-catalog generation, redundant package clones,
eager table wrappers, and allocation-heavy replacement safety scans. Cached
package and style state remains bounded, preserves runtime style overrides, and
is covered on .NET Framework 4.7.2, .NET 8, and .NET 10.

## Equivalent contract

The public comparison uses only OfficeIMO.Word and the MIT-licensed Open XML
SDK. It does not execute or publish the combined DocX or NPOI lanes.

Both creation implementations start from the same rich Office-compatible
package shell. Paragraph creation preserves equivalent paragraph properties.
Report creation produces the same title, summary, header, footer, two-column
table, cell values, table style, grid, cell widths, and table-look settings.
Read traverses the same styled input and must return the same paragraph count,
character count, and deterministic checksum. Replace starts from the same rich
input, replaces every token, preserves the style-catalog count, and serializes
the result.

Every output is reopened and structurally validated before measurement. The
comparison rejects minimal raw-SDK packages that omit the package and table
defaults emitted by OfficeIMO because their work and output are not equivalent.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 default job. The final source was clean
at commit `f2d6ee38bebc4648c119a6363365f3f679148db0`.

Ratios below are OfficeIMO divided by Open XML SDK. Lower is better.

| Workload | Scale | OfficeIMO mean | SDK mean | Time ratio | OfficeIMO allocation | SDK allocation | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Create paragraphs | 100 | 816.4 us | 396.7 us | 2.06x | 773.60 KiB | 426.85 KiB | 1.81x |
| Create paragraphs | 1,000 | 3.895 ms | 2.806 ms | 1.39x | 2,615.69 KiB | 1,588.48 KiB | 1.65x |
| Create report | 100 | 1.455 ms | 883.7 us | 1.65x | 1,224.88 KiB | 747.41 KiB | 1.64x |
| Create report | 1,000 | 8.359 ms | 6.381 ms | 1.31x | 5,815.67 KiB | 4,096.73 KiB | 1.42x |
| Read | 100 | 314.5 us | 929.3 us | 0.34x | 241.55 KiB | 318.21 KiB | 0.76x |
| Read | 1,000 | 1.833 ms | 2.111 ms | 0.87x | 1,320.45 KiB | 1,583.93 KiB | 0.83x |
| Replace and save | 100 | 889.3 us | 583.4 us | 1.52x | 754.15 KiB | 430.46 KiB | 1.75x |
| Replace and save | 1,000 | 3.257 ms | 2.434 ms | 1.34x | 2,180.00 KiB | 1,471.42 KiB | 1.48x |

## Isolated peak-memory and output evidence

The isolated runner starts a fresh child process for each workload, scale,
implementation, and repetition. Values below are ratios of the three-run
medians. Managed peak is sampled over the operation batch; process peak is the
absolute child working-set peak.

| Workload | Scale | Time ratio | Allocation ratio | Managed-peak ratio | Process-peak ratio | OfficeIMO output | SDK output |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Create paragraphs | 100 | 1.72x | 1.79x | 1.79x | 1.01x | 18,704 B | 18,678 B |
| Create paragraphs | 1,000 | 1.13x | 1.65x | 1.65x | 1.01x | 21,413 B | 21,383 B |
| Create report | 100 | 1.48x | 1.55x | 1.55x | 1.04x | 19,860 B | 19,840 B |
| Create report | 1,000 | 2.02x | 1.33x | 1.33x | 1.01x | 24,439 B | 24,518 B |
| Read | 100 | 0.54x | 0.81x | 0.81x | 0.98x | Same 1,912-1,913 B input | Same input |
| Read | 1,000 | 0.94x | 0.88x | 0.88x | 1.00x | Same 4,512-4,514 B input | Same input |
| Replace and save | 100 | 1.50x | 1.80x | 1.80x | 1.07x | 18,658 B | 18,657 B |
| Replace and save | 1,000 | 1.03x | 1.48x | 1.48x | 1.06x | 21,247 B | 21,246 B |

The isolated 1,000-row report time is launch-sensitive: its three-run median is
2.02x while the controlled BenchmarkDotNet mean is 1.31x with narrow confidence
intervals. Allocation, managed peak, process peak, and output size agree across
the two measurement approaches. No time regression gate should be based on a
single cold child-process observation.

All output-size differences are below 0.2%. The packages carry the same
validated semantic and Office-compatible structural contract; smaller output
is not achieved by omitting required content.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- validate-openxml
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- --filter '*Word*OpenXmlEvidenceBenchmarks*'
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\word\openxml-evidence.json
```

The benchmark project stays outside the normal solution. Open XML SDK is
already the format foundation used by OfficeIMO.Word; DocX and NPOI remain
isolated comparison-only dependencies. Raw reports remain ignored
machine-local artifacts.
