# OfficeIMO.Word Open XML comparison evidence (2026-08-24)

## Result

The representative Word create, read, and rich replace workflows are within
the 2x contender ceiling against direct Open XML SDK code. Creating 100
paragraphs, formerly the only breach, is 1.86x by controlled BenchmarkDotNet
time and 1.41x by allocation. The 1,000-paragraph lane is 1.26x by time and
1.54x by allocation. These are tolerable contender margins, not an optimality
claim; further movement toward parity remains worthwhile.

Structured-report creation is 1.31-1.65x by mean time and 1.42-1.64x by
allocation. Rich load-replace-save is 1.34-1.52x by mean time and 1.48-1.75x by
allocation. OfficeIMO read is faster and lower-allocation at both measured
scales.

The work removed repeated style-catalog generation, redundant package clones,
the duplicate package-wide commit after Open XML parts were already saved,
no-op relationship, section, numbering, and table normalization passes, eager
table wrappers, and allocation-heavy replacement safety scans. Loaded external
documents retain relationship and table compatibility normalization. Cached
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
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 default job. The paragraph lanes were
rerun from clean commit `308e7606a275490c09f34bd2fc9acbc4a9821117`.
The other controlled rows retain the earlier clean
`f2d6ee38bebc4648c119a6363365f3f679148db0` run; the complete isolated table
below was refreshed at the newer commit.

Ratios below are OfficeIMO divided by Open XML SDK. Lower is better.

| Workload | Scale | OfficeIMO mean | SDK mean | Time ratio | OfficeIMO allocation | SDK allocation | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Create paragraphs | 100 | 566.6 us | 305.1 us | 1.86x | 602.51 KiB | 426.85 KiB | 1.41x |
| Create paragraphs | 1,000 | 3.251 ms | 2.574 ms | 1.26x | 2,439.35 KiB | 1,588.49 KiB | 1.54x |
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
| Create paragraphs | 100 | 1.60x | 1.41x | 1.41x | 1.00x | 18,705 B | 18,678 B |
| Create paragraphs | 1,000 | 0.99x | 1.54x | 1.54x | 1.01x | 21,416 B | 21,386 B |
| Create report | 100 | 1.87x | 1.33x | 1.33x | 1.04x | 19,858 B | 19,843 B |
| Create report | 1,000 | 1.76x | 1.29x | 1.29x | 1.01x | 24,435 B | 24,518 B |
| Read | 100 | 0.56x | 0.76x | 0.76x | 0.99x | Same 1,912-1,914 B input | Same input |
| Read | 1,000 | 1.16x | 0.83x | 0.84x | 1.00x | Same 4,512-4,514 B input | Same input |
| Replace and save | 100 | 1.20x | 1.63x | 1.63x | 1.06x | 18,658 B | 18,658 B |
| Replace and save | 1,000 | 0.86x | 1.45x | 1.45x | 1.06x | 21,246 B | 21,248 B |

Cold child-process time remains launch-sensitive, but the three-run medians now
remain below 2x in every lane. Allocation, managed peak, process peak, and output
size agree across the two measurement approaches. No time regression gate
should be based on a single cold child-process observation.

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
