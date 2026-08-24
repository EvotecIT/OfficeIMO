# OfficeIMO.OpenDocument ODS comparison evidence (2026-08-24)

## Result

OfficeIMO is comfortably within contender margins for equivalent dense-string
ODS creation and complete read traversal against OpenStandardLibrary. The
controlled short run measures OfficeIMO at 0.08-0.18x comparison time for
creation and 0.03-0.07x for reading. Creation allocation is lower; read
allocation is 1.25-1.26x.

Clean process-isolated evidence confirms the result while adding managed and
process peaks. Every allocation and peak ratio is under 1.63x. OfficeIMO's
validated ODS packages are 35.2-41.3% smaller. No Windows performance
remediation is justified by this evidence; non-Windows coverage remains useful.

## Equivalent contract

Both creation lanes build one `Data` sheet containing the same deterministic
100x8 or 1,000x8 dense string-cell corpus and return a complete in-memory ODS
package. Validation reopens both packages through OfficeIMO and requires the
same sheet, expanded row and cell counts, total content length, and first and
last boundary markers.

Both read lanes receive the exact same OfficeIMO-generated package, open it,
enumerate every populated cell through the implementation's public model, and
return the same content-length checksum. Fixture creation and semantic
validation remain outside the timed operation.

OpenStandardLibrary is benchmark-only and supports the equivalent spreadsheet
contract used here. No ODT or ODP comparison is inferred from the ODS result.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. The measured product and
workload source was clean at commit
`c336b7356fc29b9c71ce6d531c0edf10f598ae22`; the later evidence-runner commit
does not change either timed implementation.

Ratios below are OfficeIMO divided by OpenStandardLibrary. Lower is better.

| Workload | Scale | OfficeIMO mean | Comparison mean | Time ratio | OfficeIMO allocation | Comparison allocation | Allocation ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | Small, 800 cells | 1.239 ms | 15.404 ms | 0.08x | 1.27 MiB | 1.64 MiB | 0.77x |
| Create | Normal, 8,000 cells | 14.065 ms | 78.853 ms | 0.18x | 10.21 MiB | 14.50 MiB | 0.70x |
| Read | Small, 800 cells | 486.3 us | 15.132 ms | 0.03x | 494.99 KiB | 392.94 KiB | 1.26x |
| Read | Normal, 8,000 cells | 4.696 ms | 70.085 ms | 0.07x | 4,180.36 KiB | 3,331.77 KiB | 1.25x |

The short job is sufficient to classify these large differences, but its three
samples are not a basis for narrow universal throughput claims. The isolated
runner below is the provenance-bound peak and size record.

## Isolated peak-memory and size evidence

Each workload, scale, implementation, and repetition runs in a fresh child
process. Values are ratios of three-run medians from clean commit
`e5af8a1969dc753e20c1a31ffb509c64ef650c95`.

| Workload | Scale | Time ratio | Allocation ratio | Managed-peak ratio | Process-peak ratio | OfficeIMO output | Comparison output |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | Small | 0.27x | 0.82x | 1.63x | 0.81x | 4,520 B | 6,974 B |
| Create | Normal | 0.17x | 0.75x | 0.68x | 0.72x | 23,750 B | 40,482 B |
| Read | Small | 0.09x | 1.12x | 1.11x | 0.96x | Same 4,520 B input | Same input |
| Read | Normal | 0.11x | 1.17x | 1.17x | 0.83x | Same 23,750 B input | Same input |

Retained managed growth after full collection is near zero in every lane and
is not used to rank the implementations. Small creation has the weakest managed
peak ratio at 1.63x, while its absolute process peak remains lower. Normal
creation is lower on allocation and both peak measures.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\opendocument\validation.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- --filter '*Ods*ComparisonBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- evidence --repeat 3 --json .benchmark-artifacts\opendocument\ods-comparison-evidence.json
```

The comparison project stays outside the normal solution, so
OpenStandardLibrary does not enter OfficeIMO runtime or package graphs. Raw
reports remain ignored machine-local artifacts.
