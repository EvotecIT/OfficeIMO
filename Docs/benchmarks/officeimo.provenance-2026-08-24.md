# OfficeIMO.Provenance structural-carrier evidence (2026-08-24)

## Result

The bounded structural inspection and selective-removal paths now have
reproducible coverage for PNG, TIFF, SVG, ZIP packages, and structured text.
The most severe local incident was PNG CRC processing: inspecting a PNG with a
1 MiB C2PA manifest fell from 99.78 ms to 6.22 ms, a 16.0x speedup, while
removal fell from 178.94 ms to 13.05 ms and from 1,046.05 KiB to 21.54 KiB
allocated. The optimized removal path therefore uses 48.6x less managed
allocation for that corpus.

This is an owner baseline rather than a contender-ratio claim. The measured
contract detects structurally valid C2PA carriers and selectively removes them
while preserving the surrounding format. It does not perform cryptographic
verification, trust evaluation, or a command-line process invocation. No
managed .NET implementation with the same public contract was accepted as an
equivalent comparison lane.

## Validated contract

Each deterministic fixture contains exactly one structurally valid C2PA JUMBF
manifest store. Preflight requires the expected format, one valid carrier, one
reported removal, no carrier after removal, and the exact expected output byte
count. The matrix covers 4 KiB and 1 MiB manifests in:

- a PNG `caBX` chunk;
- a TIFF C2PA tag whose payload is cleared without changing the container size;
- an SVG metadata element;
- a ZIP `META-INF/content_credential.c2pa` entry alongside preserved content;
- a structured-text C2PA manifest block.

The public removal result continues to snapshot caller-provided arrays. The
owning remover now transfers its private output buffer into the immutable result
without making a redundant copy. PNG validation, writing, and provenance share
one table-driven CRC implementation. PNG removal also sizes its output buffer
for the removed carrier, and non-XML text avoids exception-driven XML probing
without changing format-precedence behavior.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. The final source was clean
at commit `d14aab3d722f115c4490fee1079b52806f522a1b`.

| Format | Scale | Inspect mean | Inspect allocation | Remove mean | Remove allocation |
| --- | --- | ---: | ---: | ---: | ---: |
| PNG | 4 KiB manifest | 28.95 us | 8.23 KiB | 60.09 us | 21.54 KiB |
| PNG | 1 MiB manifest | 6.22 ms | 8.23 KiB | 13.05 ms | 21.54 KiB |
| TIFF | 4 KiB manifest | 2.17 us | 6.91 KiB | 5.50 us | 20.02 KiB |
| TIFF | 1 MiB manifest | 2.07 us | 6.91 KiB | 51.60 us | 1,040.87 KiB |
| SVG | 4 KiB manifest | 20.38 us | 94.10 KiB | 43.78 us | 214.71 KiB |
| SVG | 1 MiB manifest | 3.04 ms | 9,324.81 KiB | 7.14 ms | 24,107.41 KiB |
| ZIP | 4 KiB manifest | 7.13 us | 18.81 KiB | 89.89 us | 148.50 KiB |
| ZIP | 1 MiB manifest | 5.72 ms | 2,062.18 KiB | 39.62 ms | 8,311.74 KiB |
| Text | 4 KiB manifest | 24.50 us | 18.89 KiB | 43.14 us | 44.04 KiB |
| Text | 1 MiB manifest | 5.05 ms | 3,759.07 KiB | 9.01 ms | 8,887.71 KiB |

The same-machine initial production source was commit
`d569e99ba89b2355c2df7e97d336f6550860059e`, with the new benchmark harness
present only to exercise it. The representative changes are:

| Workload | Time change | Allocation change |
| --- | ---: | ---: |
| PNG inspect, 1 MiB | 99.78 ms to 6.22 ms, 93.8% lower | 8.54 KiB to 8.23 KiB |
| PNG remove, 1 MiB | 178.94 ms to 13.05 ms, 92.7% lower | 1,046.05 KiB to 21.54 KiB, 97.9% lower |
| PNG inspect, 4 KiB | 330.88 us to 28.95 us, 91.3% lower | unchanged at 8.23 KiB |
| TIFF remove, 1 MiB | 133.42 us to 51.60 us, 61.3% lower | 2,065.31 KiB to 1,040.87 KiB, 49.6% lower |
| Text inspect, 4 KiB | 24.03 us to 24.50 us | 33.74 KiB to 18.89 KiB, 44.0% lower |
| Text remove, 4 KiB | 52.89 us to 43.14 us, 18.4% lower | 61.89 KiB to 44.04 KiB, 28.8% lower |

Large SVG time improved by about 23%, but its allocation remained essentially
flat. Large ZIP removal allocation fell by 11.8%, while its `ShortRun` mean was
7.8% slower; that time result is not presented as an improvement and remains a
profiling target. Large structured text improved by about 20-21% in time, but
its multi-megabyte syntax and decoding allocation also remains open work.

## Isolated memory and output evidence

The evidence runner starts a fresh child process for every format, scale,
operation, and repetition. The following values are medians over three
repetitions for the 1 MiB-manifest corpus. Managed peak is sampled over the
eight-operation measurement batch; process peak is the absolute child-process
working-set peak. The source tree was clean at the final commit.

| Format | Operation | Retained heap | Managed batch peak | Process peak | Input bytes | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| PNG | Inspect | 1.20 KiB | 92.00 KiB | 35.43 MiB | 1,048,657 | n/a |
| PNG | Remove | 1.67 KiB | 195.90 KiB | 35.55 MiB | 1,048,657 | 69 |
| TIFF | Inspect | 1.21 KiB | 75.97 KiB | 34.03 MiB | 1,048,651 | n/a |
| TIFF | Remove | 1,025.70 KiB | 8,348.82 KiB | 35.35 MiB | 1,048,651 | 1,048,651 |
| SVG | Inspect | 1.02 KiB | 14,008.95 KiB | 49.32 MiB | 1,398,261 | n/a |
| SVG | Remove | 1.56 KiB | 44,314.16 KiB | 74.72 MiB | 1,398,261 | 117 |
| ZIP | Inspect | 1.04 KiB | 6,185.81 KiB | 45.60 MiB | 2,097,412 | n/a |
| ZIP | Remove | 1,025.93 KiB | 6,258.85 KiB | 48.27 MiB | 2,097,412 | 1,049,021 |
| Text | Inspect | 0.98 KiB | 7,525.27 KiB | 44.79 MiB | 1,398,205 | n/a |
| Text | Remove | 1.40 KiB | 6,500.99 KiB | 44.75 MiB | 1,398,205 | 13 |

TIFF intentionally preserves its original byte length because removal clears
the C2PA tag payload in place. The other large removal outputs prove that the
manifest carrier is physically absent while unrelated payload is retained.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- --filter '*ProvenanceBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\provenance\evidence.json
```

Raw BenchmarkDotNet and process evidence remain ignored machine-local
artifacts. This note retains the compact reproducible result and exact source
provenance.
