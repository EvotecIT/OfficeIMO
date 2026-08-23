# RTF-to-HTML comparison evidence

This 2026-08-24 run compares complete RTF parsing and HTML rendering through
OfficeIMO.Rtf plus OfficeIMO.Html.Rtf and through RtfPipe. Both implementations
receive the same RTF. Validation requires the same ordered semantic tokens,
record markers, table count, cell count, and ordered cell content before any
timing or output-size evidence is accepted.

This was a full BenchmarkDotNet run from clean source commit
`db376a1ca2d90a90b57adc677bc6bbdd851860d9` on Windows 11
10.0.26200.9168, an AMD Ryzen 9 9950X3D2 with 16 physical cores, .NET SDK
10.0.111, and .NET 10.0.11 x64. Allocation is managed allocation per
operation. Output size is the UTF-8 byte count of the validated HTML.

| Corpus | Implementation | Mean | Allocation | Output bytes | OfficeIMO time ratio | OfficeIMO allocation ratio | OfficeIMO size ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | OfficeIMO | 108.30 us | 323.01 KB | 2,559 | 0.25x | 0.58x | 0.75x |
| Small | RtfPipe | 437.50 us | 552.57 KB | 3,405 |  |  |  |
| Medium | OfficeIMO | 1,640.90 us | 5,094.82 KB | 48,763 | 0.38x | 0.74x | 0.91x |
| Medium | RtfPipe | 4,309.90 us | 6,890.45 KB | 53,655 |  |  |  |
| Large | OfficeIMO | 31,022.20 us | 40,268.49 KB | 390,139 | 0.55x | 0.75x | 0.92x |
| Large | RtfPipe | 56,630.50 us | 53,565.74 KB | 424,781 |  |  |  |
| Producer | OfficeIMO | 106.90 us | 305.55 KB | 3,531 | 0.30x | 0.67x | 3.14x |
| Producer | RtfPipe | 355.80 us | 454.02 KB | 1,125 |  |  |  |

OfficeIMO is faster and allocates less in every validated corpus. It also emits
less HTML for the three generated corpora. The producer fixture needs a
qualified interpretation: both outputs have the same validated content and
table structure, but OfficeIMO additionally preserves section/page behavior,
text direction, paragraph spacing, font properties, and detailed table-cell
layout that is absent from the RtfPipe output. Its 3.14x byte ratio is therefore
not an equivalent visual-fidelity size comparison, but it remains a material
output-overhead signal and a candidate for further lossless style hoisting.

Before the measured commit, OfficeIMO emitted 3,963, 78,013, 624,139, and 6,789
bytes for the same four corpora. Coalescing adjacent equal-format runs in the
semantic writer reduced those sizes to 2,559, 48,763, 390,139, and 3,531 bytes.
Trusted round-trip output retains source run boundaries.

Reproduce the validation and full run with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Rtf.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\rtf\validation.json
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload rtfhtml -RunMode full -Framework net10.0
```

Raw BenchmarkDotNet output remains local and excluded from the repository's
small committed evidence set.
