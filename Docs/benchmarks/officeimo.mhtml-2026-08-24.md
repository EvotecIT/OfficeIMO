# OfficeIMO.Mhtml comparison evidence (2026-08-24)

## Result

Validated MHTML read and write workflows are inside the 2x contender ceiling
for elapsed time and managed allocation at every scale. Reads use 0.88-1.08x
MimeKit plus AngleSharp elapsed time and 0.62-0.92x its allocation. Writes use
0.67-0.73x MimeKit elapsed time and 0.70-1.40x its allocation. OfficeIMO's
validated serialized archives are 2.7-13.6% smaller.

The comparison preflight also exposed and fixed an interoperability defect:
when a multipart/related producer placed its selected HTML root first,
OfficeIMO projected the root again as a resource. Both root-first and
resource-first archives now expose only the actual related resources.

## Equivalent contract

The read lane uses the same MimeKit-produced archive in both implementations.
OfficeIMO performs its normal bounded MIME and HTML parsing. The comparison
uses MimeKit for MIME, the same AngleSharp version for an HTML DOM, and retains
every decoded related-resource payload. The write lane starts from equivalent
prepared multipart/related models.

Before timing, both outputs are read through both implementations and must
agree on subject, selected root Content-ID and Content-Location, source HTML,
HTML element count, ordered resource metadata, decoded byte lengths, and
SHA-256 payload hashes.

| Scale | HTML source | Resources | Decoded resource bytes | OfficeIMO output | MimeKit output | Size ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | 4 KiB | 2 | 8,192 | 18,132 | 20,980 | 0.864x |
| Normal | 64 KiB | 16 | 524,288 | 814,295 | 856,745 | 0.950x |
| Large | 512 KiB | 64 | 8,388,608 | 12,231,016 | 12,569,213 | 0.973x |

Loaded resources now share the private decoded payload already owned by the
private MIME model. Public construction and `Content` access still return
independent snapshots, so callers cannot mutate archive state. This removed
redundant full-resource copies without weakening the immutability contract.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. The final source was clean
at commit `c678944eb4b4b7474857351d7f39ae0c516ea04d`.

### Read

| Scale | OfficeIMO mean | Comparison mean | Time ratio | OfficeIMO allocation | Comparison allocation | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | 155.4 us | 162.3 us | 0.96x | 234.03 KiB | 254.12 KiB | 0.92x |
| Normal | 2.790 ms | 3.172 ms | 0.88x | 3,503.64 KiB | 4,801.83 KiB | 0.73x |
| Large | 36.493 ms | 33.736 ms | 1.08x | 36,596.39 KiB | 59,250.00 KiB | 0.62x |

The initial same-machine working source was based on parent commit
`3d395b5069545b033684c84c8061180a61717c81`, with the new cross-producer root
fix and benchmark harness present but the payload-ownership optimization absent.
It allocated 258.41 KiB, 5,042.73 KiB, and 61,189.38 KiB for Small, Normal, and
Large. The optimized source therefore reduced allocation by 9.4%, 30.5%, and
40.2%. Timing remains close enough that no broad read-speed improvement is
inferred from the short samples.

### Write

| Scale | OfficeIMO mean | MimeKit mean | Time ratio | OfficeIMO allocation | MimeKit allocation | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | 15.26 us | 22.71 us | 0.67x | 99.09 KiB | 141.58 KiB | 0.70x |
| Normal | 570.82 us | 777.27 us | 0.73x | 4,563.93 KiB | 3,431.05 KiB | 1.33x |
| Large | 7.118 ms | 10.536 ms | 0.68x | 70,262.58 KiB | 50,188.06 KiB | 1.40x |

Large-write allocation is within the contender ceiling but remains the weakest
margin. The writer produces a 12.23 MB returned byte array and allocates about
68.6 MiB while applying OfficeIMO's deterministic MIME policy.

## Isolated memory and output evidence

The isolated runner starts a fresh child process for each scale, operation,
implementation, and repetition. Values below are medians over three repetitions.
Managed peak is sampled over the measurement batch; process peak is the absolute
child working-set peak.

| Operation | Scale | Implementation | Allocation/op | Retained heap | Managed batch peak | Process peak |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| Read | Small | OfficeIMO | 0.27 MiB | 0.11 MiB | 8.61 MiB | 58.44 MiB |
| Read | Small | MimeKit + AngleSharp | 0.28 MiB | 1.38 MiB | 9.12 MiB | 59.66 MiB |
| Read | Normal | OfficeIMO | 3.86 MiB | 1.76 MiB | 15.21 MiB | 86.34 MiB |
| Read | Normal | MimeKit + AngleSharp | 5.06 MiB | 8.21 MiB | 40.59 MiB | 100.27 MiB |
| Read | Large | OfficeIMO | 37.90 MiB | 17.69 MiB | 75.90 MiB | 315.03 MiB |
| Read | Large | MimeKit + AngleSharp | 59.91 MiB | 39.97 MiB | 105.55 MiB | 320.64 MiB |
| Write | Small | OfficeIMO | 0.10 MiB | 0.02 MiB | 3.15 MiB | 53.83 MiB |
| Write | Small | MimeKit | 0.14 MiB | 0.02 MiB | 4.45 MiB | 55.00 MiB |
| Write | Normal | OfficeIMO | 4.46 MiB | 0.78 MiB | 9.05 MiB | 88.90 MiB |
| Write | Normal | MimeKit | 3.35 MiB | 0.82 MiB | 6.21 MiB | 88.20 MiB |
| Write | Large | OfficeIMO | 68.62 MiB | 11.67 MiB | 137.25 MiB | 325.77 MiB |
| Write | Large | MimeKit | 49.01 MiB | 11.99 MiB | 98.02 MiB | 327.07 MiB |

All retained-heap, managed-peak, and absolute process-peak ratios are within
the 2x contender ceiling. Large-write managed peak is approximately 1.40-1.46x
MimeKit across the isolated samples. The normal MimeKit read retained/peak
samples varied with collection timing, but OfficeIMO remained lower in all
three repetitions.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- --filter '*Mhtml*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- evidence --repeat 3 --json .benchmark-artifacts\mhtml\evidence.json
```

The comparison project stays outside the normal solution. MimeKit and the
benchmark tooling therefore do not enter OfficeIMO runtime or package graphs.
Raw reports remain ignored machine-local artifacts.
