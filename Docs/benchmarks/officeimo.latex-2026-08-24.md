# OfficeIMO LaTeX performance evidence (2026-08-24)

## Scope

This evidence covers `OfficeIMO.Latex` lossless parsing and
parse-plus-preserve-write over deterministic documents with 2, 100, and 1,000
sections. The inputs are 1,470 bytes, 68,692 bytes, and 701,399 bytes. Each run
requires error-free lossless parsing, the expected headings and content markers,
and byte-identical preserve output before measurements are accepted.

The measured public parse contract builds the token stream, complete lossless
syntax tree, and OfficeIMO semantic model. It is not a tokenizer-only lane.

## Optimization result

The initial and final BenchmarkDotNet short runs used .NET 10.0.11 on Windows
11 with an AMD Ryzen 9 9950X3D2. The same benchmark corpus and validation were
used before and after the product changes.

| Workload | Initial mean | Final mean | Time reduction | Initial allocation | Final allocation | Allocation reduction |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small parse | 180.86 us | 55.86 us | 69.1% | 400.68 KiB | 162.29 KiB | 59.5% |
| Small parse + write | 145.89 us | 59.01 us | 59.6% | 401.45 KiB | 163.06 KiB | 59.4% |
| Normal parse | 16.07 ms | 4.07 ms | 74.7% | 17.49 MiB | 6.57 MiB | 62.4% |
| Normal parse + write | 16.55 ms | 4.03 ms | 75.6% | 17.49 MiB | 6.58 MiB | 62.4% |
| Large parse | 297.04 ms | 91.47 ms | 69.2% | 176.13 MiB | 67.20 MiB | 61.8% |
| Large parse + write | 266.21 ms | 91.52 ms | 65.6% | 176.17 MiB | 67.24 MiB | 61.8% |

Allocation tracing identified recursive traversal iterator frames, eager copies
of exact source slices, repeated complete-tree scans, and per-heading/per-list
command scans. The optimized implementation uses a bounded-depth iterative
traversal, lazy source-backed token and syntax text, retains already contiguous
child collections, performs one semantic discovery pass, and indexes labels and
list items once.

## Isolated memory and size evidence

Commit `1d482f33f447d0a72ae66295c32a7eed93fa1a32` was measured from a clean tree
with three fresh child processes per lane. All checked-in budgets passed.

| Workload | Median elapsed | Allocation | Retained heap | Managed peak growth | Process peak | Output |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small parse | 0.84 ms | 0.16 MiB | 0.13 MiB | 0.18 MiB | 25.35 MiB | n/a |
| Small parse + write | 0.87 ms | 0.16 MiB | approximately 0 MiB | 0.18 MiB | 25.54 MiB | 1.44 KiB |
| Normal parse | 23.57 ms | 6.58 MiB | 5.68 MiB | 6.62 MiB | 41.13 MiB | n/a |
| Normal parse + write | 24.08 ms | 6.59 MiB | approximately 0 MiB | 6.62 MiB | 36.79 MiB | 67.08 KiB |
| Large parse | 199.61 ms | 67.20 MiB | 57.03 MiB | 65.70 MiB | 104.25 MiB | n/a |
| Large parse + write | 214.03 ms | 67.24 MiB | approximately 0 MiB | 65.74 MiB | 97.96 MiB | 684.96 KiB |

The isolated elapsed values include the deliberately forced collection and
fresh-process measurement boundary and are therefore not interchangeable with
the steady-state BenchmarkDotNet means. Allocation and retained/peak memory are
the main evidence from this lane. Parse-plus-write retains only the returned
string after forced collection; plain parse retains the complete public model.
All preserve outputs exactly match their inputs, so output inflation is 1.00x.

## Comparison boundary and remaining target

No equivalent general .NET competitor was accepted. AsciiDocNet is unrelated to
this format and its documented profile is incomplete. LaTeX.Net exposes only a
small modifier-oriented model and does not expose its scanner as a public
NuGet-facing entry point, so reflection or a reduced syntax corpus would not
measure the same contract. Typesetting products such as Aspose.TeX produce PDF
or images and also perform materially different work.

This is an owner-level baseline, not a contender ratio or an optimality claim.
The large public model still retains about 57 MiB for a 701 KiB source because
it exposes 256,036 exact tokens plus the lossless syntax and semantic graphs.
Reducing that retained graph without weakening source spans, navigation,
editing, or exact preservation remains open work.
