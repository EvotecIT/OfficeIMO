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

A later source-graph pass replaced the full line/column coordinates retained by
every token and syntax node with compact offsets while preserving the public
`LatexSourceSpan` contract. Token text/value and syntax source/original-text
state now share their common backing slots, and full parsing reuses one source
line map. Clean commit `eb186c672db974023ed867169d580d35cac173e4` produced:

| Workload | Prior mean | Current mean | Time reduction | Prior allocation | Current allocation | Allocation reduction |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small parse | 55.86 us | 48.68 us | 12.9% | 162.29 KiB | 134.01 KiB | 17.4% |
| Small parse + write | 59.01 us | 50.30 us | 14.8% | 163.06 KiB | 134.75 KiB | 17.4% |
| Normal parse | 4.07 ms | 2.83 ms | 30.6% | 6.57 MiB | 5.30 MiB | 19.3% |
| Normal parse + write | 4.03 ms | 2.57 ms | 36.3% | 6.58 MiB | 5.31 MiB | 19.3% |
| Large parse | 91.47 ms | 60.96 ms | 33.4% | 67.20 MiB | 54.47 MiB | 18.9% |
| Large parse + write | 91.52 ms | 65.48 ms | 28.5% | 67.24 MiB | 54.51 MiB | 18.9% |

## Isolated memory and size evidence

Commit `eb186c672db974023ed867169d580d35cac173e4` was measured from a clean tree
with three fresh child processes per lane. All checked-in budgets passed.

| Workload | Median elapsed | Allocation | Retained heap | Managed peak growth | Process peak | Output |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small parse | 0.46 ms | 0.13 MiB | 0.10 MiB | 0.15 MiB | 25.44 MiB | n/a |
| Small parse + write | 0.47 ms | 0.13 MiB | approximately 0 MiB | 0.15 MiB | 25.45 MiB | 1.44 KiB |
| Normal parse | 11.69 ms | 5.30 MiB | 4.42 MiB | 5.34 MiB | 40.98 MiB | n/a |
| Normal parse + write | 12.51 ms | 5.31 MiB | approximately 0 MiB | 5.34 MiB | 36.38 MiB | 67.08 KiB |
| Large parse | 123.49 ms | 54.47 MiB | 44.43 MiB | 54.61 MiB | 91.32 MiB | n/a |
| Large parse + write | 136.16 ms | 54.51 MiB | approximately 0 MiB | 54.65 MiB | 87.82 MiB | 684.96 KiB |

The isolated elapsed values include the deliberately forced collection and
fresh-process measurement boundary and are therefore not interchangeable with
the steady-state BenchmarkDotNet means. Allocation and retained/peak memory are
the main evidence from this lane. Parse-plus-write retains only the returned
string after forced collection; plain parse retains the complete public model.
All preserve outputs exactly match their inputs, so output inflation is 1.00x.
Against the preceding clean graph baseline, the large parse now allocates 18.9%
less, retains 22.1% less managed memory, and lowers managed peak growth by
16.9%. Fresh-process elapsed time also fell, but the steady-state
BenchmarkDotNet result above is the primary throughput evidence.

## Comparison boundary and remaining target

No equivalent general .NET competitor was accepted. AsciiDocNet is unrelated to
this format and its documented profile is incomplete. LaTeX.Net exposes only a
small modifier-oriented model and does not expose its scanner as a public
NuGet-facing entry point, so reflection or a reduced syntax corpus would not
measure the same contract. Typesetting products such as Aspose.TeX produce PDF
or images and also perform materially different work.

This is an owner-level baseline, not a contender ratio or an optimality claim.
The large public model still retains about 44.4 MiB for a 701 KiB source because
it exposes 256,036 exact tokens plus the lossless syntax and semantic graphs.
Reducing that retained graph without weakening source spans, navigation,
editing, or exact preservation remains open work.
