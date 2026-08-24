# OfficeIMO AsciiDoc performance evidence (2026-08-24)

## Scope

This evidence covers `OfficeIMO.AsciiDoc` lossless parsing and
parse-plus-preserve-write over deterministic documents with 2, 100, and 1,000
sections. The UTF-8 inputs are 1,173 bytes, 56,323 bytes, and 571,429 bytes.
They exercise document attributes, anchors, headings, Unicode text, nested
inline formatting, cross references, lists, and tables.

Every accepted run requires error-free lossless parsing, the expected heading
and table counts, required semantic markers, and byte-identical preserve output.
The workload builds the public syntax tree, semantic block model, and inline
model; it is not a line scanner or tokenizer-only lane.

## Optimization result

Allocation tracing found eager source-slice copies in every syntax node and
literal inline. The optimized implementation keeps syntax and unchanged literal
text source-backed, retains already contiguous child collections, validates
node ownership and spans without allocating comparison substrings, and uses an
iterative bounded-depth syntax traversal.

The initial and final BenchmarkDotNet short runs used .NET 10.0.11 on Windows
11 with an AMD Ryzen 9 9950X3D2. Managed allocation fell by 31.7-32.3% across
all six workloads.

| Workload | Initial allocation | Final allocation | Reduction |
| --- | ---: | ---: | ---: |
| Small parse | 102.11 KiB | 69.72 KiB | 31.7% |
| Small parse + write | 102.11 KiB | 69.72 KiB | 31.7% |
| Normal parse | 4.66 MiB | 3.14 MiB | 32.5% |
| Normal parse + write | 4.66 MiB | 3.14 MiB | 32.5% |
| Large parse | 47.16 MiB | 31.95 MiB | 32.3% |
| Large parse + write | 47.16 MiB | 31.95 MiB | 32.3% |

The three-iteration elapsed run was noisy, especially at large scale, so it is
not used to claim an elapsed improvement. A longer final parse run with two
launches, five warmups, and ten measured iterations produced means of 18.20 us,
1.557 ms, and 49.615 ms for small, normal, and large. Relative to the initial
short-run means of 22.54 us, 1.869 ms, and 47.888 ms, the first two improved and
the large result differs by about 3.6%. That is not a material regression, but a
matched longer before/after lane is still required before claiming a time gain.

## Isolated memory and size evidence

Commit `0b134012239a504e1ff98fa3d68c9ba89195af9f` was measured from a clean
tree with three fresh child processes per lane on .NET 10.0.11. All checked-in
budgets passed on .NET 8 and .NET 10.

| Workload | Median elapsed | Allocation | Retained heap | Managed peak growth | Process peak | Output |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small parse | 0.23 ms | 0.07 MiB | 0.05 MiB | 0.09 MiB | 24.17 MiB | n/a |
| Small parse + write | 0.21 ms | 0.07 MiB | approximately 0 MiB | 0.09 MiB | 23.95 MiB | 1.15 KiB |
| Normal parse | 8.27 ms | 3.18 MiB | 2.24 MiB | 3.21 MiB | 30.14 MiB | n/a |
| Normal parse + write | 7.77 ms | 3.18 MiB | approximately 0 MiB | 3.21 MiB | 27.65 MiB | 55.00 KiB |
| Large parse | 82.21 ms | 32.31 MiB | 22.46 MiB | 32.43 MiB | 73.31 MiB | n/a |
| Large parse + write | 76.27 ms | 32.31 MiB | approximately 0 MiB | 32.43 MiB | 73.34 MiB | 558.04 KiB |

The isolated elapsed values include forced collection and a fresh-process
measurement boundary, so they are not interchangeable with steady-state
BenchmarkDotNet means. Allocation and retained/peak memory are the main evidence
from this lane. Parse-plus-write retains only the returned string after forced
collection; parse retains the complete public model. Preserve output is exactly
the input at all scales, for 1.00x output inflation.

## Comparison boundary and remaining target

No competitor ratio is published. AsciiDocNet's own documentation identifies
tables and list continuations as unsupported, while both constructs are required
by this corpus and OfficeIMO's measured public contract. Removing those features
to create a comparison would measure less work rather than competitive standing.

This is an owner-level baseline, not a contender or optimality claim. The large
558 KiB source still allocates 32.31 MiB and retains 22.46 MiB. Syntax nodes,
child arrays, inline arrays, and semantic strings remain the principal targets.
Any future comparison must validate the same output and public workflow; when
one exists, the repository-wide contender ceiling remains at most 2x in both
elapsed time and allocation.
