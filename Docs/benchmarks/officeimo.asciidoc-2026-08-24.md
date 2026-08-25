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

The initial and current BenchmarkDotNet short runs used .NET 10.0.11 on Windows
11 with an AMD Ryzen 9 9950X3D2. A later syntax-graph pass retained compact
offsets instead of full line/column coordinates and shared the source/original
text backing slot while preserving the public syntax contract. Managed
allocation is now 38.1-38.9% below the initial implementation across all six
workloads.

| Workload | Initial allocation | Current allocation | Reduction |
| --- | ---: | ---: | ---: |
| Small parse | 102.11 KiB | 63.17 KiB | 38.1% |
| Small parse + write | 102.11 KiB | 63.17 KiB | 38.1% |
| Normal parse | 4.66 MiB | 2.85 MiB | 38.9% |
| Normal parse + write | 4.66 MiB | 2.85 MiB | 38.9% |
| Large parse | 47.16 MiB | 28.98 MiB | 38.5% |
| Large parse + write | 47.16 MiB | 28.98 MiB | 38.5% |

The current short parse run produced means of 14.62 us, 0.923 ms, and 39.293 ms
for small, normal, and large. The preceding published parse run used a different
longer job, so the two BenchmarkDotNet sets are not treated as a matched timing
comparison. The fresh-process lane below did use the same harness before and
after this pass; its large median fell from 82.21 ms to 57.95 ms. Allocation and
retained memory remain the primary evidence for this graph change.

## Isolated memory and size evidence

Commit `91a42cbe28d6c3c63f3dde2a2b87da5ce7386132` was measured from a clean
tree with three fresh child processes per lane on .NET 10.0.11. All checked-in
budgets passed on .NET 8 and .NET 10.

| Workload | Median elapsed | Allocation | Retained heap | Managed peak growth | Process peak | Output |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small parse | 0.17 ms | 0.06 MiB | 0.04 MiB | 0.09 MiB | 24.12 MiB | n/a |
| Small parse + write | 0.17 ms | 0.06 MiB | approximately 0 MiB | 0.09 MiB | 24.33 MiB | 1.15 KiB |
| Normal parse | 6.17 ms | 2.89 MiB | 1.94 MiB | 2.91 MiB | 29.69 MiB | n/a |
| Normal parse + write | 5.94 ms | 2.89 MiB | approximately 0 MiB | 2.91 MiB | 27.75 MiB | 55.00 KiB |
| Large parse | 57.95 ms | 29.34 MiB | 19.50 MiB | 29.45 MiB | 69.88 MiB | n/a |
| Large parse + write | 60.94 ms | 29.34 MiB | approximately 0 MiB | 29.45 MiB | 69.65 MiB | 558.04 KiB |

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
558 KiB source still allocates 29.34 MiB and retains 19.50 MiB. Child
collections, inline arrays, and semantic strings remain the principal targets.
Any future comparison must validate the same output and public workflow; when
one exists, the repository-wide contender ceiling remains at most 2x in both
elapsed time and allocation.
