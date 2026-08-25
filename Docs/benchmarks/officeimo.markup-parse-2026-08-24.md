# OfficeIMO Markup parse performance evidence (2026-08-24)

## Scope

This evidence compares the Markdown-compatible semantic parse workflow in
`OfficeIMO.Markup` with Markdig 1.3.2. Both lanes use CommonMark plus pipe
tables and parse the same deterministic headings, paragraphs, formatted inline
text, lists, and tables. Office-specific directives are outside this comparison
because Markdig does not produce the corresponding document semantics.

The small, normal, and large UTF-8 inputs are 1,004, 48,706, and 495,412 bytes.
Before measurement, both parsers must produce identical canonical semantic
event streams with 17, 801, and 8,001 events and matching SHA-256 digests. A
result is rejected when the event count or digest differs.

Markdig remains an opt-in benchmark dependency in
`OfficeIMO.Markup.Benchmarks`; it is not part of the normal solution, runtime
projects, or OfficeIMO packages.

## Diagnostic baseline and optimization

The first exploratory run showed OfficeIMO at 20.3-28.3x Markdig's mean time
and 17.6-19.5x its managed allocation. That run was useful as an incident
signal, but it was not valid contender evidence: the OfficeIMO lane enabled its
broader default grammar while the Markdig lane used CommonMark plus tables.

The final harness makes the enabled grammar equivalent. The implementation
also removes work that `OfficeMarkupParser` discarded: it uses an internal
semantic-projection parse, skips source-tree binding and rich table-cell models,
avoids absent fenced-block transforms, reuses stateless parser and default-option
instances, reduces inline and table scans, and lazily renders the public
`SourceText` value. Focused tests preserve registered abbreviations and exact
heading, paragraph, list, and table `SourceText` output.

## BenchmarkDotNet result

Commit `b087466c7cf2d0521ea9e8f3979520e454c0711c` was measured from a clean
tree on .NET 10.0.11 and Windows 11 using an AMD Ryzen 9 9950X3D2. The run used
one launch, three warmups, and five measured iterations.

| Scale | OfficeIMO mean | Markdig mean | Time ratio | OfficeIMO allocation | Markdig allocation | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | 17.71 us | 15.41 us | 1.15x | 38.68 KiB | 30.60 KiB | 1.26x |
| Normal | 1.070 ms | 0.817 ms | 1.31x | 1,826.46 KiB | 1,533.74 KiB | 1.19x |
| Large | 12.278 ms | 10.555 ms | 1.16x | 18,658.74 KiB | 16,473.38 KiB | 1.13x |

Every workload is below the 2x contender ceiling in both dimensions. The
checked-in regression budget is tighter at 1.75x elapsed and 1.60x allocation.
Contender status is not an optimality claim: normal elapsed time and small
allocation remain the largest steady-state margins.

## Isolated memory evidence

Three fresh child processes per engine and scale recorded allocation, retained
managed heap, sampled managed-heap peak, and absolute process peak. Median
OfficeIMO-to-Markdig ratios were:

| Scale | Elapsed | Allocation | Retained heap | Managed peak | Process peak |
| --- | ---: | ---: | ---: | ---: | ---: |
| Small | 0.63x | 1.29x | 1.20x | 1.29x | 1.06x |
| Normal | 0.86x | 1.21x | 1.19x | 1.13x | 1.25x |
| Large | 1.26x | 1.11x | 1.19x | 1.31x | 1.28x |

The isolated elapsed lane includes forced collection and process-boundary
effects, so BenchmarkDotNet is the primary steady-state timing evidence.
Allocation and retained/peak memory are the primary evidence from this lane.

Parsing creates no file, so output byte size is not applicable. Input bytes,
semantic event counts, and semantic digests are recorded instead as the
validated output contract. The next useful work is to move the remaining
1.25-1.31x margins toward parity and add Linux and macOS evidence.
