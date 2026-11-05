# PDF reading performance

## Workload and baseline

The browser-print workload has 115 pages, 150 sections, four paragraphs and an
eight-row bordered table per section. Measure a fresh `Load` and `Read().Text`,
with one warmup and three samples per profile. Extract Markdown separately and
compare complete Text and Markdown with the baseline. Timing is opt-in evidence,
outside correctness tests.

Baseline: `186819ce958a928a0d3b13cd41e9d8f808d1481b`, macOS ARM64,
.NET 10 Release, 5 November 2026. The supplied reproduction harness measures
the exact input through the public API; run its `verify <fixtures> <output>`
mode. Structured read samples were 6423, 6099 and 6078 ms. Fast samples were
4075, 3784 and 4909 ms, with images and form widgets disabled. Load is excluded
from those numbers. The lock serializes cooperating heavy validation, but these
samples do not establish a quiet-host comparison.

## Investigation

Text, invoked-resource discovery, graphics-effect transitions and visual
primitives interpret the same root content independently. The visual path also
scans invocations before and after primitive extraction. These passes carry
different transforms, visibility, callbacks and budgets; a global token cache
is not established as safe.

First candidate: reuse existing root invoked-resource evidence in ruling
analysis and skip invocation-only traversal when the evidence proves there are
no invocations. Keep the canonical visual parser, clipping and paint validation.
Measure before deciding whether broader interpretation sharing is justified.

Stop when both profiles improve materially with byte-identical Text/Markdown
on the workload and focused correctness proof passes. Revert an optimization
that changes extraction, cancellation or resource-limit enforcement. Record
unsuccessful attempts here before trying another approach.
