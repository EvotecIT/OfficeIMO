# CommonMark semantic parse comparison evidence

This 2026-08-24 evidence compares source-free CommonMark semantic parsing
through OfficeIMO.Markdown and Markdig 1.3.2. Both implementations receive the
same Markdown. Setup renders both parsed documents to normalized semantic HTML
and rejects a corpus unless the results match before measurement.

The performance acceptance bands used for this work are:

- at or below 2x the comparison implementation in both elapsed time and managed
  allocation: contender range, though further improvement may still be useful;
- above 2x through 5x in either dimension: material remediation gap;
- above 5x in either dimension: unacceptable unless the observable contracts
  differ and the result is reported as a non-equivalent diagnostic.

A 40x ratio is an incident threshold, never a success boundary.

## Current equivalent-work result

The table is a bounded BenchmarkDotNet `ShortRun` from clean source commit
`3068f305f1dd7474fd7a576b56fa6d937b8c7296` on Windows 11
10.0.26200.9168, an AMD Ryzen 9 9950X3D2 with 16 physical cores, .NET SDK
10.0.111, and .NET 10.0.11 x64. Allocation is managed allocation per operation.

| Corpus | OfficeIMO mean | Markdig mean | Time ratio | OfficeIMO allocation | Markdig allocation | Allocation ratio | Classification |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| Large table | 1,726.90 us | 12,380.21 us | 0.14x | 1,874.26 KB | 1,017.04 KB | 1.84x | Contender |
| Long nested list | 1,045.62 us | 699.45 us | 1.49x | 2,549.54 KB | 1,754.82 KB | 1.45x | Contender |
| Normalization stress | 397.15 us | 256.17 us | 1.55x | 741.86 KB | 713.66 KB | 1.04x | Contender |
| Portable README | 75.74 us | 52.66 us | 1.44x | 154.10 KB | 84.27 KB | 1.83x | Contender |
| Rich AST | 124.70 us | 69.58 us | 1.79x | 261.20 KB | 145.49 KB | 1.80x | Contender |
| Technical document | 97.77 us | 62.57 us | 1.56x | 193.99 KB | 130.55 KB | 1.49x | Contender |
| Transcript | 73.81 us | 57.20 us | 1.29x | 194.70 KB | 101.96 KB | 1.91x | Contender |

All seven corpora meet the 2x elapsed-time and managed-allocation ceilings.
That establishes a contender baseline rather than a claim of optimality.
Transcript allocation at 1.91x, Large table at 1.84x, Portable README at 1.83x,
and Rich AST at 1.80x still warrant improvement for safer regression headroom.

## Retained and peak memory

An independent process-isolated run from the same clean commit used three fresh
child processes per engine and corpus. It validated the normalized semantic
output before measurement, warmed each process, timed an 8 MiB aggregate input
batch, and retained a separate parsed-document batch for heap measurement.

| Corpus | Retained managed-heap ratio | Sampled managed-heap peak ratio | Absolute process-peak ratio |
| --- | ---: | ---: | ---: |
| Large table | 1.51x | 1.59x | 1.41x |
| Long nested list | 1.52x | 1.42x | 1.59x |
| Normalization stress | 1.15x | 1.12x | 1.06x |
| Portable README | 1.30x | 1.83x | 1.40x |
| Rich AST | 1.10x | 1.24x | 1.27x |
| Technical document | 1.22x | 1.41x | 1.38x |
| Transcript | 1.41x | 1.81x | 1.52x |

Every retained-heap, sampled managed-heap peak, and absolute process-peak ratio
is within 2x. Working-set growth from the process baseline is retained in the
raw diagnostic output, but absolute process peak is the comparison metric
because independently launched runtimes do not have identical starting sets.

Each corpus produced one semantic SHA-256 fingerprint across both engines and
all repeats. Input byte counts and normalized semantic HTML byte counts also
matched. Parsing does not create a file, so semantic HTML bytes validate output
equivalence but are not claimed as a file-size comparison.

## What changed

The parser now avoids no-op reference-definition scans, redundant paragraph
construction, semantic-only source metadata, repeated nested option clones, and
general inline parsing for conservative flat CommonMark shapes. It also stores
rare node syntax metadata lazily, avoids the no-extension inline delegate, and
uses compact delimiter-closing indexes with pooled suffix summaries on modern
runtimes. The final pass also reuses repeated immutable input lines, skips
impossible emphasis frames, handles code spans nested in simple emphasis, and
uses a bounded semantic fast path for sequential double-asterisk pairs plus
unmatched literal delimiter runs. Nested and richer inline forms still fall
back to the general parser.

Against the earlier clean full-run evidence, managed allocation fell from
3,011.81 KB to 1,874.26 KB for Large table, 3,302.23 KB to 2,549.54 KB for Long
nested list, 2,611.46 KB to 741.86 KB for Normalization stress, 237.59 KB to
154.10 KB for Portable README, 403.20 KB to 261.20 KB for Rich AST, 321.44 KB
to 193.99 KB for Technical document, and 291.97 KB to 194.70 KB for Transcript.

The `ShortRun` uses three measurement iterations, so timing results close to a
ceiling remain directional rather than release-grade statistical proof. The
fresh-process runner supplies repeatability, retained-heap, peak-memory, and
provenance evidence, but its wall-clock timing includes more process and GC
noise than BenchmarkDotNet. Both lanes place all seven allocation ratios inside
2x; BenchmarkDotNet remains the primary timing evidence.

## Source-backed follow-up

Clean source commit `7a7c2a166a1248ef8ccf8918e2f4cb5795b6d834` reduces
the additional graph retained by `MarkdownReader.Parse` and
`ParseWithSyntaxTree` without changing their snapshot semantics. The parser now
reuses an associated syntax node for unmodified source metadata, packs optional
source positions, allocates rare syntax metadata and generated diagnostics only
when needed, and avoids materializing empty list-item block collections.

The full BenchmarkDotNet job for `LongNestedList` measured:

| API | Mean | Allocation | Time vs Markdig | Allocation vs Markdig | Time vs semantic | Allocation vs semantic |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `ParseSemantic` | 0.908 ms | 2.37 MB | 1.06x | 1.39x | 1.00x | 1.00x |
| `Parse` source-backed | 3.501 ms | 7.90 MB | 4.09x | 4.62x | 3.86x | 3.33x |
| `ParseWithSyntaxTree` | 5.571 ms | 11.35 MB | 6.51x | 6.64x | 6.14x | 4.79x |
| Markdig | 0.856 ms | 1.71 MB | 1.00x | 1.00x | 0.94x | 0.72x |

The equivalent semantic lane is inside the 2x contender band. Markdig does not
retain OfficeIMO's complete trivia and syntax-ownership contract, so its
source-backed ratios are a diagnostic floor rather than an equivalent-product
ranking. They nevertheless expose a material internal cost: source-backed
parsing is still 3.86x semantic time and 3.33x semantic allocation on this
workload. The explicit syntax-tree lane remains above 5x Markdig in both
dimensions and is not acceptable as a general performance margin.

Against the earlier clean `LongNestedList` source-backed result, mean time fell
from 5.896 ms to 3.501 ms (40.6%) and allocation from 9.08 MB to 7.90 MB
(13.0%). The explicit syntax-tree lane fell from 9.228 ms to 5.571 ms (39.6%)
and from 13.15 MB to 11.35 MB (13.7%). A same-size 64-document heap probe also
reduced source-backed retained managed memory from 361,220,688 bytes to
264,058,896 bytes (26.9%).

The fresh-process runner then measured all seven corpora three times. The table
below uses per-corpus medians and shows the source-backed result against Markdig
as a diagnostic floor:

| Corpus | Time | Allocation | Retained heap | Managed peak | Absolute process peak |
| --- | ---: | ---: | ---: | ---: | ---: |
| Portable README | 3.79x | 7.69x | 3.32x | 3.81x | 2.82x |
| Transcript | 3.73x | 6.76x | 3.59x | 3.91x | 2.70x |
| Technical document | 4.46x | 7.38x | 3.42x | 3.47x | 2.60x |
| Rich AST | 3.50x | 7.17x | 2.94x | 2.83x | 2.16x |
| Long nested list | 5.33x | 4.61x | 3.67x | 3.39x | 2.71x |
| Large table | 0.44x | 6.35x | 4.03x | 4.46x | 2.92x |
| Normalization stress | 1.83x | 6.94x | 4.35x | 3.85x | 2.55x |

Every run produced the same normalized semantic fingerprint for all engines.
The source-backed allocation and retained-memory ratios remain the primary open
Markdown performance problem; clearing a 40x incident threshold is not a
completion criterion.

Reproduce the validation and current bounded runs with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --validate-equivalence
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --filter "*MarkdownParseBenchmarks*CommonMark*" --exporters json --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --parse-evidence --repeat 3 --json .benchmark-artifacts\markdown-parse-evidence.json
```

Raw BenchmarkDotNet and process-runner output remains local and excluded from
the repository's small committed evidence set.
