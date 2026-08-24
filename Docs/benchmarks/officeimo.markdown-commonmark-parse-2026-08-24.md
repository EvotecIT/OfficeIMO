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

This work does not close the separate source-backed parser, which retains source
spans, trivia, and syntax ownership and needs its own richer performance
contract.

Reproduce the validation and current bounded runs with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --validate-equivalence
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --filter "*MarkdownParseBenchmarks*ParseSemantic*CommonMark*" --job Short --exporters json --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --parse-evidence --repeat 3 --json .benchmark-artifacts\markdown-parse-evidence.json
```

Raw BenchmarkDotNet and process-runner output remains local and excluded from
the repository's small committed evidence set.
