# CommonMark semantic parse comparison evidence

This 2026-08-24 run compares source-free CommonMark semantic parsing through
OfficeIMO.Markdown and Markdig 1.3.2. Both implementations receive the same
Markdown. Benchmark setup renders both parsed documents to normalized semantic
HTML and rejects a corpus unless the results match before measurement.

The performance acceptance bands used for this work are:

- at or below 2x the comparison implementation in both mean time and managed
  allocation: contender range;
- above 2x through 5x in either dimension: material remediation gap;
- above 5x in either dimension: unacceptable unless the observable contracts
  differ and the result is reported as a non-equivalent diagnostic.

The seven-corpus matrix was a full BenchmarkDotNet run from clean source commit
`a9280609bbdd9cf508245689fed3e94ae8e1afd8` on Windows 11
10.0.26200.9168, an AMD Ryzen 9 9950X3D2 with 16 physical cores, .NET SDK
10.0.111, and .NET 8.0.30 x64. Allocation is managed allocation per operation.
The Rich AST row was then superseded by a focused full run from clean source
commit `926b52f87eb6d617780a9653c8725657706c0aca` on the same host and runtime.

| Corpus | Implementation | Mean | Allocation | OfficeIMO time ratio | OfficeIMO allocation ratio | Classification |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| Large table | OfficeIMO.Semantic | 1,770.92 us | 3,011.81 KB | 0.14x | 2.96x | Material allocation gap |
| Large table | Markdig | 12,926.22 us | 1,016.97 KB |  |  |  |
| Long nested list | OfficeIMO.Semantic | 1,373.37 us | 3,302.23 KB | 1.09x | 1.88x | Contender |
| Long nested list | Markdig | 1,258.03 us | 1,754.82 KB |  |  |  |
| Normalization stress | OfficeIMO.Semantic | 1,493.25 us | 2,611.46 KB | 2.79x | 3.66x | Material gap |
| Normalization stress | Markdig | 535.61 us | 713.66 KB |  |  |  |
| Portable README | OfficeIMO.Semantic | 118.69 us | 237.59 KB | 1.79x | 2.82x | Material allocation gap |
| Portable README | Markdig | 66.45 us | 84.27 KB |  |  |  |
| Rich AST | OfficeIMO.Semantic | 187.96 us | 403.20 KB | 2.34x | 2.77x | Material gap |
| Rich AST | Markdig | 80.48 us | 145.49 KB |  |  |  |
| Technical document | OfficeIMO.Semantic | 108.81 us | 321.44 KB | 1.48x | 2.46x | Material allocation gap |
| Technical document | Markdig | 73.52 us | 130.55 KB |  |  |  |
| Transcript | OfficeIMO.Semantic | 92.96 us | 291.97 KB | 1.04x | 2.86x | Material allocation gap |
| Transcript | Markdig | 89.44 us | 101.96 KB |  |  |  |

Only the long nested-list corpus is currently inside the contender range in
both dimensions. Four other corpora meet the time ceiling but miss the
allocation ceiling. Rich AST and normalization stress remain material gaps in
both dimensions. These results do not close the
separate source-backed parser, which retains source spans, trivia, and syntax
ownership and needs its own performance contract.

The measured parser change avoids no-op reference-definition scans, redundant
paragraph construction, semantic-only source metadata, repeated nested option
clones, and general inline parsing for conservative flat CommonMark shapes. The
managed allocation reduction against clean commit
`6607a31f674f5f07191b8cc9ce9e4080264635d6` was 62.0% for Rich AST, 60.4% for
Technical document, 38.7% for Portable README, 11.3% for Transcript, and
3.2-5.6% for the other large synthetic corpora. The focused Rich AST follow-up
added a cheap per-line reference-definition marker guard and skipped multiline
inline-preservation construction when no candidate delimiter exists.

BenchmarkDotNet flagged seven full-matrix method distributions as bimodal or
multimodal and also flagged both Rich AST follow-up methods. It extended several
methods to 100 iterations. Allocation counts were stable, but close timing
ratios should be treated as directional until the runner gains a bounded
repeatability job. Peak process memory is not captured by this
BenchmarkDotNet lane. Output byte size is not a file-creation metric for parsing;
the accepted output contract is normalized semantic equivalence.

Reproduce the validation and full run with:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Markdown.Benchmarks -- --validate-equivalence
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload markdownparse -RunMode full -Framework net8.0
```

Raw BenchmarkDotNet output remains local and excluded from the repository's
small committed evidence set.
