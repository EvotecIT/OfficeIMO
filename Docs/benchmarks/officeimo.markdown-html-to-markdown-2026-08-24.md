# HTML-to-Markdown comparison evidence

This 2026-08-24 run compares complete HTML-to-Markdown conversion through
OfficeIMO.Markdown and ReverseMarkdown. Both timed lanes accept the same raw
HTML and include HTML parsing. Before measurement, the validation runner renders
both Markdown results through the same Markdown pipeline and requires identical
normalized semantic HTML for every corpus.

The performance acceptance bands used for this work are:

- at or below 2x the comparison implementation in both mean time and managed
  allocation: contender range;
- above 2x through 5x in either dimension: material remediation gap;
- above 5x in either dimension: unacceptable unless the observable contracts
  differ and the result is reported as a non-equivalent diagnostic.

This was a full BenchmarkDotNet run from clean source commit
`c0fcf3abb1fd5d352dd0a41f7a09182233db8787` on Windows 11
10.0.26200.9168, an AMD Ryzen 9 9950X3D2 with 16 physical cores, .NET SDK
10.0.111, and .NET 10.0.11 x64. Allocation is managed allocation per
operation.

| Corpus | Implementation | Mean | Allocation | OfficeIMO time ratio | OfficeIMO allocation ratio |
| --- | --- | ---: | ---: | ---: | ---: |
| Article | OfficeIMO.Markdown | 38.25 us | 82.93 KB | 1.50x | 1.92x |
| Article | ReverseMarkdown | 25.49 us | 43.27 KB |  |  |
| Large article | OfficeIMO.Markdown | 2,066.55 us | 3,212.66 KB | 1.39x | 1.86x |
| Large article | ReverseMarkdown | 1,490.23 us | 1,730.53 KB |  |  |
| Table | OfficeIMO.Markdown | 3,249.50 us | 4,934.08 KB | 1.64x | 1.78x |
| Table | ReverseMarkdown | 1,983.47 us | 2,764.71 KB |  |  |

All three corpora are inside the contender range in both dimensions. This is a
bounded conclusion for the validated HTML-to-Markdown contract; it does not
make the separate source-backed Markdown parser competitive. That parser
retains source spans, trivia, and syntax ownership and remains an active
optimization target.

Text byte count is not used as a package-size metric here. The implementations
may choose different Markdown spellings while preserving the validated semantic
result. Raw BenchmarkDotNet output remains local and excluded from the
repository's small committed evidence set.

Reproduce the validation and full run with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markdown.Benchmarks -- --validate-equivalence
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload htmltomarkdown -RunMode full -Framework net10.0
```
