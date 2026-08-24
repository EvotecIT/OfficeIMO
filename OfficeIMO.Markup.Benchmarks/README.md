# OfficeIMO Markup benchmarks

This opt-in BenchmarkDotNet project measures the Markdown-compatible semantic
parse workflow in `OfficeIMO.Markup` against Markdig. Both lanes use CommonMark
plus pipe tables. The corpus is deliberately limited to headings, paragraphs,
formatted inline text, lists, and pipe tables supported by both implementations.

Every run first builds a canonical semantic event stream for each AST and
requires identical event counts and SHA-256 digests. A timing result is invalid
if the two parsers do not produce the same semantic projection.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- --filter '*OfficeMarkupParseBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- --evidence --repeat 3 --json .\artifacts\markup-evidence.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Markup.Benchmarks -- --verify-budgets --repeat 3
```

The evidence runner uses a fresh child process for every engine, scale, and
repeat. It records elapsed time and managed allocation per document, retained
managed heap growth, sampled managed-heap and working-set peaks, absolute
process peak working set, input bytes, semantic event count and digest, runtime,
OS, architecture, commit, and dirty-tree state. Parsing does not create a file,
so output byte size is not an applicable metric for this suite; the validated
semantic event stream is its output contract.

The checked-in budget keeps elapsed and allocation ratios below 1.75x. The
broader project acceptance boundary remains 2x, which is a contender threshold,
not an optimality claim.

Office-specific directives are excluded from the comparison because Markdig
does not produce the corresponding Word, workbook, or presentation semantic
model. They require a separate OfficeIMO-only evidence lane.
