# OfficeIMO.Markdown.Benchmarks

`OfficeIMO.Markdown.Benchmarks` contains benchmark and comparison workloads for the Markdown builder, reader, renderer, and related conversion paths. It is not a NuGet-facing runtime package.

## Use

Run benchmarks from the repository root with the repo's normal .NET SDK:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0
```

Run a narrower benchmark by class when you only need one lane:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *MarkdownTransformBenchmarks*
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlToMarkdownBenchmarks*
```

For a quick harness smoke without publication-grade timing, use BenchmarkDotNet's dry job:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlToMarkdownBenchmarks* --job Dry --warmupCount 1 --iterationCount 1
```

## Corpus

The benchmark corpus is intentionally stable and reviewable in source. It covers README-style docs, chat/transcript documents, technical docs, mixed AST-heavy content, long nested lists, large pipe tables, and normalization-heavy transcript artifacts.

Benchmark classes currently cover:

- parse cost across the configured implementations
- syntax-tree parse cost
- HTML render cost across the configured implementations
- document normalization transform cost, including syntax-tree diagnostics
- HTML-to-Markdown conversion cost across OfficeIMO output profiles and the current ReverseMarkdown benchmark-only baseline

The CommonMark parse and HTML comparison classes run an untimed setup preflight for every corpus. Both measured HTML paths render without automatic heading identifiers or an outer body wrapper, matching CommonMark's heading and fragment output rather than hiding different work during validation. The preflight parses with both OfficeIMO and Markdig, renders both results, normalizes line endings, equivalent break-tag spelling, and HTML-collapsible whitespace while preserving preformatted blocks, and rejects the benchmark case unless the remaining HTML is identical. Timings therefore begin only after both implementations have proven the same observable output for that input.

Run the same preflight without collecting timings:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --validate-equivalence
```

## Interpretation

Use benchmark results together with correctness tests and representative document fixtures. Timing alone does not establish syntax coverage, output fidelity, or safe handling of untrusted input.
