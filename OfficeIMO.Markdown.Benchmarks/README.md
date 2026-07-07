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

- parse cost against the current Markdig baseline
- syntax-tree parse cost
- HTML render cost against the current Markdig baseline
- document normalization transform cost, including syntax-tree diagnostics
- HTML-to-Markdown conversion cost across OfficeIMO output profiles and the current ReverseMarkdown benchmark-only baseline

## Boundaries

- Benchmark scenarios belong here.
- Runtime Markdown behavior belongs in `OfficeIMO.Markdown`.
- Renderer host behavior belongs in `OfficeIMO.MarkdownRenderer`.
- ReverseMarkdown is a benchmark-only comparison package in this project and must not become a runtime dependency.
- Release decisions should use benchmark evidence together with correctness tests and representative document fixtures.
