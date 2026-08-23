# OfficeIMO.Markdown.Benchmarks

`OfficeIMO.Markdown.Benchmarks` contains benchmark and comparison workloads for the Markdown builder, reader, renderer, and related conversion paths. It is not a NuGet-facing runtime package.

This opt-in project is intentionally outside `OfficeIMO.sln`. Markdig and
ReverseMarkdown remain comparison-only dependencies and do not enter normal
solution restore, build, or runtime packages.

## Use

Run benchmarks from the repository root with the repo's normal .NET SDK:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0
```

Run a narrower benchmark by class when you only need one lane:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *MarkdownTransformBenchmarks*
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlToMarkdownBenchmarks*
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlToMarkdownOfficeProfileBenchmarks*
```

Select one Markdown corpus without changing the benchmark source by setting
`OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS` for that process:

```powershell
$env:OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS = 'LongNestedList'
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *MarkdownParseBenchmarks.*_ParseSemantic*_CommonMark
Remove-Item Env:OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS
```

For a quick harness smoke without publication-grade timing, use BenchmarkDotNet's dry job:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlToMarkdownBenchmarks* --job Dry --noOverwrite
```

## Corpus

The benchmark corpus is intentionally stable and reviewable in source. It covers README-style docs, chat/transcript documents, technical docs, mixed AST-heavy content, long nested lists, large pipe tables, and normalization-heavy transcript artifacts.

Benchmark classes currently cover:

- semantic parse cost across OfficeIMO and the configured comparison implementation
- syntax-tree parse cost
- HTML render cost across the configured implementations
- document normalization transform cost, including syntax-tree diagnostics
- comparable HTML-to-Markdown default conversion cost for OfficeIMO and the current ReverseMarkdown benchmark-only baseline
- OfficeIMO-specific HTML-to-Markdown profile costs in a separate, non-competitive benchmark class

The competitive CommonMark parse lane measures `MarkdownReader.ParseSemantic`, which builds OfficeIMO's typed semantic document without source spans, trivia, or a syntax tree. Markdig builds its normal document tree. This is an equivalent rendered-output comparison, not a claim that the two libraries retain identical source metadata. The same class reports OfficeIMO's source-backed `Parse` and `ParseWithSyntaxTree` methods as separate lanes so their additional cost remains visible.

The CommonMark parse and HTML comparison classes run an untimed setup preflight for every corpus. Both compared paths render without automatic heading identifiers or an outer body wrapper, matching CommonMark's heading and fragment output rather than hiding different work during validation. The preflight parses with both OfficeIMO and Markdig, renders both results, normalizes line endings, equivalent break-tag spelling, and HTML-collapsible whitespace while preserving preformatted blocks, and rejects the benchmark case unless the remaining HTML is identical.

The competitive HTML-to-Markdown class likewise checks every included corpus before timing. It renders both generated Markdown results through the same pipeline and requires identical normalized semantic HTML. Both timed methods accept the same raw HTML and include parsing; OfficeIMO does not reuse a document parsed during setup while ReverseMarkdown parses inside its measured call. Nested-list and feature-rich inputs remain in the OfficeIMO-only prepared-document profile class because those paths produce structurally different output and would make cross-library ratios misleading. Timings therefore begin only after compared implementations have proven the same observable output for that input.

Run the same preflight without collecting timings:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release -f net8.0 -- --validate-equivalence
```

The benchmark classes do not hard-code a runtime job. The selected target
framework controls the runtime, while `--job Dry` and `--job Short` select the
requested execution policy without also running an implicit full job.

Run the equivalent comparison lanes through the repository's shared evidence
runner when provenance and normalized JSON/CSV/Markdown output are needed:

```powershell
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload markdownparse -RunMode full -Framework net8.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload markdownhtml -RunMode full -Framework net8.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload htmltomarkdown -RunMode full -Framework net8.0
```

## Interpretation

Use benchmark results together with correctness tests and representative document fixtures. Timing alone does not establish syntax coverage, output fidelity, or safe handling of untrusted input.
