# OfficeIMO.Markdown.Benchmarks

`OfficeIMO.Markdown.Benchmarks` contains benchmark and comparison workloads for the Markdown builder, reader, renderer, and related conversion paths. It is not a NuGet-facing runtime package.

## Use

Run benchmarks from the repository root with the repo's normal .NET SDK:

```powershell
dotnet run --project OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj -c Release
```

## Boundaries

- Benchmark scenarios belong here.
- Runtime Markdown behavior belongs in `OfficeIMO.Markdown`.
- Renderer host behavior belongs in `OfficeIMO.MarkdownRenderer`.
- Release decisions should use benchmark evidence together with correctness tests and representative document fixtures.
