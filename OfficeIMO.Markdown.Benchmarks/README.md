# OfficeIMO.Markdown.Benchmarks

Internal benchmark harness for `OfficeIMO.Markdown`.

It measures representative parse and HTML-render workloads for:

- OfficeIMO default reader behavior
- OfficeIMO portable reader profile
- the internal comparison baseline used in parity work

Run with:

```powershell
dotnet run -c Release --project .\OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj
```

Filter a specific benchmark class with:

```powershell
dotnet run -c Release --project .\OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj -- --filter *MarkdownParseBenchmarks*
```
