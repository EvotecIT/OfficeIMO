# OfficeIMO.OneNote benchmarks

This permanent BenchmarkDotNet suite tracks how native desktop `.one` reading, writing, and semantic Markdown projection scale from one page to a representative multi-page section. The useful comparison is input scale and later before/after runs; exact timings vary by machine.

Validate the suite quickly:

```powershell
dotnet run -c Release --framework net8.0 -- --filter "*OneNoteReadWriteBenchmarks*" --job Dry --noOverwrite
```

For measurements, use `--job Short` while iterating and the default job for recorded release evidence. Keep benchmark artifacts outside the repository or delete them after summarizing the result.
