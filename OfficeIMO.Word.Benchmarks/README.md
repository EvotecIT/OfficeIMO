# OfficeIMO.Word benchmarks

This project measures allocation-sensitive Word workflows over deterministic generated DOCX inputs. The 100- and 1,000-item parameters make scaling regressions visible without introducing external corpus or desktop-Office dependencies.

```bash
dotnet run -c Release -f net8.0 --project OfficeIMO.Word.Benchmarks -- --filter '*WordWorkflowBenchmarks*'
```

Use `--job Dry` for a quick execution check. A dry run is not a stable performance baseline; use the normal BenchmarkDotNet job on an otherwise idle machine before comparing timings or setting budgets.

The suite covers package load, field refresh, mail merge, structured comparison, Word-to-HTML including package load, and Word-to-HTML over an already loaded document. Global setup executes each workflow once and rejects unexpected field, merge, paragraph, HTML-output, or comparison results before BenchmarkDotNet starts timing. Temporary packages are created under the operating-system temporary folder and removed by benchmark cleanup.
