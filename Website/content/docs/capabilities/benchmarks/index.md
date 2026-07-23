---
title: "Benchmark evidence and reproduction"
description: "Run OfficeIMO comparison benchmarks, regression baselines, and performance guardrails with the same validation contracts used for published evidence."
meta.seo_title: "OfficeIMO .NET benchmark evidence and reproduction guide"
order: 25
---

OfficeIMO uses three kinds of performance evidence. A **comparison benchmark** measures equivalent work across libraries and validates the resulting files or data. A **regression baseline** records an OfficeIMO workflow so future changes can be compared with the same scenario. A **performance guardrail** fails when a representative workload exceeds a documented time, allocation, memory, or I/O budget.

Those categories answer different questions. A regression result is not presented as a competitor ranking, and a timing from one machine is not a promise for every environment.

## Published comparisons

The [public benchmark page](/benchmarks/) contains the committed Excel and CSV comparison snapshots.

### Excel reports and data pipelines

The Excel suite covers 25,000-row creation, `IDataReader` writes, typed reads, styling, formulas, tables, and charts. Result artifacts are validated before a measurement is accepted.

```shell
dotnet run -c Release --project OfficeIMO.Excel.Benchmarks -- --help
```

Use the scenario links on the [benchmark page](/benchmarks/#excel-evidence) to inspect the recorded runtime, machine, workload, and result matrix.

### CSV reads and writes

The CSV suite traverses every field for read scenarios and validates the semantic output of write scenarios. This prevents a fast partial read or incomplete file from being counted as equivalent work.

```shell
dotnet run -c Release --project OfficeIMO.CSV.Benchmarks -- --filter *CsvWideBenchmarks*
```

The committed comparison is available on the [benchmark page](/benchmarks/#csv-evidence).

## Reader regression baseline

`OfficeIMO.Reader.Benchmarks` measures format detection, extraction, chunking, and transport across a mixed document corpus. The committed baseline records 25 cases across 14 formats with the runtime and machine context retained.

```shell
dotnet run -c Release --project OfficeIMO.Reader.Benchmarks -- evidence --help
```

[Inspect the committed Reader evidence](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/benchmarks/officeimo.reader.foundation-2026-07-10.md) before comparing it with a local run.

## Performance guardrails

PDF and RTF expose explicit budget-verification modes:

```shell
dotnet run -c Release --project OfficeIMO.Pdf.Benchmarks -- --verify-budgets
dotnet run -c Release --project OfficeIMO.Rtf.Benchmarks -- --verify-budgets
```

Email performance tests cover representative MIME, MSG, and mbox workloads. They assert both the budget and the workload envelope, such as source size or message count, so a smaller fixture cannot accidentally make the test pass.

```shell
dotnet test OfficeIMO.Email.Tests -c Release --filter FullyQualifiedName~EmailPerformanceEvidenceTests
```

[Read the Email performance contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.email-performance.md) for the fixture sizes, assertions, and environment controls.

## Additional benchmark projects

The repository also contains dedicated projects for Markdown, HTML, OneNote, OpenDocument, and drawing workloads:

- [`OfficeIMO.Markdown.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Markdown.Benchmarks)
- [`OfficeIMO.Html.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Html.Benchmarks)
- [`OfficeIMO.OneNote.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.OneNote.Benchmarks)
- [`OfficeIMO.OpenDocument.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.OpenDocument.Benchmarks)
- [`OfficeIMO.Drawing.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Drawing.Benchmarks)

Run a project with `--help` first. Its command surface is the source of truth for filters and output options.

## Word and PowerPoint

Word and PowerPoint have performance-focused tests for known operations, but the repository does not currently contain a publication-grade cross-library comparison artifact for either family. Their tests are useful regression protection; they are not used to claim that OfficeIMO is faster than another library.

For evaluation, start with the [Word production workflows](/docs/word/market-readiness/) and [PowerPoint designer guide](/docs/powerpoint/designer/). Those pages lead to runnable examples and validation proof while the performance evidence remains scoped to what is actually committed.

## Reproduce results responsibly

1. Use a Release build and record the .NET runtime, operating system, processor, and available memory.
2. Keep the input shape, row or item count, enabled features, and validation work equivalent.
3. Run on an otherwise quiet machine and retain warmup and measured iteration counts.
4. Inspect the generated artifact or semantic validation result, not only elapsed time.
5. Compare local results with the same suite and contract. Do not compare unrelated microbenchmarks.

For package-level evaluation, install the released package you intend to deploy and run the same workload in an environment close to production.
