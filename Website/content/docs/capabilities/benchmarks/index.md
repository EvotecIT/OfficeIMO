---
title: "Benchmark evidence and reproduction"
description: "Run OfficeIMO comparison benchmarks, regression baselines, and performance guardrails with the same validation contracts used for published evidence."
meta.seo_title: "OfficeIMO .NET benchmark evidence and reproduction guide"
order: 25
---

OfficeIMO uses three kinds of performance evidence. A **comparison benchmark** measures equivalent work across libraries and validates the resulting files or data. A **regression baseline** records an OfficeIMO workflow so later changes can be compared with the same scenario. A **performance guardrail** fails when a representative workload exceeds a documented time, allocation, memory, or I/O budget.

Those categories answer different questions. A regression result is not presented as a competitor ranking, and a timing from one machine is not a promise for every environment.

## Published comparisons

The [public benchmark page](/benchmarks/) separates three evidence layers:

- **Source suites** show which workloads the current repository can run. A project in the tree is coverage, not a performance result.
- **Cataloged artifacts** show measured results with a source commit, runtime, hardware, operating system, run mode, and validation output. Only full artifacts marked for publication support public comparisons; diagnostic artifacts only prove that a lane executes and validates.
- **Historical engineering evidence** preserves the larger Excel and CSV matrix. It remains useful for scenario and library coverage, but rows without complete environment provenance are not mixed into current platform rankings.

The page keeps missing workload and platform lanes visible. That makes the evidence boundary inspectable instead of quietly reducing the comparison to the few cells that happen to be available.

### Excel reports and data pipelines

The Excel suite covers 25,000-row creation, `IDataReader` writes, typed reads, styling, formulas, tables, and charts. Result artifacts are validated before a measurement is accepted.

```shell
dotnet run -c Release --project OfficeIMO.Excel.Benchmarks -- --help
```

Use the selector on the [benchmark page](/benchmarks/#library-comparison-evidence) to inspect the measured source revision, runtime, hardware, operating system, run mode, workload, and result rows. Use the historical matrix on the same page when you need the broader scenario and library inventory.

### CSV reads and writes

The CSV suite traverses every field for read scenarios and validates the semantic output of write scenarios. This prevents a fast partial read or incomplete file from being counted as equivalent work.

```shell
dotnet run -c Release --project OfficeIMO.CSV.Benchmarks -- --filter *CsvWideBenchmarks*
```

The committed comparison is available on the [benchmark page](/benchmarks/#library-comparison-evidence). Full and diagnostic artifacts are labeled separately and are never combined into one ranking.

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

The PDF budget gate uses deterministic allocation, retained-memory, output, and
cached-allocation-savings contracts plus generous elapsed-time ceilings in
ordinary CI. Run it with `--verify-timing-budgets` on a controlled benchmark
host to additionally enforce the relative cached speedup target.

Email performance tests cover representative MIME, MSG, and mbox workloads. They assert both the budget and the workload envelope, such as source size or message count, so a smaller fixture cannot accidentally make the test pass.

```shell
dotnet test OfficeIMO.Email.Tests -c Release --filter FullyQualifiedName~EmailPerformanceEvidenceTests
```

[Read the Email performance contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.email-performance.md) for the fixture sizes, assertions, and environment controls.

## Current source-suite inventory

The repository contains 16 benchmark projects. Some are comparison adapters, while others are OfficeIMO regression or guardrail suites:

- Data and document workloads: [`OfficeIMO.CSV.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.CSV.Benchmarks), [`OfficeIMO.Excel.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Excel.Benchmarks), [`OfficeIMO.Reader.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Reader.Benchmarks), [`OfficeIMO.Word.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Word.Benchmarks), and [`OfficeIMO.PowerPoint.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.PowerPoint.Benchmarks).
- Publishing and interchange workloads: [`OfficeIMO.Pdf.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Pdf.Benchmarks), [`OfficeIMO.Rtf.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Rtf.Benchmarks), [`OfficeIMO.Markdown.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Markdown.Benchmarks), [`OfficeIMO.Html.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Html.Benchmarks), [`OfficeIMO.OneNote.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.OneNote.Benchmarks), [`OfficeIMO.OpenDocument.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.OpenDocument.Benchmarks), and [`OfficeIMO.Drawing.CodeGlyphX.Benchmarks`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Drawing.CodeGlyphX.Benchmarks).
- Cross-library adapters: [`OfficeIMO.Excel.Benchmarks.LegacyEpPlus`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Excel.Benchmarks.LegacyEpPlus), [`OfficeIMO.Excel.Benchmarks.NPOI`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Excel.Benchmarks.NPOI), [`OfficeIMO.Pdf.Benchmarks.Comparisons`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Pdf.Benchmarks.Comparisons), and [`OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler`](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler).

This inventory describes current source, independent of a package release. Run a project with `--help` first; its checked-in command surface is the source of truth for filters and output options. A suite belongs in the public comparison selector only after a validated, provenance-complete result artifact has been committed to its catalog.

## Word and PowerPoint

Word and PowerPoint have dedicated benchmark projects and performance-focused tests for known operations, but the website does not currently contain a publication-grade cross-library result artifact for either family. Their source suites are useful regression coverage; they are not used to claim that OfficeIMO is faster than another library.

For evaluation, start with the [Word production workflows](/docs/word/market-readiness/) and [PowerPoint designer guide](/docs/powerpoint/designer/). Those pages lead to runnable examples and validation proof while the performance evidence remains scoped to what is actually committed.

## Reproduce results responsibly

1. Use a Release build and record the .NET runtime, operating system, processor, and available memory.
2. Keep the input shape, row or item count, enabled features, and validation work equivalent.
3. Run on an otherwise quiet machine and retain warmup and measured iteration counts.
4. Inspect the generated artifact or semantic validation result, not only elapsed time.
5. Compare local results with the same suite and contract. Do not compare unrelated microbenchmarks.

For deployment evaluation, build the source revision you intend to assess—or install the package you intend to deploy—and run the same workload in an environment close to production. Record that revision or package version with the result rather than assuming the website and a released package represent the same code.
