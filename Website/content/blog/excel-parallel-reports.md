---
title: "Building Excel Reports with Parallel Compute"
description: "Learn how OfficeIMO.Excel uses a parallel threading model to accelerate bulk cell writes, AutoFit, and formula recalculation for large workbooks."
date: 2025-03-20
tags: [excel, performance, threading]
categories: [Tutorial, Deep Dive]
author: "Przemyslaw Klys"
---

When your reporting pipeline needs to fill a workbook with tens of thousands of rows, single-threaded cell writes can become a bottleneck. **OfficeIMO.Excel** includes a parallel compute layer that can use multiple cores for bulk-oriented workloads such as row writes and column AutoFit.

## The Problem with Large Workbooks

A typical financial report might contain tens of thousands of rows across multiple sheets. Writing each cell sequentially, formatting it, and then running AutoFit on every column does not always make good use of the hardware available on CI runners or application servers. Parallel execution is one way to reduce that bottleneck when the workbook shape is a good fit.

## Enabling Parallel Mode

OfficeIMO.Excel exposes a `ParallelOptions` property on the workbook:

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("BigReport.xlsx");

workbook.ParallelOptions = new ExcelParallelOptions
{
    MaxDegreeOfParallelism = Environment.ProcessorCount,
    MinRowsPerPartition = 1000
};
```

Setting `MinRowsPerPartition` prevents the scheduler from creating too many tiny work items, which would hurt performance through overhead.

## Bulk Cell Writes

The `WriteBulk` API accepts a two-dimensional array and distributes row ranges across threads:

```csharp
var sheet = workbook.AddSheet("Sales");

// Build the data
var data = new object[50_000, 5];
for (int r = 0; r < 50_000; r++)
{
    data[r, 0] = $"SKU-{r:D6}";
    data[r, 1] = Random.Shared.Next(1, 500);
    data[r, 2] = Math.Round(Random.Shared.NextDouble() * 99.99, 2);
    data[r, 3] = $"=B{r + 2}*C{r + 2}";
    data[r, 4] = DateTime.Today.AddDays(-Random.Shared.Next(0, 365));
}

// Headers
sheet.SetRow(0, new object[] { "SKU", "Quantity", "Unit Price", "Total", "Date" });

// Parallel bulk write starting at row 1
sheet.WriteBulk(data, startRow: 1);
```

Internally, `WriteBulk` partitions the row range, and each partition writes directly to a separate region of the underlying XML. Because Open XML stores sheet data in row order, there is no contention between threads.

## Parallel AutoFit

Column width calculation requires measuring the rendered width of every cell in a column. OfficeIMO parallelises this per-column:

```csharp
sheet.AutoFitColumns(); // uses workbook.ParallelOptions automatically
```

The exact benefit depends on the number of rows, the width of the sheet, the host machine, and the fonts involved in width calculation.

## What To Measure

The benefit of parallel mode depends on workbook shape, formula density, formatting work, and the number of cores available. In practice, the biggest wins usually show up when you:

- write large rectangular datasets
- run `AutoFitColumns()` across wide or busy sheets
- generate several independent sheets in one job

Treat the feature as something to benchmark on your own workload rather than assuming one universal speedup number.

## Thread Safety Notes

Each sheet maintains its own write lock, so you can safely populate different sheets from different threads as well:

```csharp
Parallel.ForEach(sheetConfigs, config =>
{
    var sheet = workbook.AddSheet(config.Name);
    sheet.WriteBulk(config.Data, startRow: 1);
    sheet.AutoFitColumns();
});
```

Just be sure to call `AddSheet` inside a lock or pre-create the sheets before entering the parallel loop, since the sheet collection itself is not concurrent.

## Conclusion

Parallel compute in OfficeIMO.Excel can help when your report generation workload is large enough to benefit from partitioning. Enable it, test it against your real workbook shapes, and keep the sequential path as the baseline for comparison.

## Continue with

- [OfficeIMO.Excel](/products/excel/) for the package overview and supported workflow surface.
- [Excel documentation](/docs/excel/) for sheets, tables, formulas, and formatting patterns.
- [Tables and ranges guide](/docs/excel/tables-ranges/) for more structured workbook layouts.
- [Benchmarks](/benchmarks/) for the current site-wide benchmark summary and caveats.
