# OfficeIMO.Data.Arrow

`OfficeIMO.Data.Arrow` converts any forward-only `DbDataReader`—including readers from
`OfficeIMO.Excel` and `OfficeIMO.CSV`—into bounded Apache Arrow record batches.

```powershell
dotnet add package OfficeIMO.Data.Arrow
```

```csharp
using Apache.Arrow;
using OfficeIMO.Data.Arrow;
using OfficeIMO.Excel;
using System.Data.Common;

using DbDataReader reader = ExcelDocument.OpenDataReader("report.xlsx");

await foreach (RecordBatch batch in reader.ReadArrowBatchesAsync(
    new ArrowReadOptions { BatchSize = 65_536 },
    cancellationToken)) {
    // Send the batch to an analytics, IPC, or columnar processing pipeline.
}
```

The adapter does not load the complete input into an Arrow table. Batch size and supported
fallback behavior are controlled through `ArrowReadOptions`. Native adapters cover Boolean,
integer, floating-point, decimal, timestamp, `DateOnly`, `TimeOnly`, `Guid`, binary, and text
columns. Unsupported CLR types are converted to invariant text by default; set
`ConvertUnsupportedTypesToString` to `false` when the pipeline should reject them instead.
CLR `DateTime` columns become timezone-less Arrow timestamps so spreadsheet and database
wall-clock values retain their original meaning. `DateTimeOffset` columns become UTC-aware
timestamps and preserve their instant.
