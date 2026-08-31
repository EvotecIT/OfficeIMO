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

using DbDataReader reader = ExcelDocument.OpenDataReader(
    "report.xlsx",
    new ExcelReadOptions { InferSchema = true });

await foreach (RecordBatch batch in reader.ReadArrowBatchesAsync(
    new ArrowReadOptions { BatchSize = 65_536 },
    cancellationToken)) {
    using (batch) {
        // Send the batch to an analytics, IPC, or columnar processing pipeline.
    }
}
```

The adapter does not accumulate the complete input in an Arrow table. Batch size and supported
fallback behavior are controlled through `ArrowReadOptions`; the source reader retains ownership
of its own buffering and memory policy. Native adapters cover Boolean,
integer, floating-point, decimal, timestamp, `DateOnly`, `TimeOnly`, `Guid`, binary, and text
columns. Unsupported CLR types are converted to invariant text by default; set
`ConvertUnsupportedTypesToString` to `false` when the pipeline should reject them instead.
Decimal values must be exactly representable at `ArrowReadOptions.DecimalScale`; conversion
fails instead of silently rounding an inexact value. Increase the scale when the source contains
more significant fractional digits.
When the source schema is already known, set `ArrowReadOptions.ColumnTypes` in ordinal order and
leave reader-side inference disabled. The adapter snapshots and validates the explicit types before
reading, so the conversion does not pay a schema-sampling pass.
CLR `DateTime` columns become timezone-less Arrow timestamps so spreadsheet and database
wall-clock values retain their original meaning. `DateTimeOffset` columns become UTC-aware
timestamps and preserve their instant. Temporal columns use nanoseconds by default so CLR
100-nanosecond precision is retained. Set `ArrowReadOptions.TemporalUnit` to
`TimeUnit.Microsecond` when the wider microsecond timestamp range is required and accepting
sub-microsecond precision loss is appropriate.

## Managed and C streams

For consumers that expect Apache Arrow's stream contract, open a managed
`IArrowArrayStream`. It keeps the same bounded-batch behavior and leaves the source reader
caller-owned:

```csharp
using Apache.Arrow.Ipc;

using IArrowArrayStream stream = reader.OpenArrowStream(
    new ArrowReadOptions { BatchSize = 16_384 });

while (await stream.ReadNextRecordBatchAsync(cancellationToken) is { } batch) {
    using (batch) {
        // Consume one bounded batch.
    }
}
```

Native engines can consume the same pipeline through the Arrow C Data Interface:

```csharp
using ArrowCArrayStreamOwner stream = reader.ExportArrowCStream(
    new ArrowReadOptions { BatchSize = 16_384 });

nint address = stream.Address;
// Pass address to a native ArrowArrayStream consumer while stream remains alive.
```

`ArrowCArrayStreamOwner` owns both the unmanaged stream struct and its managed callbacks.
Keep it alive for the complete native call and dispose it afterwards. Native code may invoke
the stream's release callback, but it must not free the struct allocation. Each `get_next`
call produces at most the configured batch size; the full worksheet or CSV is never collected
into one `RecordBatch`.
