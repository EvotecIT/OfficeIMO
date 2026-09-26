using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using Apache.Arrow;
using Apache.Arrow.C;
using Apache.Arrow.Ipc;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using OfficeIMO.Data.Arrow;

CsvDocument document = CsvDocument.Parse("Name,Score,Date,Time\nAlice,42,2026-08-06,14:35:12\n");
if (!document.Header.SequenceEqual(new[] { "Name", "Score", "Date", "Time" })) {
    throw new InvalidOperationException("The CSV parser did not preserve the header schema.");
}

CsvSmokeRow mapped = document.RowsAs<CsvSmokeRow>().Single();
if (mapped.Name != "Alice" || mapped.Score != 42 ||
    mapped.Date != new DateOnly(2026, 8, 6) || mapped.Time != new TimeOnly(14, 35, 12)) {
    throw new InvalidOperationException("The CSV typed-row mapper lost a value.");
}

CsvSmokeRow explicitlyMapped = document.RowsAs<CsvSmokeRow>(mapper => mapper
    .FromColumn<string>("Name", static (row, value) => { row.Name = value; return row; })
    .FromColumn<int>("Score", static (row, value) => { row.Score = value; return row; })
    .FromColumn<DateOnly>("Date", static (row, value) => { row.Date = value; return row; })
    .FromColumn<TimeOnly>("Time", static (row, value) => { row.Time = value; return row; }))
    .Single();
if (explicitlyMapped.Name != "Alice" || explicitlyMapped.Score != 42 ||
    explicitlyMapped.Date != new DateOnly(2026, 8, 6) || explicitlyMapped.Time != new TimeOnly(14, 35, 12)) {
    throw new InvalidOperationException("The explicit CSV typed-row mapper lost a value.");
}

using var reader = CsvDocument.OpenTextDataReader("Name,Score,Date,Time\nAlice,42,2026-08-06,14:35:12\n");
CsvSmokeRow streamed = reader.RowsAs<CsvSmokeRow>().Single();
if (streamed.Name != "Alice" || streamed.Score != 42 ||
    streamed.Date != new DateOnly(2026, 8, 6) || streamed.Time != new TimeOnly(14, 35, 12)) {
    throw new InvalidOperationException("The forward-only CSV typed-row mapper lost a value.");
}

using var explicitReader = CsvDocument.OpenTextDataReader("Name,Score,Date,Time\nAlice,42,2026-08-06,14:35:12\n");
CsvSmokeRow explicitlyStreamed = explicitReader.RowsAs<CsvSmokeRow>(mapper => mapper
    .FromColumn<string>("Name", static (row, value) => { row.Name = value; return row; })
    .FromColumn<int>("Score", static (row, value) => { row.Score = value; return row; })
    .FromColumn<DateOnly>("Date", static (row, value) => { row.Date = value; return row; })
    .FromColumn<TimeOnly>("Time", static (row, value) => { row.Time = value; return row; }))
    .Single();
if (explicitlyStreamed.Name != "Alice" || explicitlyStreamed.Score != 42 ||
    explicitlyStreamed.Date != new DateOnly(2026, 8, 6) || explicitlyStreamed.Time != new TimeOnly(14, 35, 12)) {
    throw new InvalidOperationException("The explicit forward-only typed-row mapper lost a value.");
}

CsvSchema parallelSchema = new CsvSchemaBuilder()
    .Column("Name").AsString()
    .Column("Score").AsInt32()
    .Column("Date").AsType(typeof(DateOnly))
    .Column("Time").AsType(typeof(TimeOnly))
    .Done()
    .Build();
using (var parallelReader = CsvDocument.OpenTextDataReader(
           "Name,Score,Date,Time\nAlice,42,2026-08-06,14:35:12\nBob,43,2026-08-07,15:36:13\n",
           readerOptions: new CsvDataReaderOptions {
               Schema = parallelSchema,
               ParallelProcessing = new CsvDataReaderParallelOptions {
                   MaxDegreeOfParallelism = 2,
                   BatchSize = 1
               }
           })) {
    int score = 0;
    while (parallelReader.Read()) score += parallelReader.GetInt32(1);
    if (score != 85) {
        throw new InvalidOperationException("The AOT-safe parallel CSV data reader lost a value.");
    }
}

var exportedRows = new DataTable();
exportedRows.Columns.Add("Name", typeof(string));
exportedRows.Columns.Add("Score", typeof(decimal));
exportedRows.Columns.Add("Created", typeof(DateTime));
exportedRows.Columns.Add("Enabled", typeof(bool));
exportedRows.Rows.Add("Alice", 42.5m, new DateTime(2026, 8, 6, 14, 35, 12, DateTimeKind.Utc), true);
using (var exportedReader = exportedRows.CreateDataReader())
using (var exportedText = new StringWriter(CultureInfo.InvariantCulture)) {
    CsvDocument.WriteDataReader(
        exportedText,
        exportedReader,
        new CsvSaveOptions { NewLine = "\n", DateTimeFormat = "yyyy-MM-dd HH:mm:ss" });
    if (exportedText.ToString() != "Name,Score,Created,Enabled\nAlice,42.5,2026-08-06 14:35:12,True\n") {
        throw new InvalidOperationException("The AOT-safe typed DataReader writer lost a value.");
    }
}

using (var parallelExportedReader = exportedRows.CreateDataReader())
using (var parallelExportedText = new StringWriter(CultureInfo.InvariantCulture)) {
    CsvDocument.WriteDataReaderParallel(
        parallelExportedText,
        parallelExportedReader,
        new CsvSaveOptions { NewLine = "\n", DateTimeFormat = "yyyy-MM-dd HH:mm:ss" },
        new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 });
    if (parallelExportedText.ToString() != "Name,Score,Created,Enabled\nAlice,42.5,2026-08-06 14:35:12,True\n") {
        throw new InvalidOperationException("The AOT-safe parallel DataReader writer lost a value.");
    }
}

using var arrowReader = CsvDocument.OpenTextDataReader(
    "Id,Name\n1,Ada\n2,Grace\n",
    readerOptions: new CsvDataReaderOptions { InferSchema = true });
using ArrowCArrayStreamOwner arrowOwner = arrowReader.ExportArrowCStream(
    new ArrowReadOptions { BatchSize = 1 });
IArrowArrayStream importedArrowStream = arrowOwner.ImportArrayStream();
using (importedArrowStream) {
    using RecordBatch importedArrowBatch =
        (await importedArrowStream.ReadNextRecordBatchAsync())!;
    if (importedArrowBatch.Length != 1
        || importedArrowBatch.Schema.FieldsList.Count != 2
        || importedArrowBatch.Column(0).Length != 1) {
        throw new InvalidOperationException("The bounded Arrow C stream did not survive NativeAOT.");
    }
}

using var nativeArrowReader = CsvDocument.OpenTextDataReader(
    "Id,Name\n1,Ada\n2,Grace\n",
    readerOptions: new CsvDataReaderOptions { InferSchema = true });
using ArrowCArrayStreamOwner nativeArrowOwner = nativeArrowReader.ExportArrowCStream(
    new ArrowReadOptions { BatchSize = 1 });
using (ArrowCArrayStreamOwner.ArrowCArrayStreamLease nativeLease = nativeArrowOwner.AcquireLease()) {
    ConsumeArrowAbiDirectly(nativeLease.Address, expectedColumns: 2, expectedRows: 2, maximumBatchRows: 1);
}

Console.WriteLine("PASS | CSV parse, schema inspection, mapping, managed import and direct Arrow C ABI consumption, and sequential/parallel typed DataReader writing");

static unsafe void ConsumeArrowAbiDirectly(nint address, long expectedColumns, long expectedRows, long maximumBatchRows) {
    var stream = (NativeArrowArrayStream*)address;
    if (stream == null || stream->get_schema == null || stream->get_next == null || stream->release == null) {
        throw new InvalidOperationException("The exported ArrowArrayStream did not expose the required ABI callbacks.");
    }

    var schema = new NativeArrowSchema();
    int status = stream->get_schema(stream, &schema);
    try {
        ThrowForArrowStatus(stream, status, "get_schema");
        if (schema.release == null || schema.n_children != expectedColumns) {
            throw new InvalidOperationException("The native Arrow schema did not preserve the expected columns or release contract.");
        }
    } finally {
        if (schema.release != null) schema.release(&schema);
    }

    long rows = 0L;
    int batches = 0;
    while (true) {
        var array = new NativeArrowArray();
        status = stream->get_next(stream, &array);
        try {
            ThrowForArrowStatus(stream, status, "get_next");
            if (array.release == null) break;
            if (array.length < 0 || array.length > maximumBatchRows || array.n_children != expectedColumns) {
                throw new InvalidOperationException("The native Arrow callback violated the configured batch or schema bound.");
            }
            rows += array.length;
            batches++;
        } finally {
            if (array.release != null) array.release(&array);
        }
    }

    if (rows != expectedRows || batches != 2) {
        throw new InvalidOperationException("The direct Arrow C ABI consumer did not observe the complete bounded stream.");
    }

    stream->release(stream);
    if (stream->release != null) {
        throw new InvalidOperationException("The native Arrow stream release callback did not clear itself.");
    }
}

static unsafe void ThrowForArrowStatus(NativeArrowArrayStream* stream, int status, string operation) {
    if (status == 0) return;
    string? detail = stream->get_last_error == null
        ? null
        : Marshal.PtrToStringUTF8((nint)stream->get_last_error(stream));
    throw new InvalidOperationException($"Arrow C ABI {operation} failed with status {status}: {detail ?? "no native detail"}.");
}

[StructLayout(LayoutKind.Sequential)]
internal unsafe struct NativeArrowArrayStream {
    public delegate* unmanaged<NativeArrowArrayStream*, NativeArrowSchema*, int> get_schema;
    public delegate* unmanaged<NativeArrowArrayStream*, NativeArrowArray*, int> get_next;
    public delegate* unmanaged<NativeArrowArrayStream*, byte*> get_last_error;
    public delegate* unmanaged<NativeArrowArrayStream*, void> release;
    public void* private_data;
}

[StructLayout(LayoutKind.Sequential)]
internal unsafe struct NativeArrowSchema {
    public byte* format;
    public byte* name;
    public byte* metadata;
    public long flags;
    public long n_children;
    public NativeArrowSchema** children;
    public NativeArrowSchema* dictionary;
    public delegate* unmanaged<NativeArrowSchema*, void> release;
    public void* private_data;
}

[StructLayout(LayoutKind.Sequential)]
internal unsafe struct NativeArrowArray {
    public long length;
    public long null_count;
    public long offset;
    public long n_buffers;
    public long n_children;
    public byte** buffers;
    public NativeArrowArray** children;
    public NativeArrowArray* dictionary;
    public delegate* unmanaged<NativeArrowArray*, void> release;
    public void* private_data;
}

internal sealed class CsvSmokeRow {
    public string Name { get; set; } = string.Empty;
    public int Score { get; set; }
    public DateOnly Date { get; set; }
    public TimeOnly Time { get; set; }
}
