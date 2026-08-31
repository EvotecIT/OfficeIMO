using System.Data;
using System.Globalization;
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
IArrowArrayStream importedArrowStream;
unsafe {
    importedArrowStream = CArrowArrayStreamImporter.ImportArrayStream(
        arrowOwner.DangerousGetPointer());
}
arrowOwner.Dispose();
using (importedArrowStream) {
    using RecordBatch importedArrowBatch =
        (await importedArrowStream.ReadNextRecordBatchAsync())!;
    if (importedArrowBatch.Length != 1
        || importedArrowBatch.Schema.FieldsList.Count != 2
        || importedArrowBatch.Column(0).Length != 1) {
        throw new InvalidOperationException("The bounded Arrow C stream did not survive NativeAOT.");
    }
}

Console.WriteLine("PASS | CSV parse, schema inspection, mapping, Arrow C stream, and sequential/parallel typed DataReader writing");

internal sealed class CsvSmokeRow {
    public string Name { get; set; } = string.Empty;
    public int Score { get; set; }
    public DateOnly Date { get; set; }
    public TimeOnly Time { get; set; }
}
