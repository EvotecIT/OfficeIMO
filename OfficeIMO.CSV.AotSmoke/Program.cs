using OfficeIMO.CSV;
using OfficeIMO.Data;

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

Console.WriteLine("PASS | CSV parse, schema inspection, automatic mapping, and explicit AOT mapping");

internal sealed class CsvSmokeRow {
    public string Name { get; set; } = string.Empty;
    public int Score { get; set; }
    public DateOnly Date { get; set; }
    public TimeOnly Time { get; set; }
}
