using OfficeIMO.CSV;

CsvDocument document = CsvDocument.Parse("Name,Score\nAlice,42\n");
if (!document.Header.SequenceEqual(new[] { "Name", "Score" })) {
    throw new InvalidOperationException("The CSV parser did not preserve the header schema.");
}

CsvSmokeRow mapped = document.RowsAs<CsvSmokeRow>().Single();
if (mapped.Name != "Alice" || mapped.Score != 42) {
    throw new InvalidOperationException("The CSV typed-row mapper lost a value.");
}

using var reader = CsvDocument.OpenTextDataReader("Name,Score\nAlice,42\n");
CsvSmokeRow streamed = reader.RowsAs<CsvSmokeRow>().Single();
if (streamed.Name != "Alice" || streamed.Score != 42) {
    throw new InvalidOperationException("The forward-only CSV typed-row mapper lost a value.");
}

Console.WriteLine("PASS | CSV parse, schema inspection, and typed-row mapping");

internal sealed class CsvSmokeRow {
    public string Name { get; set; } = string.Empty;
    public int Score { get; set; }
}
