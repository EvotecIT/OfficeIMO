using OfficeIMO.CSV;

CsvDocument document = CsvDocument.Parse("Name,Score\nAlice,42\n");
if (!document.Header.SequenceEqual(new[] { "Name", "Score" })) {
    throw new InvalidOperationException("The CSV parser did not preserve the header schema.");
}

Console.WriteLine("PASS | CSV parse and schema inspection");
