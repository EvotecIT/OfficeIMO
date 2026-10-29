using OfficeIMO.Data;
using OfficeIMO.Excel;

internal static class RecordReadSmoke {
    internal static void Run() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Records");
        var expected = new[] { new Row(7, "North", 12.5m), new Row(8, "South", 25m) };
        ExcelTable table = sheet.AddTable(expected);
        if (!sheet.RowsAs<Row>().SequenceEqual(expected)
            || !sheet.RowsAsParallel<Row>().SequenceEqual(expected)
            || !table.RowsAs<Row>().SequenceEqual(expected)) {
            throw new InvalidOperationException("Constructor-bound records lost values in typed worksheet or table reads.");
        }
        table.Resize("A1:C4"); // Reserve the final row for native table totals.
        table.SetTotals(new Dictionary<string, ExcelTableTotalsFunction> {
            [nameof(Row.Total)] = ExcelTableTotalsFunction.Sum
        });
        if (!table.RowsAs<Row>().SequenceEqual(expected)) {
            throw new InvalidOperationException("Typed table reads included the totals row.");
        }
        ExcelSheet getter = document.AddSheet("Getter");
        getter.CellValue(1, 1, "Name");
        getter.CellValue(2, 1, " north ");
        if (getter.RowsAs<GetterRow>().Single().Name != "NORTH") {
            throw new InvalidOperationException("Getter-only records skipped constructor normalization.");
        }
        ExcelSheet values = document.AddSheet("Values");
        values.CellValue(1, 1, "Value");
        values.CellValue(2, 1, 7);
        if (values.RowsAs<ValueRow>().Single().Value != 70) {
            throw new InvalidOperationException("Record structs skipped their primary constructor.");
        }
        bool missingConstructorColumnRejected = false;
        try {
            _ = values.RowsAs<MeasuredValue>().ToArray();
        } catch (DataMappingException) {
            missingConstructorColumnRejected = true;
        }
        if (!missingConstructorColumnRejected) {
            throw new InvalidOperationException("A record struct bypassed a required constructor column.");
        }
        PlainValue plain = values.RowsAs<PlainValue>().Single();
        if (plain.Value != 7 || plain.Unit is not null) {
            throw new InvalidOperationException("Plain mutable structs lost their default construction path.");
        }
        foreach (bool readFirst in new[] { false, true }) {
            using var hiddenDocument = ExcelDocument.Create();
            ExcelTable hidden = hiddenDocument.AddSheet("Hidden").AddTable(expected,
                new OfficeIMO.OfficeTableOptions { IncludeHeaders = false });
            if (readFirst && !hidden.RowsAs<Row>().SequenceEqual(expected)) {
                throw new InvalidOperationException("Materializing a headerless table lost its native column schema.");
            }
            using var saved = new MemoryStream();
            hiddenDocument.Save(saved);
            saved.Position = 0;
            using var loaded = ExcelDocument.Load(saved);
            if (!loaded.GetSheet("Hidden").Table(hidden.NameOrRange).RowsAs<Row>().SequenceEqual(expected)) {
                throw new InvalidOperationException("Saving a headerless table lost its native column schema.");
            }
        }
        Console.WriteLine("PASS | Constructor-bound worksheet and table records retain values and native schemas");
    }

    private sealed record Row(int Id, string Customer, decimal Total);
    private sealed record GetterRow(string Name) {
        public string Name { get; } = Name.Trim().ToUpperInvariant();
    }
    private struct PlainValue {
        public int Value { get; set; }
        public string? Unit { get; set; }
    }
    private readonly record struct MeasuredValue(int Value, string Unit);
    private readonly record struct ValueRow(int Value) {
        public int Value { get; init; } = Value * 10;
    }
}
