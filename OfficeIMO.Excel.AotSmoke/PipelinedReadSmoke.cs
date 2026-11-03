using OfficeIMO.Data;
using OfficeIMO.Excel;

internal static class PipelinedReadSmoke {
    private const int Count = 50_000;

    internal static void Run() {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output,
            new ExcelWriterOptions { IncludeCellReferences = false, UseSharedStrings = false, CreateTable = false }, leaveOpen: true)) {
            writer.WriteTable("Large", Rows(), ["Name", "Id", "CreatedOn", "Amount"],
                static (rowWriter, row) => rowWriter.Write(row.Name).Write(row.Id).Write(row.CreatedOn).Write(row.Amount));
        }
        byte[] bytes = output.ToArray();
        using (var input = new MemoryStream(bytes))
        using (var reader = ExcelDocument.OpenReader(input)) {
            int count = 0;
            while (reader.Read()) {
                Row expected = Create(count++);
                if (reader.GetString(0) != expected.Name || reader.GetInt32(1) != expected.Id
                    || reader.GetDateTime(2) != expected.CreatedOn || reader.GetDouble(3) != expected.Amount)
                    throw new InvalidOperationException("The large compressed reader lost an indexed value.");
            }
            if (count != Count) throw new InvalidOperationException("The large compressed reader lost rows.");
        }
        using (var input = new MemoryStream(bytes))
        using (var reader = ExcelDocument.OpenReader(input)) {
            int count = 0;
            foreach (Row row in reader.RowsAs<Row>()) {
                if (row != Create(count++)) throw new InvalidOperationException("The large compressed reader lost a typed row.");
            }
            if (count != Count) throw new InvalidOperationException("The large compressed typed reader lost rows.");
        }
        Console.WriteLine("PASS | Large compressed implicit worksheets retain every scalar and typed row under NativeAOT");
    }

    private static Row Create(int index) => new("Name " + index, index + 1,
        new DateTime(2024, 1, 1).AddDays(index % 365), index * 1.25);

    private static IEnumerable<Row> Rows() {
        for (int index = 0; index < Count; index++) yield return Create(index);
    }

    private sealed record Row(string Name, int Id, DateTime CreatedOn, double Amount);
}
