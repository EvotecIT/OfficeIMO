using System.Data;
using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Shares identical cell values and independent validation across export profiles.</summary>
internal static class ExcelTextWriteFixture {
    internal static DataTable Create(int textLength, string textShape) {
        var table = new DataTable();
        table.Columns.Add("Id", typeof(string));
        table.Columns.Add("Notes", typeof(string));
        for (int index = 0; index < 1000; index++) {
            string marker = textShape == "Escaped" ? " & <tag> Łódź\n" : " plain Łódź ";
            string payload = textShape == "Markup"
                ? string.Concat(Enumerable.Repeat("<tag>&data</tag>", Math.Max(1, textLength / 16)))
                : new string('n', textLength / 2) + marker + new string('x', textLength / 2);
            table.Rows.Add(index.ToString(CultureInfo.InvariantCulture), index + payload);
        }
        return table;
    }

    internal static void Validate(DataTable table, string method, byte[] bytes) {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read() || reader.FieldCount != 2 || reader.GetString(0) != "Id" || reader.GetString(1) != "Notes")
            throw new InvalidOperationException($"{method} header mismatch.");
        foreach (DataRow row in table.Rows) {
            if (!reader.Read() || reader.GetString(0) != (string)row[0] || reader.GetString(1) != (string)row[1])
                throw new InvalidOperationException($"{method} row mismatch.");
        }
        if (reader.Read()) throw new InvalidOperationException($"{method} wrote extra rows.");
        Console.WriteLine($"Validated {method}: {table.Rows.Count} rows, {bytes.Length} XLSX bytes.");
    }
}
