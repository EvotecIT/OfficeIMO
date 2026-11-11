using System.Globalization;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private sealed record MultiSheetRow(string Name, int Id, DateTime CreatedOn, double? Amount);

    [Fact]
    public void WriteData_OnSeveralSheets_SavesEverySheetsValues() {
        var rows = Enumerable.Range(0, 300)
            .Select(i => new MultiSheetRow("Name " + i.ToString(CultureInfo.InvariantCulture), i + 1,
                new DateTime(2024, 1, 1).AddDays(i), i % 7 == 0 ? null : i * 1.25))
            .ToArray();
        using var output = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create()) {
            for (int sheet = 1; sheet <= 3; sheet++) {
                document.AddSheet("Sheet" + sheet.ToString(CultureInfo.InvariantCulture)).WriteData(rows,
                    ("Name", r => r.Name), ("Id", r => r.Id), ("CreatedOn", r => r.CreatedOn), ("Amount", r => r.Amount));
            }
            document.Save(output);
        }

        for (int sheet = 1; sheet <= 3; sheet++) {
            output.Position = 0;
            using var reader = ExcelDocument.OpenReader(output, new ExcelReaderOptions { SheetName = "Sheet" + sheet.ToString(CultureInfo.InvariantCulture) });
            int index = 0;
            while (reader.Read()) {
                MultiSheetRow expected = rows[index++];
                Assert.Equal(expected.Name, reader.GetString(0));
                Assert.Equal(expected.Id, Convert.ToInt32(reader.GetValue(1), CultureInfo.InvariantCulture));
                Assert.Equal(expected.CreatedOn, reader.GetDateTime(2));
                if (expected.Amount is double amount) Assert.Equal(amount, reader.GetDouble(3));
                else Assert.True(reader.IsDBNull(3));
            }
            Assert.Equal(rows.Length, index);
        }
    }

    [Fact]
    public void Save_OfSeveralWrittenSheets_IsTheSameWithAndWithoutWorkerThreads() {
        static byte[] Save(bool sequential) {
            using var output = new MemoryStream();
            using (ExcelDocument document = ExcelDocument.Create()) {
                if (sequential) document.Execution.Mode = ExcelExecutionMode.Sequential;
                for (int sheet = 1; sheet <= 6; sheet++) {
                    var rows = Enumerable.Range(0, 400)
                        .Select(i => new MultiSheetRow("Row " + (i * sheet).ToString(CultureInfo.InvariantCulture) + (i % 5 == 0 ? " & <b>\nnext" : ""),
                            i, new DateTime(2024, 1, 1).AddDays(i), i % 3 == 0 ? null : i / 8d))
                        .ToArray();
                    document.AddSheet("Sheet" + sheet.ToString(CultureInfo.InvariantCulture)).WriteData(rows,
                        ("Name", r => r.Name), ("Id", r => r.Id), ("CreatedOn", r => r.CreatedOn), ("Amount", r => r.Amount));
                }
                document.Save(output);
            }
            return output.ToArray();
        }

        static List<(string Name, byte[] Content)> Parts(byte[] package) {
            using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            return archive.Entries.Select(entry => {
                using Stream stream = entry.Open();
                using var content = new MemoryStream();
                stream.CopyTo(content);
                return (entry.FullName, content.ToArray());
            }).ToList();
        }

        var parallel = Parts(Save(sequential: false));
        var sequential = Parts(Save(sequential: true));

        Assert.Equal(sequential.Select(part => part.Name), parallel.Select(part => part.Name));
        foreach (var (expected, actual) in sequential.Zip(parallel, (expected, actual) => (expected, actual))) {
            // Relationship ids and creation times are generated per save; every other part must match exactly.
            if (expected.Name.EndsWith(".rels", StringComparison.Ordinal) || expected.Name is "xl/workbook.xml" or "docProps/core.xml") continue;
            Assert.True(expected.Content.SequenceEqual(actual.Content), expected.Name + " differs.");
        }
    }

    [Fact]
    public void RowWriter_FailureInALaterLargeSheet_SurfacesTheCallersError() {
        var rows = Enumerable.Range(0, 20_000).ToArray();
        using var output = new MemoryStream();
        var writer = ExcelDocument.CreateWriter(output, leaveOpen: true);
        writer.WriteTable("First", rows, ["Value", "Text"], static (row, value) => row.Write(value).Write("Row text " + value));

        var error = Assert.Throws<InvalidOperationException>(() => writer.WriteTable("Second", rows, ["Value", "Text"], static (row, value) => {
            if (value == 15_000) throw new InvalidOperationException("Row 15000 failed.");
            row.Write(value).Write("Row text " + value);
        }));

        Assert.Equal("Row 15000 failed.", error.Message);
        writer.Dispose();
    }

    [Theory]
    [InlineData("Q1::Sales//North", "Q1_Sales_North")]
    [InlineData("__Totals__", "Totals")]
    [InlineData("a_ _b", "a_ _b")]
    public void AddSheet_SanitizeMode_ReplacesInvalidCharactersAndCollapsesUnderscores(string requested, string expected) {
        using ExcelDocument document = ExcelDocument.Create();

        ExcelSheet sheet = document.AddSheet(requested, ExcelSheetNameValidationMode.Sanitize);

        Assert.Equal(expected, sheet.Name);
    }
}
