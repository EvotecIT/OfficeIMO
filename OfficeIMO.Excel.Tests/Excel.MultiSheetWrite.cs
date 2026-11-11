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
