using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData("xlsx", 1900)]
    [InlineData("xls", 1900)]
    [InlineData("xlsb", 1900)]
    [InlineData("xlsx", 1904)]
    [InlineData("xls", 1904)]
    [InlineData("xlsb", 1904)]
    public void Reader_ProducerCalendarAndTime_RetainsCalendarAndUnshiftedDuration(string extension, int system) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "SpreadsheetDateSerialCorpus", $"serials-{system}.{extension}");
        using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Data" });
        for (int i = 0; i < 4; i++) Assert.True(reader.Read());
        DateTime expectedCalendar = system == 1900 ? new DateTime(1900, 1, 1, 12, 0, 0) : new DateTime(1904, 1, 2, 12, 0, 0);
        DateTime expectedDuration = new DateTime(1899, 12, 31, 12, 0, 0);
        Assert.Equal(expectedCalendar, reader.GetDateTime(0));
        foreach (int ordinal in new[] { 1, 2, 3, 4 }) Assert.Equal(expectedDuration, reader.GetDateTime(ordinal));
        Assert.Equal(expectedCalendar, reader.GetDateTime(5));
        for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) Assert.Equal(1.5d, reader.GetDouble(ordinal));
    }

    [Theory]
    [InlineData("xlsx", 1900)]
    [InlineData("xls", 1900)]
    [InlineData("xlsb", 1900)]
    [InlineData("xlsx", 1904)]
    [InlineData("xls", 1904)]
    [InlineData("xlsb", 1904)]
    public void Reader_ProducerCalendarAndTime_PreservesEverySerialAndCachedFormula(string extension, int system) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "SpreadsheetDateSerialCorpus", $"serials-{system}.{extension}");
        double[] serials = { -0.5, 0, 1, 1.5, 59, 59.5, 60, 60.5, 61, 61.5, 1462, 45000, 1.5 };
        using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Data" });
        foreach (double serial in serials) {
            Assert.True(reader.Read());
            for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) {
                if (ordinal < 6) {
                    DateTime value = reader.GetDateTime(ordinal);
                    if (ordinal is >= 1 and <= 4) Assert.Equal(DateTime.FromOADate(serial), value);
                    else if (system == 1904) Assert.Equal(new DateTime(1904, 1, 1).AddDays(serial), value);
                    else if (serial < 60) Assert.Equal(new DateTime(1899, 12, 31).AddDays(serial), value);
                    else Assert.Equal(DateTime.FromOADate(serial), value);
                }
                Assert.Equal(serial, reader.GetDouble(ordinal));
            }
        }
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData("xls", 1900)]
    [InlineData("xls", 1904)]
    [InlineData("xlsb", 1900)]
    [InlineData("xlsb", 1904)]
    public void Load_ProducerCalendarAndTime_PreservesDateSystemSerialsAndStyles(string extension, int system) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "SpreadsheetDateSerialCorpus", $"serials-{system}.{extension}");
        using var document = ExcelDocument.Load(path);
        Assert.Equal(system == 1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred, document.DateSystem);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var package = SpreadsheetDocument.Open(output, false);
        var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
            .Where(cell => int.Parse(new string(cell.CellReference!.Value!.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture) > 1)
            .ToArray();
        double[] serials = { -0.5, 0, 1, 1.5, 59, 59.5, 60, 60.5, 61, 61.5, 1462, 45000, 1.5 };
        Assert.Equal(serials.Length * 7, cells.Length);
        foreach (Cell cell in cells) {
            int row = int.Parse(new string(cell.CellReference!.Value!.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
            Assert.Equal(serials[row - 2], double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture));
            if (!cell.CellReference.Value.StartsWith("G", StringComparison.Ordinal)) Assert.True(cell.StyleIndex?.Value > 0);
        }
    }
}
