using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>
    /// Workbook text with an unpaired surrogate cannot be stored by any Excel format. It is rejected where it is
    /// supplied, rather than accepted and later failing inside the XML writer at save time.
    /// </summary>
    [Theory]
    [InlineData("sheet")]
    [InlineData("rename")]
    [InlineData("numberFormat")]
    [InlineData("rangeNumberFormat")]
    [InlineData("fontName")]
    [InlineData("rangeFontName")]
    public void UnpairedSurrogates_AreRejectedWhenSupplied(string target) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue(1.5);

        ArgumentException exception = Assert.Throws<ArgumentException>(() => {
            switch (target) {
                case "sheet": document.AddSheet("Bad\uD800Name", ExcelSheetNameValidationMode.Strict); break;
                case "rename": sheet.Name = "Bad\uDC00Name"; break;
                case "numberFormat": sheet.Cell(1, 1).SetNumberFormat("0.00\"\uDC00\""); break;
                case "rangeNumberFormat": sheet.Range("A1:A2").SetNumberFormat("0\uD800"); break;
                case "fontName": sheet.Cell(1, 1).SetFontFamily("Ar\uD800ial"); break;
                default: sheet.Range("A1:A2").SetFontFamily("Ar\uDC00ial"); break;
            }
        });

        Assert.Contains("unpaired UTF-16 surrogate", exception.Message, StringComparison.Ordinal);
        Assert.Equal(new[] { "Data" }, document.Sheets.Select(item => item.Name).ToArray());
    }

    [Fact]
    public void SanitizedSheetNames_ReplaceUnpairedSurrogatesAndKeepPairs() {
        using var document = ExcelDocument.Create();

        ExcelSheet sheet = document.AddSheet("A\uD800B 😀", ExcelSheetNameValidationMode.Sanitize);

        Assert.Equal("A_B 😀", sheet.Name);
        using var saved = new MemoryStream();
        document.Save(saved);
    }
}
