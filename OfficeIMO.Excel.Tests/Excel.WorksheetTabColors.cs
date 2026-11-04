using System.IO;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace OfficeIMO.Tests;

public sealed class ExcelWorksheetTabColors {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StatusTabColorsUseTheSelectedCreationPalette(bool fromTemplate) {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(success: OfficeColor.Parse("#123456")));
        using var templateBytes = new MemoryStream();
        using (var template = ExcelDocument.Create()) {
            template.AddSheet("Data");
            template.Save(templateBytes);
        }
        templateBytes.Position = 0;
        using var book = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme, TemplateStream = fromTemplate ? templateBytes : null });
        var sheet = fromTemplate ? book["Data"] : book.AddSheet("Data");
        sheet.SetTabColor(OfficeColor.Success);
        using var bytes = new MemoryStream();
        book.Save(bytes);
        bytes.Position = 0;
        using var package = SpreadsheetDocument.Open(bytes, false);
        var color = package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProperties>()!.TabColor!;
        Assert.Null(color.Theme);
        Assert.Equal("FF123456", color.Rgb!.Value);
    }

    [Theory]
    [InlineData("Accent1", 4)]
    [InlineData("Accent6", 9)]
    [InlineData("Text", 1)]
    [InlineData("TextMuted", 3)]
    [InlineData("Background", 0)]
    [InlineData("Success", null)]
    [InlineData("#33669980", null)]
    public void SharedColorsPersistAsNativeThemeSlotsOrArgbLiterals(string value, int? themeIndex) {
        using var book = ExcelDocument.Create();
        var sheet = book.AddSheet("Data");
        sheet.SetTabColor(OfficeColor.Parse(value));
        using var bytes = new MemoryStream();
        book.Save(bytes);
        bytes.Position = 0;
        using var package = SpreadsheetDocument.Open(bytes, false);
        var color = package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProperties>()!.TabColor!;
        if (themeIndex.HasValue) {
            Assert.Equal((uint)themeIndex.Value, color.Theme!.Value);
            Assert.Null(color.Rgb);
        } else {
            Assert.Null(color.Theme);
            Assert.Equal(value == "Success" ? OfficeTheme.Default.Palette.Success.ToArgbHex() : "80336699", color.Rgb!.Value);
        }
    }
}
