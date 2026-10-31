using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

internal static class ExcelPackageWriterFixtures {
    /// <summary>Creates the valid theme-free package profile required by tests of the Simple serializer.</summary>
    internal static void WithoutOptionalTheme(ExcelDocument document) {
        WorkbookPart workbook = document.OpenXmlDocument.WorkbookPart!;
        if (workbook.ThemePart is { } theme) workbook.DeletePart(theme);
    }
}

public sealed class ExcelModernPackageWriterTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NewWorkbookRetainsModernThemeThroughTheExtendedFastWriter(bool asynchronous) {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create();
        document.AddSheet("Modern").CellValue(1, 1, "Modern workbook");
        string theme = document.OpenXmlDocument.WorkbookPart!.ThemePart!.Theme!.OuterXml;
        var options = new ExcelSaveOptions { ValidateOpenXml = true };
        if (asynchronous) await document.SaveAsync(stream, options: options);
        else document.Save(stream, options);

        Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
        Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
        byte[] bytes = stream.ToArray();
        using (var source = new MemoryStream(bytes, writable: false))
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(source, false)) {
            Assert.Equal(theme, package.WorkbookPart!.ThemePart!.Theme!.OuterXml);
        }
        using var input = new MemoryStream(bytes, writable: false);
        using var reopened = ExcelDocument.Load(input);
        Assert.Equal("Modern workbook", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
    }
}
