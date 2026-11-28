using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void LegacyXls_PlainWorksheetLoadsAndSavesAsXlsxThroughOwnedParts() {
        using var stream = new MemoryStream(CreateMinimalLegacyXlsCompound());
        using var document = ExcelDocument.Load(stream);

        Assert.Equal(ExcelFileFormat.Xls, document.SourceFormat);
        Assert.Empty(document.OpenXmlSdkFallbacks);
        ExcelSheet sheet = Assert.Single(document.Sheets);
        Assert.True(sheet.TryGetCellText(1, 1, out string? text));
        Assert.Equal("Name", text);
        Assert.True(sheet.TryGetCellText(2, 2, out string? number));
        Assert.Equal("42", number);
        Assert.False(sheet.TryGetCellText(1, (1 << 15) + 1, out _));
        Assert.Empty(document.OpenXmlSdkFallbacks);

        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsx })));
        Assert.True(reopened.Sheets[0].TryGetCellText(2, 2, out string? saved));
        Assert.Equal(number, saved);
        Assert.Empty(document.OpenXmlSdkFallbacks);
    }

    [Theory]
    [InlineData("Before\u0001After", false)]
    [InlineData("First\nSecond", true)]
    public void LegacyXls_OwnedWorksheetKeepsTextAndWrapping(string value, bool wrap) {
        byte[] workbook = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream(value);
        using var stream = new MemoryStream(LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbook));
        using var document = ExcelDocument.Load(stream);

        Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? text));
        Assert.Equal(value, text);
        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsx })));
        Assert.Empty(document.OpenXmlSdkFallbacks);
        Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string? saved));
        Assert.Equal(value, saved);
        Assert.Equal(wrap, reopened.Sheets[0].Cell(1, 1).GetStyle().WrapText);
    }
}
