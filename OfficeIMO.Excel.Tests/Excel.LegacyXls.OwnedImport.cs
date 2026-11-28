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
        Assert.Empty(document.OpenXmlSdkFallbacks);

        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsx })));
        Assert.True(reopened.Sheets[0].TryGetCellText(2, 2, out string? saved));
        Assert.Equal(number, saved);
        Assert.Empty(document.OpenXmlSdkFallbacks);
    }
}
