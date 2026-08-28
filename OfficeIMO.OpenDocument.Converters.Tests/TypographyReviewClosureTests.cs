using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class TypographyReviewClosureTests {
    [Fact]
    public void OdsTextCaseMaterializationUsesDocumentLanguage() {
        OdsDocument source = OdsDocument.Create();
        source.Metadata.Language = "tr-TR";
        OdsCell cell = source.AddSheet("Text").Cell(0, 0);
        cell.SetString("i");
        cell.TextTransform = OdfTextTransform.Uppercase;

        using ExcelDocument target = source.ToExcelDocumentResult().Value;

        Assert.Equal("İ", target.Sheets.Single().CellAt(1, 1).GetValue().Value);
    }

    [Fact]
    public void OdsTextCaseMaterializationFallsBackToInvariantCulture() {
        OdsDocument source = OdsDocument.Create();
        source.Metadata.Language = "not-a-culture";
        OdsCell cell = source.AddSheet("Text").Cell(0, 0);
        cell.SetString("i");
        cell.TextTransform = OdfTextTransform.Uppercase;

        using ExcelDocument target = source.ToExcelDocumentResult().Value;

        Assert.Equal("I", target.Sheets.Single().CellAt(1, 1).GetValue().Value);
    }
}
