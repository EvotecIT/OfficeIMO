using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void ExcelHtml_GenericImportRetainsWrappedNestedListText() {
        HtmlToExcelResult result = HtmlConversionDocument
            .Parse("<ul><li>Parent<div><ol><li>Child</li></ol></div></li></ul>")
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        Assert.True(sheet.TryGetCellValueSnapshot(2, 1, out ExcelCellValueSnapshot? value));
        Assert.Equal("• Parent\n  1. Child", value!.Text);
    }

    [Fact]
    public void PowerPointHtml_RendersHeadingOnlySemanticSections() {
        HtmlToPowerPointResult result = HtmlConversionDocument
            .Parse("<p>First slide</p><h1>Heading-only slide</h1>")
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Equal(2, result.Slides);
        PowerPointSlide second = presentation.Slides[1];
        Assert.Contains(second.TextBoxes, textBox =>
            textBox.Text.Contains("Heading-only slide", StringComparison.Ordinal));
    }
}
