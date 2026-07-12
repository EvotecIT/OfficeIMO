using OfficeIMO.Excel.Html;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersStructuredResults {
    [Fact]
    public void SharedConversionDocumentFeedsExcelAndPowerPointAdapters() {
        HtmlConversionDocument excelSource = HtmlConversionDocumentBuilder.Build("""
            <main><section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1:A1">
              <table><tr><td>42</td></tr></table>
            </section></main>
            """);
        HtmlConversionDocument powerPointSource = HtmlConversionDocumentBuilder.Build("""
            <main><section class="officeimo-slide"><p>Prepared slide</p></section></main>
            """);

        HtmlToExcelResult excelResult = excelSource.ToExcelDocumentResult();
        using OfficeIMO.Excel.ExcelDocument workbook = excelResult.Workbook;
        HtmlToPowerPointResult powerPointResult = powerPointSource.ToPowerPointPresentationResult();
        using OfficeIMO.PowerPoint.PowerPointPresentation presentation = powerPointResult.Presentation;

        Assert.True(excelResult.Succeeded);
        Assert.True(powerPointResult.Succeeded);
        Assert.True(Assert.Single(workbook.Sheets).TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? value));
        Assert.Equal("42", value!.Text);
        Assert.Contains(Assert.Single(presentation.Slides).TextBoxes, textBox => textBox.Text == "Prepared slide");
    }

    [Fact]
    public void ExcelHtml_ConvenienceImportThrowsWhenSemanticEnvelopeIsMissing() {
        HtmlConversionException exception = Assert.Throws<HtmlConversionException>(() =>
            "<main><p>Not an Excel envelope</p></main>".ToExcelDocument());

        HtmlDiagnostic diagnostic = Assert.Single(exception.Diagnostics);
        Assert.Equal(HtmlConversionDiagnosticCodes.SemanticContentMissing, diagnostic.Code);
        Assert.Equal(HtmlDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal(HtmlConversionLossKind.Failure, diagnostic.LossKind);
    }

    [Fact]
    public void PowerPointHtml_ConvenienceImportThrowsWhenSemanticEnvelopeIsMissing() {
        HtmlConversionException exception = Assert.Throws<HtmlConversionException>(() =>
            "<main><p>Not a PowerPoint envelope</p></main>".ToPowerPointPresentation());

        HtmlDiagnostic diagnostic = Assert.Single(exception.Diagnostics);
        Assert.Equal(HtmlConversionDiagnosticCodes.SemanticContentMissing, diagnostic.Code);
        Assert.Equal(HtmlDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal(HtmlConversionLossKind.Failure, diagnostic.LossKind);
    }
}
