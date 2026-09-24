using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlGenericFormTextFallbackTests {
    [Fact]
    public void GenericEditableTargetsKeepFormTextAndReportInteractionLoss() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<main><form><h2>Was this page helpful?</h2><label><input type='radio' name='answer'>Yes</label></form></main>");

        HtmlToExcelResult excel = source.ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using (ExcelDocument workbook = excel.RequireValue()) {
            using var artifact = workbook.ToStream();
            using ExcelDocument reopened = ExcelDocument.Load(artifact);
            Assert.Contains("Was this page helpful?", reopened.ToHtml());
        }
        Assert.Contains(excel.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);

        HtmlToPowerPointResult slides = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using (PowerPointPresentation presentation = slides.RequireValue()) {
            using var artifact = presentation.ToStream();
            using PowerPointPresentation reopened = PowerPointPresentation.Load(artifact);
            Assert.Contains("Was this page helpful?", reopened.ToHtml());
        }
        Assert.Contains(slides.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }
}
