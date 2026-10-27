using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInspectionProjectionConversionScopeTests {
    [Theory]
    [InlineData("word", false)]
    [InlineData("word", true)]
    [InlineData("excel", false)]
    [InlineData("excel", true)]
    [InlineData("powerpoint-tables", false)]
    [InlineData("powerpoint-tables", true)]
    [InlineData("powerpoint-editable", false)]
    [InlineData("powerpoint-editable", true)]
    [InlineData("html-semantic", false)]
    [InlineData("html-semantic", true)]
    [InlineData("html-positioned", false)]
    [InlineData("html-positioned", true)]
    [InlineData("rtf", false)]
    [InlineData("rtf", true)]
    public void ConversionCannotReportLossFromUnrequestedSourceProjection(string converter, bool omitWidgets) {
        PdfDocumentReadResult logical = PdfDocument.Load(PdfInspectionProjectionFixture.Build())
            .Read(new PdfLoadOptions { IncludeImages = omitWidgets, IncludeFormWidgets = !omitWidgets });
        Assert.Contains("Native page text", logical.Text);
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => Convert(logical, converter));
        Assert.Contains("unrequested projection", error.Message);
    }

    [Theory]
    [InlineData(PowerPointPdfImportMode.EditableContent)]
    [InlineData(PowerPointPdfImportMode.HybridVisualAndEditableTables)]
    public void OpenedPdfCannotBuildConversionFromUninspectedLogicalContent(PowerPointPdfImportMode mode) {
        PdfDocument pdf = PdfDocument.Load(PdfInspectionProjectionFixture.Build());
        var options = new PowerPointPdfImportOptions { Mode = mode,
            ReadOptions = new PdfLoadOptions { IncludeImages = false } };
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => pdf.ConvertToPowerPointWithEvidence(options));
        Assert.Contains("unrequested projection", error.Message);
    }

    [Fact]
    public void FullSourceConversionsRetainKnownContentAndOmissionDiagnostics() {
        PdfDocumentReadResult logical = PdfDocument.Load(PdfInspectionProjectionFixture.Build()).Read();
        PdfExcelTableImportResult excel = logical.ConvertToExcelWithEvidence();
        using (excel.Value) {
            Assert.Equal(1, excel.Report.SourceScope.ImageCount);
            Assert.Equal(1, excel.Report.SourceScope.FormWidgetCount);
            Assert.True(excel.HasOmittedPageContent);
        }
        PdfPowerPointConversionResult powerpoint = logical.ConvertToPowerPointWithEvidence();
        using (powerpoint.Value) {
            Assert.Equal(1, powerpoint.Report.SourceScope!.ImageCount);
            Assert.True(powerpoint.HasOmittedPageContent);
        }
        PdfWordConversionResult word = logical.ConvertToWordWithEvidence();
        using (word.Value) Assert.NotEmpty(word.Report.FidelityDiagnostics);
        Assert.Contains("Native page text", logical.ToHtmlResult().Value);
        Assert.Contains("Native page text", logical.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile()).Value);
        PdfRtfConversionResult rtf = logical.ToRtfDocumentResult();
        Assert.Contains(rtf.Report.Warnings, warning => warning.Code == "IMAGES_NOT_IMPORTED");
    }

    private static void Convert(PdfDocumentReadResult logical, string converter) {
        switch (converter) {
            case "word": using (logical.ConvertToWordWithEvidence().Value) { } break;
            case "excel": using (logical.ConvertToExcelWithEvidence().Value) { } break;
            case "powerpoint-tables": using (logical.ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
                Mode = PowerPointPdfImportMode.EditableTables, ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(1, 1) }
            }).Value) { } break;
            case "powerpoint-editable": using (logical.ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
                Mode = PowerPointPdfImportMode.EditableContent, ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(1, 1) }
            }).Value) { } break;
            case "html-semantic": logical.ToHtmlResult(); break;
            case "html-positioned": logical.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile()); break;
            case "rtf": logical.ToRtfDocumentResult(); break;
            default: throw new ArgumentOutOfRangeException(nameof(converter));
        }
    }
}
