using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNeutralProjectionImageTests {
    [Theory]
    [InlineData(1600, 400, 612, 792, 72)]
    [InlineData(400, 1600, 612, 792, 72)]
    [InlineData(1600, 400, 300, 400, 30)]
    [InlineData(1600, 1, 612, 792, 72)]
    [InlineData(1, 80, 612, 792, 72)]
    public void SupportedAssetsFitConfiguredPageWithoutChangingPixelsOrAspectRatio(
        int width, int height, double pageWidth, double pageHeight, double margin) {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(width, height, OfficeColor.Red));
        var source = new OfficeDocumentModel {
            Blocks = new[] { new OfficeDocumentModelBlock { Kind = "paragraph", Text = "Searchable source content" } },
            Assets = new[] { new OfficeDocumentModelAsset { Id = "diagram", Kind = "preview-png", PayloadBytes = png } }
        };
        var pdfOptions = new PdfOptions {
            PageWidth = pageWidth, PageHeight = pageHeight,
            MarginLeft = margin, MarginRight = margin, MarginTop = margin, MarginBottom = margin
        };

        byte[] bytes = source.ToPdfDocumentResult(new PdfProjectionOptions {
            PdfOptions = pdfOptions, IncludeMetadata = false
        }).ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(bytes);
        PdfImagePlacement placement = Assert.Single(read.Pages.SelectMany(page => page.GetImagePlacements()));
        PdfExtractedImage embedded = Assert.Single(read.Pages.SelectMany(page => page.GetImages()));

        Assert.InRange(placement.X, margin - 0.001, pageWidth - margin);
        Assert.InRange(placement.Y, margin - 0.001, pageHeight - margin);
        Assert.True(placement.X + placement.Width <= pageWidth - margin + 0.001);
        Assert.True(placement.Y + placement.Height <= pageHeight - margin + 0.001);
        Assert.InRange(Math.Abs(placement.Width / placement.Height - (double)width / height), 0, 0.001);
        Assert.Equal(width, embedded.Width);
        Assert.Equal(height, embedded.Height);
        Assert.True(OfficeRasterImageDecoder.TryDecode(embedded.Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(width / 2, height / 2));
        Assert.Contains("Searchable source content", read.ExtractText());
        Assert.Equal(pageWidth, pdfOptions.PageWidth);
        Assert.Equal(margin, pdfOptions.MarginLeft);
    }

    [Theory]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Information, "Omission", OfficeConversionLossKind.Omission)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Information, "Unassessed", OfficeConversionLossKind.Unassessed)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Warning, "Unassessed", OfficeConversionLossKind.Unassessed)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Warning, "None", OfficeConversionLossKind.None)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Warning, "Omission", OfficeConversionLossKind.Omission)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Error, "None", OfficeConversionLossKind.Failure)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Warning, "invalid", OfficeConversionLossKind.Approximation)]
    [InlineData(OfficeDocumentModelDiagnosticSeverity.Warning, "999", OfficeConversionLossKind.Approximation)]
    public void SourceLossCategorySurvivesProjectionSeparatelyFromSeverity(
        OfficeDocumentModelDiagnosticSeverity severity, string sourceLoss, OfficeConversionLossKind expectedLoss) {
        var source = new OfficeDocumentModel {
            Diagnostics = new[] { new OfficeDocumentModelDiagnostic {
                Severity = severity, Code = "source-rendering", Message = "Source rendering evidence",
                Attributes = new Dictionary<string, string> { ["lossKind"] = sourceLoss, ["sourcePage"] = "3" }
            } }
        };

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        PdfConversionWarning warning = Assert.Single(result.Warnings);

        Assert.Equal(expectedLoss, warning.LossKind);
        Assert.Equal(sourceLoss, warning.Details["lossKind"]);
        Assert.Equal("3", warning.Details["sourcePage"]);
        Assert.Equal((int)severity, (int)warning.Severity);
        Assert.Equal(expectedLoss != OfficeConversionLossKind.None, result.HasLoss);
        Assert.Equal(expectedLoss, Assert.Single(result.Report.FidelityDiagnostics).LossKind);
        if (expectedLoss == OfficeConversionLossKind.None) {
            result.RequireNoLoss();
            Assert.Equal(PdfConversionFidelityStatus.Faithful, result.Report.FidelityStatus);
        } else {
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
            Assert.Equal(PdfConversionFidelityStatus.Degraded, result.Report.FidelityStatus);
        }
    }

    [Theory]
    [InlineData(PdfConversionWarningSeverity.Information)]
    [InlineData(PdfConversionWarningSeverity.Warning)]
    public void UnassessedWarningCannotEstablishFaithfulConversion(PdfConversionWarningSeverity severity) {
        var warning = new PdfConversionWarning("source", "rendering", "image", "Fidelity was not assessed.",
            severity, OfficeConversionLossKind.Unassessed);
        var report = new PdfConversionReport();
        report.Add(warning);

        Assert.Equal(OfficeConversionLossKind.Unassessed, warning.ToFidelityDiagnostic().LossKind);
        Assert.True(report.HasLoss);
        Assert.Equal(PdfConversionFidelityStatus.Degraded, report.FidelityStatus);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(999)]
    public void WarningRejectsUndefinedFidelityCategory(int value) {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfConversionWarning(
            "source", "rendering", "image", "Invalid category.", PdfConversionWarningSeverity.Information,
            (OfficeConversionLossKind)value));
    }
}
