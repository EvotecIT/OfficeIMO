using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPdf_SourceClippingDoesNotExemptAdditionalPdfContentFromLossChecks(bool nestedUnderSourceClip) {
        var result = HtmlConversionDocument.Parse("<p>Source content</p>").ConvertToPdfWithEvidence(new HtmlToPdfOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        void AddTargetClip(PdfCore.PdfPageCanvas canvas) => canvas.Clip(0D, 0D, 1D, 1D, clipped =>
            clipped.Text("Clipped target content", 10D, 10D, 80D, 30D, fontSize: 12D));
        result.Value.Page(page => page.Size(100D, 100D).Margin(0D).Canvas(canvas => {
            if (nestedUnderSourceClip) canvas.SourceClip(0D, 0D, 100D, 100D, AddTargetClip);
            else AddTargetClip(canvas);
        }));

        using var output = new MemoryStream();
        PdfCore.PdfSaveResult saved = result.TrySaveWithEvidence(output);

        Assert.True(saved.Succeeded);
        Assert.Contains(saved.Report.Warnings, warning => warning.Code == "canvas-text-canvas-clip-paint-clipped"
            && warning.LossKind != OfficeConversionLossKind.None);
        Assert.Throws<InvalidOperationException>(() => saved.Report.RequireNoLoss());
    }

    [Fact]
    public void HtmlPdf_SourceClipDoesNotExemptAdditionalTargetTextFrameLoss() {
        var result = HtmlConversionDocument.Parse("<p>Source content</p>").ConvertToPdfWithEvidence(new HtmlToPdfOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        result.Value.Page(page => page.Size(100D, 100D).Margin(0D).Canvas(canvas =>
            canvas.SourceClip(0D, 0D, 100D, 100D, sourcePaint =>
                sourcePaint.Text("Clipped target content", 10D, 10D, 80D, 1D, fontSize: 12D))));

        using var output = new MemoryStream();
        PdfCore.PdfSaveResult saved = result.TrySaveWithEvidence(output);

        Assert.True(saved.Succeeded);
        Assert.Contains(saved.Report.Warnings, warning => warning.Code == "canvas-text-content-clipped"
            && warning.LossKind != OfficeConversionLossKind.None);
        Assert.Throws<InvalidOperationException>(() => saved.Report.RequireNoLoss());
    }
}
