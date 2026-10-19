using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlDirectionalProvider_EmptySpaceContoursRetainLogicalTextWithoutPdfLoss() {
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss,
            TextShapingProvider = new ManagedTextShapingTestAssets.RecordingProvider()
        };
        options.Fonts.Add("Ink", ManagedTextShapingTestAssets.CreateFont("אבג123".Select(c => (int)c).ToArray()));
        options.Fonts.Add("Empty Space", ManagedTextShapingTestAssets.CreateFontWithEmptyGlyphs(2, ' '));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:300px 120px;margin:0}html,body,p{margin:0;padding:0}"
            + "p{font:20px/24px 'Ink'}</style><p dir='rtl'><span>אבג</span>"
            + "<span style=\"font-family:'Empty Space'\"> </span><span>123</span></p>");
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
        HtmlRenderText space = Assert.Single(ArabicContextText(rendered), run => run.Text == " ");
        Assert.NotNull(space.PreparedShaping);
        Assert.True(space.TextAdvanceWidth > 0D);
        Assert.Empty(space.PreparedShaping!.GetContours(0D, 0D, 20D, 1000, default));
        rendered.RequireNoLoss();

        PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
        byte[] bytes = pdf.ToBytes();
        pdf.RequireNoLoss();
        Assert.Equal("אבג 123", PdfCore.PdfReadDocument.Open(bytes).ExtractText().Trim());
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
            pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
        Assert.Equal(replay.Bytes, png.Bytes);
        Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);
    }
}
