using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("\u2067אבג\u2069", "אבג")]
    [InlineData("\u2068١٢٣\u2069", "١٢٣")]
    [InlineData("LEFT\u2067אבג\u2069", "LEFTאבג")]
    [InlineData("<bdi dir='rtl'>אבג 123</bdi>", "אבג 123")]
    public void HtmlLocalBidi_ResolvedPaintKeepsItsOwnershipThroughPdfAndImages(string content, string logicalText) {
        HtmlConversionDocument document = LocalBidiDocument(content);
        HtmlToPdfOptions options = LocalBidiOptions();
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
        HtmlRenderText[] directionalPaint = LocalBidiText(rendered)
            .Where(text => OfficeTextElements.ContainsRightToLeft(text.Text)).ToArray();
        Assert.NotEmpty(directionalPaint);
        Assert.All(directionalPaint, text => {
            Assert.True(text.BidiVisualOrderResolved);
            Assert.True(text.PreservesPaintedGlyphs);
        });

        PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
        byte[] bytes = pdf.ToBytes();
        pdf.RequireNoLoss();
        Assert.Equal(logicalText, PdfCore.PdfReadDocument.Open(bytes).ExtractText().Trim());
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
            pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
        Assert.Equal(replay.Bytes, png.Bytes);
        Assert.DoesNotContain(png.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
    }

    [Fact]
    public void HtmlLocalBidi_ConfiguredProviderStillReceivesLogicalDirectionalText() {
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        HtmlToPdfOptions options = LocalBidiOptions();
        options.TextShapingProvider = provider;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(LocalBidiDocument("\u2067אבג\u2069"), options);
        HtmlRenderText[] directional = LocalBidiText(rendered)
            .Where(text => OfficeTextElements.ContainsRightToLeft(text.Text)).ToArray();

        Assert.NotEmpty(directional);
        Assert.All(directional, text => {
            Assert.NotNull(text.PreparedShaping);
            Assert.Equal(text.Text, text.PreparedShaping!.Text);
        });
        Assert.Contains(provider.Requests, request => request.Text == "אבג");
        Assert.DoesNotContain(provider.Requests, request => request.Text == "גבא");
    }

    private static HtmlToPdfOptions LocalBidiOptions() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            " LEFTאבג123١٢٣".Select(character => (int)character).Distinct()
                .ToDictionary(scalar => scalar, _ => (ushort)500), new[] { 32 });
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss,
            FontFamily = new PdfCore.PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, font)
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, font);
        return options;
    }

    private static HtmlConversionDocument LocalBidiDocument(string content) => HtmlConversionDocument.Parse(
        "<style>@page{size:300px 120px;margin:0}html,body{margin:0;padding:0}p{margin:0;"
        + "font:20px/24px '" + ManagedTextShapingTestAssets.FamilyName + "'}</style><p dir='ltr'>" + content + "</p>");

    private static IEnumerable<HtmlRenderText> LocalBidiText(HtmlRenderDocument document) => document.Pages
        .SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>();
}
