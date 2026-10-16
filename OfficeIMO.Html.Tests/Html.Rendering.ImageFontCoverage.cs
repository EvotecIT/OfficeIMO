using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlImageFontCoverageTests {
    [Fact]
    public void HtmlImages_UncoveredJoiningScriptUsesThePdfEncodingFailureContract() {
        const string html = "<p dir='rtl'>ܫܠܡ</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
        };

        HtmlConversionException pdf = Assert.Throws<HtmlConversionException>(() => document.ToPdfBytes(options));
        HtmlConversionException png = Assert.Throws<HtmlConversionException>(() => document.ToPng(options));

        foreach (HtmlConversionException failure in new[] { pdf, png }) {
            HtmlDiagnostic[] encoding = failure.Diagnostics.Where(item => item.Code == "unsupported-text-glyph").ToArray();
            Assert.NotEmpty(encoding);
            Assert.All(encoding, diagnostic => {
                Assert.Equal(HtmlDiagnosticSeverity.Error, diagnostic.Severity);
                Assert.Equal(OfficeConversionLossKind.Failure, diagnostic.LossKind);
            });
            Assert.Contains(encoding, diagnostic => diagnostic.Message.Contains("U+072B", StringComparison.Ordinal));
        }
        Assert.Equal(pdf.Diagnostics.Select(item => item.Code), png.Diagnostics.Select(item => item.Code));
        Assert.Equal(html, document.SourceHtml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async System.Threading.Tasks.Task HtmlImages_StrictPolicyRejectsAProviderThatDeclinesPdfPainting(bool pdfRoute) {
        const string text = "ܫܠܡ";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">" + text + "</p>");
        var provider = new LayoutOnlyShapingProvider();
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(text.Select(character => (int)character).ToArray()));

        HtmlConversionException synchronous = Assert.Throws<HtmlConversionException>(() => {
            if (pdfRoute) document.ToPdfBytes(options);
            else document.ToPng(options);
        });
        HtmlConversionException asynchronous = await Assert.ThrowsAsync<HtmlConversionException>(async () => {
            if (pdfRoute) await document.ToPdfBytesAsync(options);
            else await document.ToPngAsync(options);
        });

        foreach (HtmlConversionException failure in new[] { synchronous, asynchronous }) {
            Assert.Contains(failure.Diagnostics, diagnostic => diagnostic.Code == "unsupported-complex-script-shaping"
                && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        }
        Assert.True(provider.AcceptedLayout);
        Assert.True(provider.DeclinedPainting);
    }

    [Fact]
    public void HtmlImages_CallerFontAndShaperCoverJoiningScriptOnBothPdfRoutes() {
        const string text = "ܫܠܡ";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:120px 60px;margin:0}body{margin:0}</style>"
            + "<p dir='rtl' style=\"font-family:'OfficeIMO Shaping Test'\">" + text + "</p>");
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(text.Select(character => (int)character).ToArray()));

        PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult pdfPng = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
            pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));

        Assert.Contains(text, PdfCore.PdfReadDocument.Open(pdf.ToBytes()).ExtractText(), StringComparison.Ordinal);
        Assert.Equal(pdfPng.Bytes, png.Bytes);
        Assert.DoesNotContain(png.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Contains(provider.Requests, request => request.Text == text && request.Language == "syr");
    }

    private sealed class LayoutOnlyShapingProvider : IOfficeTextShapingProvider {
        private readonly ManagedTextShapingTestAssets.RecordingProvider _layout = new();
        internal bool AcceptedLayout { get; private set; }
        internal bool DeclinedPainting { get; private set; }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            if (request.FontName != ManagedTextShapingTestAssets.FamilyName) {
                DeclinedPainting = true;
                return null;
            }
            AcceptedLayout = true;
            return _layout.ShapeText(request);
        }
    }
}
