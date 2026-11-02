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

        HtmlConversionException pdf = Assert.Throws<HtmlConversionException>(() => (document.ToPdf(options)).ToBytes());
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
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async System.Threading.Tasks.Task HtmlImages_StrictPolicyRetainsLayoutShapingWithoutAliasReshaping(bool pdfRoute, bool styled) {
        const string text = "ܫܠܡ";
        string expectedText = styled ? text + " " + text : text;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:120px 60px;margin:0}body,p{margin:0}</style>"
            + "<p dir='rtl' style=\"font-family:'OfficeIMO Shaping Test'\">"
            + (styled ? "<span style='color:red'>" + text + " </span><span>" + text + "</span>" : text) + "</p>");
        var provider = new LayoutOnlyShapingProvider();
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs((text + " ").Select(character => (int)character).ToArray()));

        byte[] synchronous = pdfRoute ? (document.ToPdf(options)).ToBytes() : document.ToPng(options);
        byte[] asynchronous = pdfRoute ? (await document.ToPdfAsync(options)).Value.ToBytes() : await document.ToPngAsync(options);

        Assert.NotEmpty(synchronous);
        Assert.NotEmpty(asynchronous);
        if (pdfRoute) {
            foreach (byte[] bytes in new[] { synchronous, asynchronous }) {
                Assert.Equal(expectedText, PdfCore.PdfReadDocument.Open(bytes).ExtractText().Trim());
            }
        } else {
            Assert.Equal(synchronous, asynchronous);
            PdfCore.PdfDocumentConversionResult pdf = document.ConvertToPdfWithEvidence(options);
            pdf.RequireNoLoss();
            Assert.Equal(expectedText, PdfCore.PdfReadDocument.Open(pdf.ToBytes()).ExtractText().Trim());
            OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
                pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
            Assert.Equal(replay.Bytes, synchronous);
        }
        Assert.True(provider.AcceptedLayout);
        Assert.False(provider.DeclinedAliasShaping);
    }

    [Theory]
    [InlineData("ܫܠܡ", false)]
    [InlineData("ܫܠܡ", true)]
    [InlineData("אבג", true)]
    public void HtmlImages_CallerFontAndShaperCoverJoiningScriptOnBothPdfRoutes(string text, bool styled) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:120px 60px;margin:0}body{margin:0}</style>"
            + "<p dir='rtl' style=\"font-family:'OfficeIMO Shaping Test'\">"
            + (styled ? "<span style='color:red'>" + text + " </span><span>" + text + "</span>" : text) + "</p>");
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.PdfOptions.Language = "en-GB";
        byte[] fontData = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            (text + " ").Select(character => (int)character).ToArray());
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, fontData);

        PdfCore.PdfDocumentConversionResult pdf = document.ConvertToPdfWithEvidence(options);
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult pdfPng = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
            pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));

        Assert.Equal("en-GB", pdf.Value.Options.Language);
        Assert.Equal(styled ? text + " " + text : text,
            PdfCore.PdfReadDocument.Open(pdf.ToBytes()).ExtractText().Trim());
        Assert.Equal(pdfPng.Bytes, png.Bytes);
        Assert.DoesNotContain(png.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Contains(provider.Requests, request => request.Text.Contains(text, StringComparison.Ordinal)
            && request.FontData.SequenceEqual(fontData) && request.Language == "syr"
            && request.Direction == OfficeTextDirection.RightToLeft);
    }

    [Theory]
    [InlineData('\u202F')]
    [InlineData('\u2007')]
    public void HtmlImages_RegisteredSpacingFallbackPreservesTextAndPdfReplay(char separator) {
        string text = "A" + separator + "B";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:160px 60px;margin:0}body{margin:0}p{margin:0;font:20px Primary;white-space:pre-wrap}</style><p>"
            + text + "</p>");
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };
        options.Fonts.Add("Primary", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B'))
            .Add("Spacing Fallback", ManagedTextShapingTestAssets.CreateFontWithEmptyGlyphs(2, separator))
            .AddFallbackFamily("Spacing Fallback");

        PdfCore.PdfDocumentConversionResult pdf = document.ConvertToPdfWithEvidence(options);
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
            pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));

        PdfCore.PdfReadDocument reopened = PdfCore.PdfReadDocument.Open(pdf.ToBytes());
        Assert.Equal("A B", reopened.ExtractText()); // Plain extraction normalizes whitespace.
        Assert.Equal(text, string.Concat(Assert.Single(reopened.Pages).GetGlyphTextSpans().Select(span => span.Text)));
        Assert.Equal(replay.Bytes, png.Bytes);
        Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == "unsupported-text-glyph");
    }

    [Theory]
    [InlineData('\u202F')]
    [InlineData('\u2007')]
    public void HtmlImages_UncoveredSpacingGlyphStillFailsWithTypedLoss(char separator) {
        string html = "<p style='font-family:Primary;white-space:pre-wrap'>A" + separator + "B</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
        };
        options.Fonts.Add("Primary", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B'));

        HtmlConversionException pdf = Assert.Throws<HtmlConversionException>(() => (document.ToPdf(options)).ToBytes());
        HtmlConversionException png = Assert.Throws<HtmlConversionException>(() => document.ToPng(options));

        foreach (HtmlConversionException failure in new[] { pdf, png }) {
            Assert.Contains(failure.Diagnostics, diagnostic => diagnostic.Code == "unsupported-text-glyph"
                && diagnostic.Severity == HtmlDiagnosticSeverity.Error
                && diagnostic.LossKind == OfficeConversionLossKind.Failure
                && diagnostic.Message.Contains("U+" + ((int)separator).ToString("X4"), StringComparison.Ordinal));
        }
        Assert.Equal(html, document.SourceHtml);
    }

    private sealed class LayoutOnlyShapingProvider : IOfficeTextShapingProvider {
        private readonly ManagedTextShapingTestAssets.RecordingProvider _layout = new();
        internal bool AcceptedLayout { get; private set; }
        internal bool DeclinedAliasShaping { get; private set; }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            if (request.FontName != ManagedTextShapingTestAssets.FamilyName) {
                DeclinedAliasShaping = true;
                return null;
            }
            AcceptedLayout = true;
            return _layout.ShapeText(request);
        }
    }
}
