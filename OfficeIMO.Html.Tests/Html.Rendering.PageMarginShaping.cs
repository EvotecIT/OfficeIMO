using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top-left")]
    [InlineData("top-center")]
    [InlineData("top-right")]
    [InlineData("bottom-left")]
    [InlineData("bottom-center")]
    [InlineData("bottom-right")]
    [InlineData("left-top")]
    [InlineData("left-middle")]
    [InlineData("left-bottom")]
    [InlineData("right-top")]
    [InlineData("right-middle")]
    [InlineData("right-bottom")]
    [InlineData("top-left-corner")]
    [InlineData("top-right-corner")]
    [InlineData("bottom-left-corner")]
    [InlineData("bottom-right-corner")]
    public void HtmlPageMarginShaping_EveryPositionRetainsTheSameArabicGlyphsAsBodyText(string position) {
        const string text = "لَا";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource(position, text, "width:max-content", text)), PageMarginShapingOptions());

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup artifact = PageMarginGeometryBox(page, position);
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, artifact.Role);
        HtmlRenderText margin = Assert.Single(PageMarginShapedText(artifact));
        HtmlRenderText control = Assert.Single(PageMarginBodyText(page));
        Assert.NotNull(control.PreparedShaping);
        Assert.NotNull(margin.PreparedShaping);
        Assert.Equal(text, margin.PreparedShaping!.Text);
        Assert.Equal(control.PreparedShaping!.Result.Glyphs.Select(glyph =>
            (glyph.GlyphId, glyph.AdvanceWidth, glyph.OffsetX, glyph.OffsetY, glyph.UnicodeText)),
            margin.PreparedShaping.Result.Glyphs.Select(glyph =>
                (glyph.GlyphId, glyph.AdvanceWidth, glyph.OffsetX, glyph.OffsetY, glyph.UnicodeText)));
        Assert.Equal(margin.PreparedShaping.Measure(margin.Font.Size), margin.TextAdvanceWidth!.Value, 6);
        Assert.Equal(text, rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPageMarginShaping_MarkedCffMaxContentFitsAndExportsTheSameGlyphScene(bool explicitProvider) {
        const string text = "Cafe\u0301";
        HtmlToPdfOptions options = PageMarginShapingOptions(cff: true);
        if (explicitProvider) options.TextShapingProvider = OfficeManagedTextShapingProvider.Instance;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", text, "width:max-content;height:max-content", text));

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
        HtmlRenderSemanticGroup artifact = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderText margin = Assert.Single(PageMarginShapedText(artifact));
        Assert.NotNull(margin.PreparedShaping);
        Assert.Equal(text, margin.PreparedShaping!.Text);
        Assert.Equal(margin.PreparedShaping.Measure(margin.Font.Size), artifact.Width, 6);
        Assert.Equal(24D, artifact.Height, 6);
        rendered.RequireNoLoss();

        PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
        pdf.RequireNoLoss();
        byte[] saved = pdf.ToBytes();
        Assert.Equal(text, PdfCore.PdfReadDocument.Open(saved).ExtractText().Trim());
        OfficeImageExportResult image = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(pdf,
            OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
        Assert.Equal(replay.Bytes, image.Bytes);
        Assert.DoesNotContain(image.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
    }

    [Fact]
    public void HtmlPageMarginShaping_SavedArabicArtifactRetainsItsLogicalClusterOnce() {
        const string text = "لَا";
        HtmlToPdfOptions options = PageMarginShapingOptions();
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            PageMarginShapingSource("bottom-center", text, "width:max-content", string.Empty));
        PdfCore.PdfDocumentConversionResult pdf = document.ToPdfDocumentResult(options);
        pdf.RequireNoLoss();
        byte[] bytes = pdf.ToBytes();

        Assert.True(string.IsNullOrWhiteSpace(PdfCore.PdfReadDocument.Open(bytes).ExtractText()));
        PdfCore.PdfReadDocument includingArtifacts = PdfCore.PdfReadDocument.Open(bytes,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true });
        Assert.Equal(text, includingArtifacts.ExtractText().Trim());
        Assert.Equal(text, string.Concat(Assert.Single(includingArtifacts.Pages).GetTextSpans().Select(span => span.Text)));
    }

    [Fact]
    public void HtmlPageMarginShaping_RegisteredFallbackKeepsEachFontAndItsPositionedMarks() {
        HtmlToPdfOptions options = PageMarginShapingOptions();
        options.Fonts.Add("Latin", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf")));
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", "A\u0301 لَا", "width:max-content;font-family:Latin,Typography", string.Empty)), options);
        HtmlRenderText[] text = PageMarginShapedText(PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center"));

        Assert.Contains(text, run => run.Text == "A\u0301" && run.PreparedShaping?.Program is IOfficeCffBoundedFontProgram);
        Assert.Contains(text, run => run.Text == "لَا" && run.PreparedShaping?.Program is OfficeTrueTypeFont);
        Assert.All(text, run => {
            Assert.NotNull(run.PreparedShaping);
            Assert.Equal(run.PreparedShaping!.Measure(run.Font.Size), run.TextAdvanceWidth!.Value, 6);
        });
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginShaping_AuthoredHardLinesRetainLeadingAndTheFinalEmptyLine() {
        HtmlToPdfOptions options = PageMarginShapingOptions(cff: true);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", "A\u0301\\A A\u0301\\A ",
                "width:80px;height:max-content", "A\u0301")), options);

        HtmlRenderSemanticGroup artifact = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderText[] text = PageMarginShapedText(artifact);
        Assert.Equal(72D, artifact.Height, 6);
        Assert.Equal(2, text.Select(run => run.LayoutY).Distinct().Count());
        Assert.All(text, run => {
            Assert.Equal("A\u0301", run.Text);
            Assert.NotNull(run.PreparedShaping);
        });
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(12D)]
    public void HtmlPageMarginShaping_ShortLeadingKeepsCompletePositiveGlyphFrames(double leading) {
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", "A\u0301\\A A\u0301", "width:80px;line-height:"
                + leading.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px", "A\u0301")),
            PageMarginShapingOptions(cff: true));
        HtmlRenderText[] text = PageMarginShapedText(PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center"));

        Assert.Equal(2, text.Length);
        Assert.All(text, run => {
            Assert.Equal(leading, run.LineHeight, 6);
            Assert.True(run.Height > 0D);
            Assert.NotNull(run.PreparedShaping);
        });
        Assert.Equal(leading, text[1].LayoutY - text[0].LayoutY, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginShaping_WrappingKeepsEachCombiningClusterAndItsMeasuredAdvance() {
        HtmlToPdfOptions options = PageMarginShapingOptions(cff: true);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", "A\u0301 A\u0301", "width:min-content;height:max-content", "A\u0301")), options);
        HtmlRenderSemanticGroup artifact = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderText[] text = PageMarginShapedText(artifact);

        Assert.Equal(2, text.Length);
        Assert.Equal(2, text.Select(run => run.LayoutY).Distinct().Count());
        Assert.Equal(48D, artifact.Height, 6);
        Assert.All(text, run => {
            Assert.Equal("A\u0301", run.Text);
            Assert.NotNull(run.PreparedShaping);
            Assert.Equal(run.PreparedShaping!.Measure(run.Font.Size), run.TextAdvanceWidth!.Value, 6);
        });
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginShaping_EmergencyWrapNeverCutsALamAlefClusterIntoSeparateLetters() {
        const string text = "لَا";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", text, "width:1px", text)), PageMarginShapingOptions());
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = Assert.Single(PageMarginShapedText(PageMarginGeometryBox(page, "top-center")));
        HtmlRenderText body = Assert.Single(PageMarginBodyText(page));

        Assert.NotNull(margin.PreparedShaping);
        Assert.Equal(text, margin.PreparedShaping!.Text);
        Assert.NotNull(body.PreparedShaping);
        Assert.Equal(body.PreparedShaping!.Result.Glyphs.Select(glyph => glyph.GlyphId),
            margin.PreparedShaping.Result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(body.TextAdvanceWidth, margin.TextAdvanceWidth);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginShaping_OneOverlongMarkedTokenStopsAtTheSharedLineAdmission() {
        string content = string.Concat(Enumerable.Repeat("A\u0301", OfficeTextLayoutEngine.MaximumLayoutLines + 8));
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", content, "width:1px;line-height:0", string.Empty)),
            PageMarginShapingOptions(cff: true));

        Assert.Contains(rendered.Diagnostics, item => item.Source == "@page @top-center"
            && item.Detail == "text-layout-limit" && item.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(OfficeTextLayoutEngine.MaximumLayoutLines,
            PageMarginShapedText(PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center")).Length);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("ب")]
    [InlineData("A")]
    public void HtmlPageMarginShaping_CustomProviderControlsBothIntrinsicWidthAndRetainedPaint(string source) {
        byte[] font = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            source == "A" ? "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf"
                : "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf"));
        HtmlToPdfOptions options = PageMarginShapingOptions(fontData: font);
        var provider = new LanguageAdvanceTextShapingProvider(font);
        // The provider's exact program must cover this positive fixture. Noto
        // Arabic deliberately has no Latin A and correctly declines that text.
        Assert.True(provider.Program.TryGetGlyphMetrics(source[0], out int glyphId, out _));
        options.TextShapingProvider = provider;
        options.TextShapingLanguage = "fa";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", source, "width:max-content", source)), options);
        HtmlRenderSemanticGroup artifact = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderText margin = Assert.Single(PageMarginShapedText(artifact));
        double expected = 18D * LanguageAdvanceTextShapingProvider.Advance("fa") / provider.Program.UnitsPerEm;

        Assert.Equal(expected, artifact.Width, 6);
        Assert.NotNull(margin.PreparedShaping);
        Assert.Equal(glyphId, Assert.Single(margin.PreparedShaping!.Result.Glyphs).GlyphId);
        Assert.Equal(expected, margin.TextAdvanceWidth!.Value, 6);
        Assert.NotEmpty(provider.Requests);
        Assert.All(provider.Requests, request => Assert.Equal("fa", request.Language));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginShaping_DeclinedGeneratedTextHasAStableSourceAndCannotClaimNoLoss() {
        HtmlToPdfOptions options = PageMarginShapingOptions();
        options.TextShapingProvider = new DecliningPageMarginShapingProvider();
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", "لَا", "width:80px", string.Empty)), options);

        Assert.Contains(rendered.Diagnostics, item => item.Source == "@page @top-center"
            && item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && item.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginShaping_ZeroSizeNeverCallsTheConfiguredProvider() {
        var provider = new DecliningPageMarginShapingProvider();
        HtmlToPdfOptions options = PageMarginShapingOptions();
        options.TextShapingProvider = provider;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            PageMarginShapingSource("top-center", "لَا\\A لَا", "font-size:0;line-height:12px;height:max-content", string.Empty)), options);

        Assert.Equal(0, provider.Calls);
        Assert.Empty(PageMarginShapedText(PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center")));
        rendered.RequireNoLoss();
    }

    private static HtmlToPdfOptions PageMarginShapingOptions(bool cff = false, byte[]? fontData = null) {
        byte[] data = fontData ?? File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), cff
            ? "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf"
            : "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf"));
        var options = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, AllowSystemFontFallback = false,
            DefaultFontFamily = "Typography", FontFamily = new PdfCore.PdfEmbeddedFontFamily("Typography", data),
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("Typography", data);
        return options;
    }

    private static string PageMarginShapingSource(string position, string generated, string declaration, string body) =>
        "<style>@page{size:480px 480px;margin:100px;@" + position + "{content:'" + generated
        + "';font:18px/24px Typography;" + declaration + "}}html,body,p{margin:0;padding:0}"
        + "p{font:18px/24px Typography}</style><p id='control'>" + body + "</p>";

    private static HtmlRenderText[] PageMarginShapedText(HtmlRenderSemanticGroup artifact) =>
        EnumerateRenderVisuals(artifact.Visuals).OfType<HtmlRenderText>().Where(run => !string.IsNullOrWhiteSpace(run.Text)).ToArray();

    private static HtmlRenderText[] PageMarginBodyText(HtmlRenderPage page) {
        // Ordinary text uses semantic source names unless an interaction owns a
        // DOM identity. The Artifact scope, rather than an invented source name,
        // distinguishes the body control from the identically worded header.
        var artifacts = new HashSet<HtmlRenderText>(EnumerateRenderVisuals(page.Scene)
            .OfType<HtmlRenderSemanticGroup>().Where(group => group.Role == HtmlRenderSemanticGroupRole.Artifact)
            .SelectMany(group => EnumerateRenderVisuals(group.Visuals)).OfType<HtmlRenderText>());
        return EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>()
            .Where(run => !artifacts.Contains(run) && !string.IsNullOrWhiteSpace(run.Text)).ToArray();
    }

    private sealed class DecliningPageMarginShapingProvider : IOfficeTextShapingProvider {
        internal int Calls { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            request.CancellationToken.ThrowIfCancellationRequested();
            Calls++;
            return null;
        }
    }
}
