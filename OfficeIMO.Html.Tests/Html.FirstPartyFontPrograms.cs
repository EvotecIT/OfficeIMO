using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlFirstPartyFontProgramTests {
#if NET8_0_OR_GREATER
    [Fact]
    public void HtmlFontFaceLoaderUsesFirstPartyWoff2ProgramWithoutLoss() {
        byte[] fontData = ReadFont("OpenSans-Regular.woff2");
        string html = "<style>@font-face{font-family:OpenSansWeb;src:url('data:font/woff2;base64,"
            + Convert.ToBase64String(fontData)
            + "') format('woff2')}p{font:24px OpenSansWeb}</style><p>OfficeIMO ffi 123</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));

        OfficeFontFace face = Assert.Single(rendered.Fonts.Faces);
        Assert.Equal("OpenSansWeb", face.FamilyName);
        Assert.Equal(OfficeFontContainerFormat.Woff2, face.ContainerFormat);
        Assert.True(face.CanEmbedAsStaticPdfFont);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.FontFormatUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
        Assert.Contains("OfficeIMO ffi 123", rendered.Text, StringComparison.Ordinal);
    }
#endif

    [Fact]
    public void HtmlPdfEmbedsStaticCff1WithoutVectorFallback() {
        byte[] fontData = ReadFont("SourceSansPro-Regular.otf");
        const string text = "Static serif 123";
        string html = FontHtml("Source Sans CFF", "font/otf", fontData, text, link: false);
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(new HtmlPdfSaveOptions());
        byte[] pdf = result.ToBytes();
        string raw = System.Text.Encoding.GetEncoding(28591).GetString(pdf);

        Assert.Contains(text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("/FontFile3 ", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined);
        Assert.False(result.HasLoss, FormatWarnings(result));
    }

    [Theory]
    [InlineData("Adobe Variable CFF2", "AdobeVFPrototype-Subset.otf", "font/otf", "$$$")]
    [InlineData("Roboto Flex", "RobotoFlex.ttf", "font/ttf", "Variable OfficeIMO")]
    public void HtmlPdfVariableFontsUseAccessibleDeterministicVectorOutlines(
        string family,
        string fileName,
        string mediaType,
        string text) {
        byte[] fontData = ReadFont(fileName);
        HtmlConversionDocument source = HtmlConversionDocument.Parse(FontHtml(family, mediaType, fontData, text, link: false));
        var options = new HtmlPdfSaveOptions();
        if (fileName == "RobotoFlex.ttf") {
            options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 725F, ["wdth"] = 112F };
        }

        PdfCore.PdfDocumentConversionResult first = source.ToPdfDocumentResult(options);
        byte[] firstPdf = first.ToBytes();
        byte[] secondPdf = source.ToPdfDocumentResult(options).ToBytes();

        Assert.Equal(Hash(firstPdf), Hash(secondPdf));
        Assert.Contains(text, PdfCore.PdfReadDocument.Open(firstPdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(first.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined
            && warning.Severity == PdfCore.PdfConversionWarningSeverity.Information
            && warning.Details.TryGetValue("Representation", out string? representation)
            && representation == "vector-outlines-plus-actual-text");
        Assert.False(first.HasLoss, FormatWarnings(first));
        HtmlPdfAccessibilityValidationResult validation = HtmlPdfAccessibilityValidator.Validate(source, firstPdf);
        Assert.True(validation.IsValid, string.Join(" | ", validation.Issues.Select(issue => issue.Code + ": " + issue.Message)));
        var outlineShapes = PdfCore.PdfDocument.Open(firstPdf).Read.Drawing(1).Shapes;
        Assert.NotEmpty(outlineShapes);
        Assert.All(outlineShapes, shape => Assert.Equal(OfficeFillRule.NonZero, shape.Shape.FillRule));
    }

    [Fact]
    public void HtmlPdfTrueTypeCollectionUsesAccessibleVectorOutlines() {
        const string text = "OfficeIMO 0123456789";
        byte[] collection = ManagedTextShapingTestAssets.CreateFontCollection(
            text.Select(character => (int)character).ToArray());
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            FontHtml("Collection Face", "font/ttf", collection, text, link: false));

        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(new HtmlPdfSaveOptions());
        byte[] pdf = result.ToBytes();

        Assert.Contains(text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined
            && warning.Details.TryGetValue("StaticPdfEmbeddable", out string? staticEmbeddable)
            && staticEmbeddable == "false");
        Assert.NotEmpty(PdfCore.PdfDocument.Open(pdf).Read.Drawing(1).Shapes);
        Assert.True(HtmlPdfAccessibilityValidator.Validate(source, pdf).IsValid);
    }

    [Fact]
    public void HtmlPdfVariableFontInsideSvgDrawingUsesVectorOutlines() {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        const string marker = "VariableSvgMarker";
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 220 36'><text x='4' y='26' font-family='Variable Svg' font-size='22'>" + marker + "</text></svg>";
        string html = "<style>@font-face{font-family:'Variable Svg';src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(fontData)
            + "')}</style><img style='width:220px;height:36px' src='data:image/svg+xml;base64,"
            + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg))
            + "' alt='variable font drawing'>";
        var options = new HtmlPdfSaveOptions();
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 725F };

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();

        Assert.Contains(marker, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined
            && warning.Details.TryGetValue("Representation", out string? representation)
            && representation == "vector-outlines-plus-actual-text");
        Assert.NotEmpty(PdfCore.PdfDocument.Open(pdf).Read.Drawing(1).Shapes);
    }

    [Fact]
    public void HtmlFontFaceLoaderAppliesConfiguredVariableFontInstance() {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            FontHtml("Roboto Flex", "font/ttf", fontData, "Variable OfficeIMO", link: false));
        var light = new HtmlPdfSaveOptions();
        light.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 200F, ["wdth"] = 90F };
        var black = new HtmlPdfSaveOptions();
        black.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 900F, ["wdth"] = 120F };

        byte[] lightPdf = source.ToPdfDocumentResult(light).ToBytes();
        byte[] blackPdf = source.ToPdfDocumentResult(black).ToBytes();

        Assert.NotEqual(Hash(lightPdf), Hash(blackPdf));
    }

    [Fact]
    public void HtmlPdfOutlinedVariableFontPaintsConfiguredShapedGlyphs() {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            FontHtml("Roboto Flex", "font/ttf", fontData, "AA", link: false));
        var unshapedOptions = new HtmlPdfSaveOptions();
        unshapedOptions.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 725F };
        var shapedOptions = unshapedOptions.ClonePdf();
        var shapingProvider = new CollapsingTextShapingProvider();
        shapedOptions.TextShapingProvider = shapingProvider;
        shapedOptions.TextShapingLanguage = "pl-PL";

        byte[] unshapedPdf = source.ToPdfDocumentResult(unshapedOptions).ToBytes();
        byte[] shapedPdf = source.ToPdfDocumentResult(shapedOptions).ToBytes();
        int unshapedCommandCount = PdfCore.PdfDocument.Open(unshapedPdf)
            .Read.Drawing(1).Shapes
            .Sum(shape => shape.Shape.PathCommands.Count);
        int shapedCommandCount = PdfCore.PdfDocument.Open(shapedPdf)
            .Read.Drawing(1).Shapes
            .Sum(shape => shape.Shape.PathCommands.Count);

        Assert.True(shapedCommandCount < unshapedCommandCount);
        Assert.True(shapingProvider.Requests.Count >= 2);
        Assert.All(shapingProvider.Requests, request => Assert.Equal(725F, request.VariationCoordinates["wght"]));
        Assert.All(shapingProvider.Requests, request => Assert.Equal("pl-PL", request.Language));
        object cacheKey = Assert.IsAssignableFrom<object>(shapingProvider.Requests[0].FontProgramCacheKeyForShaping);
        Assert.All(shapingProvider.Requests, request => Assert.Same(cacheKey, request.FontProgramCacheKeyForShaping));
    }

    [Theory]
    [InlineData("A\uFE0F")]
    [InlineData("A\u0301")]
    [InlineData("A\u200DB")]
    public void HtmlPdfOutlinedRunReportsProviderDeclineAfterSuccessfulLayoutShaping(string text) {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            FontHtml("Roboto Flex", "font/ttf", fontData, text, link: false));
        var provider = new AcceptOnceThenDeclineTextShapingProvider();
        var options = new HtmlPdfSaveOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "hi"
        };
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 725F };

        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(options);

        Assert.True(provider.Requests.Count >= 2);
        Assert.Contains(result.Report.Warnings, warning =>
            warning.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && warning.Details.TryGetValue("Detail", out string? detail)
            && detail == "provider-declined");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());

        var strictProvider = new AcceptOnceThenDeclineTextShapingProvider();
        options.TextShapingProvider = strictProvider;
        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        HtmlConversionException strict = Assert.Throws<HtmlConversionException>(() =>
            source.ToPdfDocumentResult(options));
        Assert.True(strictProvider.Requests.Count >= 2);
        Assert.Contains(strict.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
    }

    [Fact]
    public void HtmlPdfOutlinedTextNormalizationPreservesVisualAndDecorationOrigins() {
        HtmlPdfRenderedConverter.OutlinedTextFrame frame = HtmlPdfRenderedConverter.ResolveOutlinedTextFrame(
            visualX: 40D,
            visualY: 30D,
            visualHeight: 20D,
            frameWidth: 100D,
            textX: 5D,
            minimumX: -3D,
            minimumY: -4D,
            maximumX: 80D,
            maximumY: 15D);

        Assert.Equal(8D, frame.OffsetX);
        Assert.Equal(4D, frame.OffsetY);
        Assert.Equal(32D, frame.X);
        Assert.Equal(26D, frame.Y);
        Assert.Equal(108D, frame.Width);
        Assert.Equal(24D, frame.Height);
        Assert.Equal(37D, frame.X + -3D + frame.OffsetX);
        Assert.Equal(26D, frame.Y + -4D + frame.OffsetY);
        Assert.Equal(45D, frame.X + 5D + frame.OffsetX);
    }

    [Fact]
    public void HtmlPdfOutlinedFallbackRunsShareOneBaselineAcrossFaceMetrics() {
        HtmlPdfRenderedConverter.OutlinedLineMetrics metrics = HtmlPdfRenderedConverter.ResolveOutlinedLineMetrics(
            visualHeight: 20D,
            lineHeights: new[] { 10D, 12D },
            baselineOffsets: new[] { 8D, 7D });

        Assert.Equal(3.5D, metrics.TextTop);
        Assert.Equal(11.5D, metrics.Baseline);
        Assert.Equal(13D, metrics.LineHeight);
        Assert.Equal(metrics.Baseline, metrics.TextTop + 8D);
        Assert.Equal(metrics.Baseline, (metrics.Baseline - 7D) + 7D);
    }

    [Fact]
    public void HtmlPdfOutlinedGlyphGeometryDoesNotAbsorbCssLetterSpacing() {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        string encodedFont = Convert.ToBase64String(fontData);
        string prefix = "<html><style>@font-face{font-family:'Spacing Variable';src:url('data:font/ttf;base64," +
            encodedFont + "')}p{font-family:'Spacing Variable';font-size:28px;font-variation-settings:'wght' 725;margin:0";
        HtmlConversionDocument plain = HtmlConversionDocument.Parse(prefix + "}</style><p>A</p></html>");
        HtmlConversionDocument spaced = HtmlConversionDocument.Parse(prefix + ";letter-spacing:18px}</style><p>A</p></html>");

        double plainWidth = OutlinedPathWidth(plain.ToPdfDocumentResult().ToBytes());
        double spacedWidth = OutlinedPathWidth(spaced.ToPdfDocumentResult().ToBytes());

        Assert.Equal(plainWidth, spacedWidth, 3);
    }

    [Fact]
    public void HtmlPdfOutlinedFallbackFaceSynthesizesRequestedBoldItalicStyle() {
        byte[] boldLatin = ReadFont("RobotoFlex.ttf");
        byte[] regularGreek = ReadFont("RobotoFlex.ttf");
        string prefix = "<html><style>" +
            "@font-face{font-family:'Fallback Collection';src:url('data:font/ttf;base64," + Convert.ToBase64String(boldLatin) +
            "');font-style:normal;font-weight:700;unicode-range:U+0041}" +
            "@font-face{font-family:'Fallback Collection';src:url('data:font/ttf;base64," + Convert.ToBase64String(regularGreek) +
            "');font-style:normal;font-weight:400;unicode-range:U+03A9}" +
            "p{font-family:'Fallback Collection';font-size:28px;margin:0";
        HtmlConversionDocument regular = HtmlConversionDocument.Parse(prefix + "}</style><p>Ω</p></html>");
        HtmlConversionDocument requested = HtmlConversionDocument.Parse(prefix + ";font-weight:700;font-style:italic}</style><p>Ω</p></html>");

        IReadOnlyList<OfficePathCommand> regularCommands = Assert.Single(
            PdfCore.PdfDocument.Open(regular.ToPdfDocumentResult().ToBytes()).Read.Drawing(1).Shapes).Shape.PathCommands;
        IReadOnlyList<OfficePathCommand> requestedCommands = Assert.Single(
            PdfCore.PdfDocument.Open(requested.ToPdfDocumentResult().ToBytes()).Read.Drawing(1).Shapes).Shape.PathCommands;

        Assert.Equal(regularCommands.Count * 2, requestedCommands.Count);
        Assert.True(OutlinedPathWidth(requestedCommands) > OutlinedPathWidth(regularCommands));
    }

    [Fact]
    public async System.Threading.Tasks.Task HtmlPdfOutlinedVariableFontPropagatesCancellationToShaping() {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            FontHtml("Roboto Flex", "font/ttf", fontData, "AA", link: false));
        using var cancellation = new System.Threading.CancellationTokenSource();
        var shapingProvider = new CollapsingTextShapingProvider((requestNumber, request) => {
            if (requestNumber < 2) return;
            cancellation.Cancel();
            request.CancellationToken.ThrowIfCancellationRequested();
        });
        var options = new HtmlPdfSaveOptions {
            TextShapingProvider = shapingProvider,
            TextShapingLanguage = "pl-PL"
        };
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 725F };

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            source.ToPdfDocumentResultAsync(options, cancellation.Token));

        Assert.True(shapingProvider.Requests.Count >= 2);
        Assert.All(shapingProvider.Requests, request => Assert.Equal("pl-PL", request.Language));
        Assert.All(shapingProvider.Requests, request => Assert.Equal(cancellation.Token, request.CancellationToken));
    }

    [Fact]
    public void HtmlPdfVariableOutlinesRetainHyperlinksAndHonorBudgets() {
        byte[] fontData = ReadFont("RobotoFlex.ttf");
        const string text = "Accessible linked outlines";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(FontHtml("Linked Variable", "font/ttf", fontData, text, link: true));
        var options = new HtmlPdfSaveOptions();
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 800F };

        byte[] pdf = source.ToPdfDocumentResult(options).ToBytes();
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
        Assert.Contains(text, read.ExtractText(), StringComparison.Ordinal);
        Assert.Equal("https://example.com/path", Assert.Single(read.Pages[0].GetLinkAnnotations()).Uri);
        Assert.True(HtmlPdfAccessibilityValidator.Validate(source, pdf).IsValid);

        var characterLimited = options.ClonePdf();
        characterLimited.MaxOutlinedTextCharactersPerRun = 4;
        InvalidOperationException characterException = Assert.Throws<InvalidOperationException>(() => source.ToPdfDocumentResult(characterLimited));
        Assert.Contains("character budget", characterException.Message, StringComparison.Ordinal);
        var commandLimited = options.ClonePdf();
        var shapingProvider = new CollapsingTextShapingProvider();
        commandLimited.TextShapingProvider = shapingProvider;
        commandLimited.MaxOutlinedTextPathCommands = 1;
        InvalidOperationException commandException = Assert.Throws<InvalidOperationException>(() => source.ToPdfDocumentResult(commandLimited));
        Assert.Contains("point budget", commandException.Message, StringComparison.Ordinal);
        Assert.True(shapingProvider.Requests.Count >= 2);
    }

    [Fact]
    public void FallbackPackFingerprintIncludesFirstPartyVariableInstance() {
        byte[] latin = ReadFont("SourceSansPro-Regular.otf");
        byte[] variable = ReadFont("RobotoFlex.ttf");
        OfficeFontFaceCollection lightFonts = CreateFallbackFonts(latin, variable, 200F);
        OfficeFontFaceCollection blackFonts = CreateFallbackFonts(latin, variable, 900F);
        var light = new OfficeFontFallbackPack("portable", "Portable Latin, Portable Variable", lightFonts);
        var identical = new OfficeFontFallbackPack("portable", "Portable Latin, Portable Variable", lightFonts);
        var black = new OfficeFontFallbackPack("portable", "Portable Latin, Portable Variable", blackFonts);

        Assert.Equal(light.Fingerprint, identical.Fingerprint);
        Assert.NotEqual(light.Fingerprint, black.Fingerprint);
        Assert.Equal(64, light.Fingerprint.Length);
        Assert.Equal(light.Fingerprint, new OfficeFontFallbackPack(
            light.Id,
            light.DefaultFamilyNames,
            light.CreateRenderingProfile().Fonts).Fingerprint);
    }

    private static OfficeFontFaceCollection CreateFallbackFonts(byte[] latin, byte[] variable, float weight) {
        var fonts = new OfficeFontFaceCollection { FontVariationResolver = request =>
            request.FamilyName == "Portable Variable" ? new Dictionary<string, float> { ["wght"] = weight } : null };
        return fonts.Add("Portable Latin", latin).Add("Portable Variable", variable).AddFallbackFamily("Portable Variable");
    }

    private static string FontHtml(string family, string mediaType, byte[] data, string text, bool link) {
        string content = link ? "<a href='https://example.com/path'>" + text + "</a>" : "<p>" + text + "</p>";
        return "<html lang='en'><style>@font-face{font-family:'" + family + "';src:url('data:" + mediaType + ";base64,"
            + Convert.ToBase64String(data) + "')}p,a{font-family:'" + family + "';font-size:28px;line-height:1.25}</style>"
            + content + "</html>";
    }

    private static byte[] ReadFont(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", name));

    private static string Hash(byte[] value) {
        using SHA256 sha256 = SHA256.Create();
        return BitConverter.ToString(sha256.ComputeHash(value)).Replace("-", string.Empty);
    }

    private static string FormatWarnings(PdfCore.PdfDocumentConversionResult result) => string.Join(
        " | ",
        result.Report.Warnings.Select(warning => warning.Code + ": " + warning.Message));

    private sealed class CollapsingTextShapingProvider : IOfficeTextShapingProvider {
        private readonly Action<int, OfficeTextShapingRequest>? _onRequest;

        internal CollapsingTextShapingProvider(Action<int, OfficeTextShapingRequest>? onRequest = null) {
            _onRequest = onRequest;
        }

        internal List<OfficeTextShapingRequest> Requests { get; } = new List<OfficeTextShapingRequest>();

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Requests.Add(request);
            _onRequest?.Invoke(Requests.Count, request);
            OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Shaping fixture", request.FontData).Faces);
            Assert.True(face.Program.TryGetGlyphMetrics('A', out int glyphId, out int advanceWidth));
            return new OfficeTextShapingResult(new[] {
                new OfficeShapedGlyph(glyphId, request.Text, 0, advanceWidth)
            });
        }
    }

    private sealed class AcceptOnceThenDeclineTextShapingProvider : IOfficeTextShapingProvider {
        internal List<OfficeTextShapingRequest> Requests { get; } = new List<OfficeTextShapingRequest>();

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Requests.Add(request);
            if (Requests.Count > 1) return null;
            return new OfficeTextShapingResult(new[] {
                new OfficeShapedGlyph(1, request.Text, 0, advanceWidth: 500)
            });
        }
    }

    private static double OutlinedPathWidth(byte[] pdf) => OutlinedPathWidth(Assert.Single(
        PdfCore.PdfDocument.Open(pdf).Read.Drawing(1).Shapes).Shape.PathCommands);

    private static double OutlinedPathWidth(IReadOnlyList<OfficePathCommand> commands) {
        double minimum = commands.Where(command => command.Kind != OfficePathCommandKind.Close).Min(command => command.Point.X);
        double maximum = commands.Where(command => command.Kind != OfficePathCommandKind.Close).Max(command => command.Point.X);
        return maximum - minimum;
    }
}
