using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
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
        Assert.NotEmpty(PdfCore.PdfDocument.Open(firstPdf).Read.Drawing(1).Shapes);
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
}
