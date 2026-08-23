using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.SixLabors;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Drawing.SixLabors.Tests;

public sealed class DrawingSixLaborsFontProgramTests {
    [Fact]
    public void Woff1Face_NormalizesProviderInputWithoutLosingSourceContainerIdentity() {
        byte[] openType = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(openType);
        var fonts = new OfficeFontFaceCollection()
            .UseSixLaborsFontPrograms()
            .Add("Source Sans WOFF", woff);

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.Woff, face.ContainerFormat);
        Assert.True(face.CanEmbedAsStaticPdfFont);
        Assert.True(face.Program.HasGlyphs("OfficeIMO ffi"));
        Assert.True(face.Program.Measure("OfficeIMO ffi", 24D) > 0D);
    }

    [Fact]
    public void Woff2Face_MeasuresAndRasterizesThroughOptionalProgram() {
        byte[] woff2 = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        var fonts = new OfficeFontFaceCollection()
            .UseSixLaborsFontPrograms();
        Assert.True(fonts.TryAddBounded(
            "Open Sans WOFF2",
            woff2,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: 16 * 1024 * 1024,
            out int decodedBytes,
            out string? error), error);
        Assert.True(decodedBytes > woff2.Length);

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.Woff2, face.ContainerFormat);
        Assert.False(face.CanEmbedAsStaticPdfFont);
        Assert.True(fonts.TryMeasureText(
            "OfficeIMO ffi 123",
            24D,
            "Open Sans WOFF2",
            OfficeFontStyle.Regular,
            out double width));
        Assert.True(width > 100D);

        var image = new OfficeRasterImage(360, 80, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);
        canvas.DrawText(
            "OfficeIMO ffi 123",
            0D,
            0D,
            image.Width,
            image.Height,
            OfficeColor.Black,
            36D,
            fontFamily: "Open Sans WOFF2");

        byte[] pixels = image.GetPixels();
        Assert.Contains(pixels, value => value < 250);
    }

    [Fact]
    public void HtmlFontFaceLoader_UsesConfiguredWoff2ProgramWithoutLossDiagnostic() {
        byte[] woff2 = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        string html = "<style>@font-face{font-family:OpenSansWeb;src:url('data:font/woff2;base64,"
            + Convert.ToBase64String(woff2)
            + "') format('woff2')}p{font:24px OpenSansWeb}</style><p>OfficeIMO ffi 123</p>";
        var options = new HtmlRenderOptions()
            .UseSixLaborsFontPrograms();

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html),
            options);

        OfficeFontFace face = Assert.Single(rendered.Fonts.Faces);
        Assert.Equal("OpenSansWeb", face.FamilyName);
        Assert.Equal(OfficeFontContainerFormat.Woff2, face.ContainerFormat);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.FontFormatUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable
            || diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Contains("OfficeIMO ffi 123", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Cff1Face_RemainsEligibleForStaticPdfEmbedding() {
        OfficeFontFaceCollection fonts = LoadFont(
            "Source Sans Pro CFF",
            "SourceSansPro-Regular.otf",
            OfficeSixLaborsFontProgramProvider.Instance);

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.OpenType, face.ContainerFormat);
        Assert.True(face.CanEmbedAsStaticPdfFont);
        AssertInk(fonts, "Source Sans Pro CFF", "CFF office ffi");
    }

    [Fact]
    public void HtmlPdf_StaticCff1UsesEmbeddedFontInsteadOfVectorFallback() {
        byte[] fontData = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        const string text = "Static serif 123";
        string html = "<html lang='en'><style>@font-face{font-family:'Source Sans CFF';src:url('data:font/otf;base64,"
            + Convert.ToBase64String(fontData)
            + "')}p{font:28px 'Source Sans CFF'}</style><p>"
            + text
            + "</p></html>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions().UseSixLaborsFontPrograms();

        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string raw = System.Text.Encoding.Latin1.GetString(pdf);

        Assert.Contains(text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("/FontFile3 ", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined);
        Assert.False(
            result.HasLoss,
            string.Join(" | ", result.Report.Warnings.Select(warning =>
                warning.Code + ": " + warning.Message)));
    }

    [Fact]
    public void HtmlPdf_StaticCff1UsesProviderOutlinesWhenScalarPdfTextWouldLoseLigatures() {
        byte[] fontData = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        const string text = "Office ffi affinity";
        string html = "<html lang='en'><style>@font-face{font-family:'Source Sans CFF';src:url('data:font/otf;base64,"
            + Convert.ToBase64String(fontData)
            + "')}p{font:28px 'Source Sans CFF'}</style><p>"
            + text
            + "</p></html>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions().UseSixLaborsFontPrograms();

        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();

        Assert.Contains(text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined
            && warning.Severity == PdfCore.PdfConversionWarningSeverity.Information
            && warning.Details.TryGetValue("StaticPdfEmbeddable", out string? value)
            && value == "true");
        Assert.DoesNotContain(result.Report.Warnings, warning =>
            warning.Code == "unsupported-font-ligature-substitution");
        Assert.False(
            result.HasLoss,
            string.Join(" | ", result.Report.Warnings.Select(warning =>
                warning.Code + ": " + warning.Message)));
        HtmlPdfAccessibilityValidationResult validation = HtmlPdfAccessibilityValidator.Validate(source, pdf);
        Assert.True(validation.IsValid, string.Join(" | ", validation.Issues.Select(issue => issue.Code + ": " + issue.Message)));
    }

    [Fact]
    public void Cff2VariableFace_RasterizesButFailsClosedForStaticPdfEmbedding() {
        var provider = new OfficeSixLaborsFontProgramProvider(_ =>
            new Dictionary<string, float> {
                ["wght"] = 700F,
                ["CNTR"] = 75F
            });
        OfficeFontFaceCollection fonts = LoadFont(
            "Adobe Variable CFF2",
            "AdobeVFPrototype-Subset.otf",
            provider);

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.False(face.CanEmbedAsStaticPdfFont);
        AssertInk(fonts, "Adobe Variable CFF2", "$$$");
    }

    [Fact]
    public void TrueTypeVariableAxes_ProduceDistinctDeterministicOutlines() {
        var lightProvider = new OfficeSixLaborsFontProgramProvider(_ =>
            new Dictionary<string, float> { ["wght"] = 200F, ["wdth"] = 75F });
        var blackProvider = new OfficeSixLaborsFontProgramProvider(_ =>
            new Dictionary<string, float> { ["wght"] = 900F, ["wdth"] = 125F });
        OfficeFontFaceCollection light = LoadFont("Roboto Flex", "RobotoFlex.ttf", lightProvider);
        OfficeFontFaceCollection black = LoadFont("Roboto Flex", "RobotoFlex.ttf", blackProvider);

        Assert.False(Assert.Single(light.Faces).CanEmbedAsStaticPdfFont);
        Assert.False(Assert.Single(black.Faces).CanEmbedAsStaticPdfFont);
        byte[] lightHash = RenderHash(light, "Roboto Flex", "Variable OfficeIMO");
        byte[] blackHash = RenderHash(black, "Roboto Flex", "Variable OfficeIMO");
        Assert.NotEqual(Convert.ToHexString(lightHash), Convert.ToHexString(blackHash));
        Assert.NotEqual(
            Assert.Single(light.Faces).Program.Fingerprint,
            Assert.Single(black.Faces).Program.Fingerprint);
        Assert.Equal(
            Convert.ToHexString(lightHash),
            Convert.ToHexString(RenderHash(light, "Roboto Flex", "Variable OfficeIMO")));
    }

    [Theory]
    [InlineData("Noto Arabic", "NotoSansArabic-Regular.ttf", "مرحبا بالعالم")]
    [InlineData("Noto Devanagari", "NotoSansDevanagari-Regular.ttf", "नमस्ते दुनिया")]
    [InlineData("Noto CJK", "NotoSansSC-BaselineSubset.ttf", "永字国")]
    [InlineData("Noto Emoji", "NotoEmoji-VariableFont_wght.ttf", "😀🚀🌍")]
    public void ComplexTextCorpus_UsesDeterministicManagedShaping(
        string family,
        string fileName,
        string text) {
        OfficeFontFaceCollection fonts = LoadFont(
            family,
            fileName,
            OfficeSixLaborsFontProgramProvider.Instance);

        Assert.Equal(family, Assert.Single(fonts.PlanFallbackRuns(text, family)).FamilyName);
        byte[] first = RenderHash(fonts, family, text);
        byte[] second = RenderHash(fonts, family, text);
        Assert.Equal(Convert.ToHexString(first), Convert.ToHexString(second));
        AssertInk(fonts, family, text);
    }

    [Theory]
    [InlineData("❤️")]
    [InlineData("👨‍👩‍👧‍👦")]
    public void EmojiFallbackCoverage_AllowsVariationSelectorsAndJoinControls(string text) {
        OfficeFontFaceCollection fonts = LoadFont(
            "Noto Emoji",
            "NotoEmoji-VariableFont_wght.ttf",
            OfficeSixLaborsFontProgramProvider.Instance);

        OfficeFontFallbackRun run = Assert.Single(fonts.PlanFallbackRuns(text, "Noto Emoji"));
        Assert.Equal(text, run.Text);
        Assert.Equal("Noto Emoji", run.FamilyName);
        Assert.True(Assert.Single(fonts.Faces).Program.HasGlyphs(text));
    }

    [Fact]
    public void CompleteFontProgramReceivesLogicalArabicInsteadOfPresentationForms() {
        byte[] fontData = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "NotoSansArabic-Regular.ttf"));
        const string text = "مرحبا بالعالم";
        string html = "<html lang='ar' dir='rtl'><style>@font-face{font-family:Arabic;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(fontData)
            + "')}p{font:32px Arabic}</style><p>"
            + text
            + "</p></html>";
        var renderOptions = new HtmlRenderOptions().UseSixLaborsFontPrograms();
        var pdfOptions = new HtmlPdfSaveOptions().UseSixLaborsFontPrograms();

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), renderOptions);
        OfficeFontFace renderedFace = Assert.Single(rendered.Fonts.Faces);
        Assert.True(renderedFace.Program.ProvidesComplexTextLayout);
        Assert.True(renderedFace.Program.HasGlyphs(text));
        IReadOnlyList<HtmlRenderText> visuals = rendered.Pages
            .SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>()
            .Where(candidate => candidate.Font.FamilyName.Contains("Arabic", StringComparison.Ordinal))
            .ToList();

        Assert.NotEmpty(visuals);
        Assert.Equal(text, rendered.Text.Trim());
        Assert.All(visuals, visual => Assert.DoesNotContain(
            visual.Text,
            character => character >= '\uFE70' && character <= '\uFEFF'));

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).ToBytes();
        Assert.Contains(text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void OutlinedFontProgramHonorsCancellationAndPointBudgets() {
        OfficeFontFace face = Assert.Single(LoadFont(
            "Open Sans WOFF2",
            "OpenSans-Regular.woff2",
            OfficeSixLaborsFontProgramProvider.Instance).Faces);
        IOfficeBoundedFontProgram bounded = Assert.IsAssignableFrom<IOfficeBoundedFontProgram>(face.Program);
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => bounded.GetTextContoursBounded(
            "OfficeIMO",
            0D,
            0D,
            24D,
            maximumPointCount: 10_000,
            cancellation.Token));
        Assert.Throws<InvalidOperationException>(() => bounded.GetTextContoursBounded(
            "OfficeIMO",
            0D,
            0D,
            24D,
            maximumPointCount: 1,
            System.Threading.CancellationToken.None));
    }

    [Fact]
    public void HtmlPdf_OutlinedTextFailsClosedAtCharacterAndPathCommandBudgets() {
        byte[] woff2 = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        string html = "<style>@font-face{font-family:Bounded;src:url('data:font/woff2;base64,"
            + Convert.ToBase64String(woff2)
            + "')}p{font:28px Bounded}</style><p>OfficeIMO outlined text</p>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        var characterLimited = new HtmlPdfSaveOptions {
            MaxOutlinedTextCharactersPerRun = 4
        }.UseSixLaborsFontPrograms();
        var commandLimited = new HtmlPdfSaveOptions {
            MaxOutlinedTextPathCommands = 1
        }.UseSixLaborsFontPrograms();

        InvalidOperationException characterException = Assert.Throws<InvalidOperationException>(() =>
            source.ToPdfDocumentResult(characterLimited));
        Assert.Contains("character budget", characterException.Message, StringComparison.Ordinal);
        InvalidOperationException commandException = Assert.Throws<InvalidOperationException>(() =>
            source.ToPdfDocumentResult(commandLimited));
        Assert.Contains("point budget", commandException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_ReorderedOutlinedRunsEmitOneLogicalActualTextOwner() {
        byte[] woff2 = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        string html = "<html lang='en'><style>@font-face{font-family:Outlined;src:url('data:font/woff2;base64,"
            + Convert.ToBase64String(woff2)
            + "')}p{font:28px Outlined}</style><p><span>\u202E</span><b>abc</b><i>def</i><span>\u202C</span></p></html>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions().UseSixLaborsFontPrograms();

        byte[] pdf = source.ToPdfDocumentResult(options).ToBytes();
        string extracted = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("abcdef", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("cba", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain("fed", extracted, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(extracted, "abcdef"));
    }

    [Theory]
    [InlineData("Open Sans WOFF2", "OpenSans-Regular.woff2", "font/woff2", "OfficeIMO ffi 123")]
    [InlineData("Adobe Variable CFF2", "AdobeVFPrototype-Subset.otf", "font/otf", "$$$")]
    [InlineData("Roboto Flex", "RobotoFlex.ttf", "font/ttf", "Variable OfficeIMO")]
    public void HtmlPdf_NonStaticProgramsUseAccessibleDeterministicVectorOutlines(
        string family,
        string fileName,
        string mediaType,
        string text) {
        byte[] fontData = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", fileName));
        string html = "<html lang='en'><style>@font-face{font-family:'" + family + "';src:url('data:" + mediaType + ";base64,"
            + Convert.ToBase64String(fontData)
            + "')}p{font-family:'" + family + "';font-size:28px;line-height:1.25}</style><p>"
            + text
            + "</p></html>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        var options = new HtmlPdfSaveOptions()
            .UseSixLaborsFontPrograms();

        PdfCore.PdfDocumentConversionResult first = source.ToPdfDocumentResult(options);
        byte[] firstPdf = first.ToBytes();
        byte[] secondPdf = source.ToPdfDocumentResult(options).ToBytes();

        Assert.Equal(Convert.ToHexString(SHA256.HashData(firstPdf)), Convert.ToHexString(SHA256.HashData(secondPdf)));
        PdfCore.PdfReadDocument readDocument = PdfCore.PdfReadDocument.Open(firstPdf);
        string extractedText = readDocument.ExtractText();
        Assert.True(
            extractedText.Contains(text, StringComparison.Ordinal),
            extractedText + " | " + string.Join(" | ", readDocument.Pages[0].GetTextSpans().Select(span => span.Text + "@" + span.X + "," + span.Y)));
        Assert.Contains(first.Report.Warnings, warning =>
            warning.Code == HtmlPdfDiagnosticCodes.FontProgramOutlined
            && warning.Severity == PdfCore.PdfConversionWarningSeverity.Information
            && warning.Details.TryGetValue("Representation", out string? representation)
            && representation == "vector-outlines-plus-actual-text");
        Assert.False(first.HasLoss);
        HtmlPdfAccessibilityValidationResult validation = HtmlPdfAccessibilityValidator.Validate(source, firstPdf);
        Assert.True(validation.IsValid, string.Join(" | ", validation.Issues.Select(issue => issue.Code + ": " + issue.Message)));
        OfficeDrawing drawing = PdfCore.PdfDocument.Open(firstPdf).Read.Drawing(1);
        Assert.NotEmpty(drawing.Shapes);
    }

    [Fact]
    public void HtmlPdf_OutlinedWebFontRetainsHyperlinkInteraction() {
        byte[] woff2 = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        const string text = "Accessible linked outlines";
        string html = "<html lang='en'><style>@font-face{font-family:Linked;src:url('data:font/woff2;base64,"
            + Convert.ToBase64String(woff2)
            + "')}a{font:24px Linked}</style><a href='https://example.com/path'>"
            + text
            + "</a></html>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlPdfSaveOptions options = new HtmlPdfSaveOptions().UseSixLaborsFontPrograms();

        byte[] pdf = source.ToPdfDocumentResult(options).ToBytes();
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);

        Assert.Contains(text, read.ExtractText(), StringComparison.Ordinal);
        PdfCore.PdfLinkAnnotation link = Assert.Single(read.Pages[0].GetLinkAnnotations());
        Assert.Equal("https://example.com/path", link.Uri);
        HtmlPdfAccessibilityValidationResult validation = HtmlPdfAccessibilityValidator.Validate(source, pdf);
        Assert.True(validation.IsValid, string.Join(" | ", validation.Issues.Select(issue => issue.Code + ": " + issue.Message)));
    }

    [Fact]
    public void FallbackPackFingerprintIncludesVariableFontInstanceAndSurvivesSnapshots() {
        byte[] latin = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        byte[] variable = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"));
        OfficeFontFaceCollection lightFonts = new OfficeFontFaceCollection()
            .UseSixLaborsFontPrograms(new OfficeSixLaborsFontProgramProvider(_ =>
                new Dictionary<string, float> { ["wght"] = 200F }))
            .Add("Portable Latin", latin)
            .Add("Portable Variable", variable)
            .AddFallbackFamily("Portable Variable");
        OfficeFontFaceCollection blackFonts = new OfficeFontFaceCollection()
            .UseSixLaborsFontPrograms(new OfficeSixLaborsFontProgramProvider(_ =>
                new Dictionary<string, float> { ["wght"] = 900F }))
            .Add("Portable Latin", latin)
            .Add("Portable Variable", variable)
            .AddFallbackFamily("Portable Variable");

        var light = new OfficeFontFallbackPack("portable", "Portable Latin, Portable Variable", lightFonts);
        var identical = new OfficeFontFallbackPack("portable", "Portable Latin, Portable Variable", lightFonts);
        var black = new OfficeFontFallbackPack("portable", "Portable Latin, Portable Variable", blackFonts);

        Assert.Equal(light.Fingerprint, identical.Fingerprint);
        Assert.NotEqual(light.Fingerprint, black.Fingerprint);
        Assert.Equal(64, light.Fingerprint.Length);
        OfficeFontFaceCollection snapshot = light.Fonts;
        snapshot.Add("Unrelated", latin);
        Assert.Equal(2, light.Fonts.Faces.Count);
        Assert.Equal(new[] { "Portable Variable" }, light.Fonts.FallbackFamilies);
        Assert.Equal(light.Fingerprint, new OfficeFontFallbackPack(
            light.Id,
            light.DefaultFamilyNames,
            light.CreateRenderingProfile().Fonts).Fingerprint);

        var options = new HtmlPdfSaveOptions().UseFontFallbackPack(light);
        Assert.Equal(light.DefaultFamilyNames, options.DefaultFontFamily);
        Assert.Equal(light.Fingerprint, new OfficeFontFallbackPack(
            light.Id,
            options.DefaultFontFamily,
            options.Fonts!).Fingerprint);
    }

    private static OfficeFontFaceCollection LoadFont(
        string family,
        string fileName,
        OfficeSixLaborsFontProgramProvider provider) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", fileName));
        return new OfficeFontFaceCollection()
            .UseSixLaborsFontPrograms(provider)
            .Add(family, data);
    }

    private static void AssertInk(OfficeFontFaceCollection fonts, string family, string text) {
        byte[] pixels = Render(fonts, family, text).GetPixels();
        Assert.Contains(pixels, value => value < 250);
    }

    private static byte[] RenderHash(OfficeFontFaceCollection fonts, string family, string text) {
        using SHA256 sha256 = SHA256.Create();
        return sha256.ComputeHash(Render(fonts, family, text).GetPixels());
    }

    private static OfficeRasterImage Render(OfficeFontFaceCollection fonts, string family, string text) {
        var image = new OfficeRasterImage(480, 100, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);
        canvas.DrawText(text, 0D, 0D, image.Width, image.Height, OfficeColor.Black, 48D, fontFamily: family);
        return image;
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(token, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += token.Length;
        }
        return count;
    }
}
