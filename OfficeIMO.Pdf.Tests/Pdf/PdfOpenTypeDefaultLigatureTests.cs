using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOpenTypeDefaultLigatureTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LookupStopsAfterFirstMatchingSubtable(bool nested, bool extension) {
        var request = new OfficeTextShapingRequest("A", "Test",
            ManagedTextShapingTestAssets.CreateFontWithLookupScanningScenario(nested, extension, false), false, 1000,
            direction: OfficeTextDirection.LeftToRight, language: "en", featureSettings: OfficeTextFeatureSettings.Default.With("calt", 1));
        var result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
        Assert.Equal(new[] { 3, 4 }, result.Glyphs.Select(g => g.GlyphId));
        Assert.Equal("A", string.Concat(result.Glyphs.Select(g => g.UnicodeText)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContextualScanSkipsInsertedOutput(bool extension) {
        var request = new OfficeTextShapingRequest("A", "Test",
            ManagedTextShapingTestAssets.CreateFontWithLookupScanningScenario(true, extension, true), false, 1000,
            direction: OfficeTextDirection.LeftToRight, language: "en", featureSettings: OfficeTextFeatureSettings.Default.With("calt", 1));
        var result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
        Assert.Equal(new[] { 3, 2 }, result.Glyphs.Select(g => g.GlyphId));
        Assert.Equal("A", string.Concat(result.Glyphs.Select(g => g.UnicodeText)));
    }
    [Theory]
    [InlineData("ccmp")]
    [InlineData("rlig")]
    public void MandatoryLanguageFeatureCannotBeDisabled(string tag) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithRequiredLigature(tag), "Test");
        var run = font.ShapeText("fi", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
            featureSettings: OfficeTextFeatureSettings.Default.With(tag, 0)));
        Assert.Equal((ushort)3, Assert.Single(run.Glyphs).GlyphId);
        Assert.Equal("fi", run.Glyphs[0].UnicodeText);
    }

    [Fact]
    public void OversizedLookupListDeclinesShapingAndRetainsDiagnostic() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithOversizedLookupList();
        var report = new PdfConversionReport();
        byte[] pdf = PdfDocument.Create(new PdfOptions().ReportDiagnosticsTo(report)
            .EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test")).Paragraph(p => p.Text("fi")).ToBytes();
        Assert.Contains("fi", PdfReadDocument.Open(pdf).ExtractText());
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContextualExpansionRequiresStableInputSlots(bool expansionLast) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithContextualExpansion(expansionLast);
        var request = new OfficeTextShapingRequest("AB", "Test", data, false, 1000,
            direction: OfficeTextDirection.LeftToRight, language: "en",
            featureSettings: OfficeTextFeatureSettings.Default.With("calt", 1));
        var result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);
        if (!expansionLast) { Assert.Null(result); return; }
        Assert.NotNull(result);
        Assert.Equal(3, result!.Glyphs.Count);
        Assert.Equal("AB", string.Concat(result.Glyphs.Select(g => g.UnicodeText)));
        Assert.Equal(3, result.Glyphs.Last().GlyphId);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultLigaturesUseFontGlyphsAndPreserveSourceText(bool cff) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        var shaping = PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures);
        PdfGlyphRun run = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Test").ShapeText("office", shaping)
            : PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText("office", shaping);
        Assert.True(run.Glyphs.Count < 6);
        Assert.Equal("office", string.Concat(run.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Null(run.ActualText);
        var report = new PdfConversionReport();
        var options = new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test");
        byte[] pdf = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("office affinity fine flow")).ToBytes();
        Assert.Contains("office affinity fine flow", PdfReadDocument.Open(pdf).ExtractText());
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
        if (!cff) {
            var font = PdfTrueTypeFontProgram.Parse(data, "Test");
            Assert.Equal(run.TotalAdvanceWidth1000 * 12D / 1000, font.MeasureTextWidth("office", 12, PdfTextShapingMode.OpenTypeLigatures), 5);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitLigatureDisableAndScalarModeKeepSeparateGlyphs(bool cff) {
        byte[] data = File.ReadAllBytes((cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont())!);
        var disabled = PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
            featureSettings: OfficeTextFeatureSettings.Default.With("liga", 0));
        var scalar = PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar);
        PdfGlyphRun Shape(PdfTextShapingOptions options) => cff ? PdfOpenTypeCffFontProgram.Parse(data, "Test").ShapeText("office", options)
            : PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText("office", options);
        Assert.Equal(6, Shape(disabled).Glyphs.Count);
        Assert.Equal(6, Shape(scalar).Glyphs.Count);
    }
    [Fact]
    public void DefaultLigaturesDoNotRequirePresentationCodePoints() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn");
        var font = PdfTrueTypeFontProgram.Parse(data, "Test");
        var run = font.ShapeText("fi", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Single(run.Glyphs);
        Assert.Equal("fi", run.Glyphs[0].UnicodeText);
        Assert.Equal(2, font.ShapeText("fi", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar)).Glyphs.Count);
    }

    [Fact]
    public void OtherScriptLigaturesDoNotWarnForLatinText() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "arab");
        var report = new PdfConversionReport();
        var options = new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test");
        byte[] pdf = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("fi")).ToBytes();
        Assert.Contains("fi", PdfReadDocument.Open(pdf).ExtractText());
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Fact]
    public void UnsupportedLookupFlagsPreserveScalarTextAndDiagnostic() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn", lookupFlags: 8);
        var run = PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText("fi",
            PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(2, run.Glyphs.Count);
        var report = new PdfConversionReport();
        var options = new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test");
        byte[] pdf = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("fi")).ToBytes();
        Assert.Contains("fi", PdfReadDocument.Open(pdf).ExtractText());
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Fact]
    public void AutomaticLigaturesRetainUnimplementedMarkPositioningWarnings() {
        byte[] data = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        var report = new PdfConversionReport();
        var options = new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test");
        byte[] pdf = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("office e\u0301")).ToBytes();
        Assert.Contains("office", PdfReadDocument.Open(pdf).ExtractText());
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-mark-positioning");
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-mark-positioning-or-joiner-shaping");
    }

    [Fact]
    public void ImplicitComplexShapingRetainsLogicalActualTextInDefaultMode() {
        const string text = "\u202Efi\u202C";
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('f', 'i');
        var run = PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText(text,
            PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
                direction: OfficeTextDirection.RightToLeft));
        Assert.Equal(text, run.ActualText);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void MixedPresentationAndSubstitutionGlyphsPreserveEachSource(bool cff, bool reverse) {
        byte[] data = File.ReadAllBytes((cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont())!);
        string first = reverse ? "fi" : "\uFB01";
        string second = reverse ? "\uFB01" : "fi";
        byte[] pdf = PdfDocument.Create(new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Paragraph(paragraph => paragraph.Text(first))
            .Paragraph(paragraph => paragraph.Text(second)).ToBytes();
        string extracted = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("\uFB01", extracted);
        Assert.Contains("fi", extracted);
    }

    [Fact]
    public void SuccessfulProgramDoesNotSuppressAnotherProgramsFallbackWarning() {
        byte[] supported = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn");
        byte[] unsupported = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn", lookupFlags: 8);
        var report = new PdfConversionReport();
        var options = new PdfOptions().ReportDiagnosticsTo(report)
            .EmbedStandardFont(PdfStandardFont.Helvetica, supported, "SharedName")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, unsupported, "SharedName");
        byte[] pdf = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("fi").Bold("fi")).ToBytes();
        Assert.Contains("unsupported-font-ligature-substitution", report.Warnings.Select(warning => warning.Code));
    }

    [Theory]
    [InlineData(0x03B1, 0x03B2)]
    [InlineData(0x0430, 0x0431)]
    public void DefaultLatinLookupsDoNotSubstituteOtherScripts(int first, int second) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature(first, second, scriptTag: "latn");
        var run = PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText(char.ConvertFromUtf32(first) + char.ConvertFromUtf32(second),
            PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(2, run.Glyphs.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PresentationCharacterDoesNotDisableNeighboringLatinLigatures(bool cff) {
        byte[] data = File.ReadAllBytes((cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont())!);
        var settings = PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures);
        var run = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Test").ShapeText("\uFB01 office", settings)
            : PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText("\uFB01 office", settings);
        Assert.True(run.Glyphs.Count < 8);
        Assert.Equal("\uFB01 office", string.Concat(run.Glyphs.Select(glyph => glyph.UnicodeText)));
    }

    [Fact]
    public void ExplicitRtlAutomaticRunKeepsOriginalLogicalText() {
        const string text = "()[]";
        byte[] data = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!);
        var run = PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText(text,
            PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures, direction: OfficeTextDirection.RightToLeft));
        Assert.Equal(text, run.ActualText);
    }

    [Fact]
    public void EnabledUnsupportedDiscretionaryLigaturesHaveFallbackDiagnostic() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', featureTag: "dlig", scriptTag: "latn", lookupFlags: 8);
        var font = PdfTrueTypeFontProgram.Parse(data, "Test");
        Assert.Contains(PdfTextDiagnostics.AnalyzeAdvancedTextLayout("fi", font, featureSettings: OfficeTextFeatureSettings.Default.With("dlig", 1)),
            diagnostic => diagnostic.Code == "unsupported-font-ligature-substitution");
        Assert.DoesNotContain(PdfTextDiagnostics.AnalyzeAdvancedTextLayout("fi", font), diagnostic => diagnostic.Code == "unsupported-font-ligature-substitution");
    }

    [Fact]
    public void LogicalWordScopeIncludesMultipleSubstitutionContinuations() {
        var glyphs = new[] {
            new PdfGlyphInfo(1, "A", 0, 600, 600, 0, 0, 0),
            new PdfGlyphInfo(2, "", 0, 600, 600, 0, 0, 0),
            new PdfGlyphInfo(3, "", 0, 600, 600, 0, 0, 0)
        };
        var output = new System.Text.StringBuilder();
        new ContentStreamBuilder(output).BeginText().TextMatrix(40, 400)
            .ShowText(new PdfGlyphRun(glyphs, System.Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true).ToTextShowCommand(), 12).EndText();
        string content = output.ToString();
        Assert.Contains("<000100020003> Tj", content);
        Assert.Equal(1, content.Split(new[] { "/ActualText" }, System.StringSplitOptions.None).Length - 1);
    }

    [Theory]
    [InlineData("A", true)]
    [InlineData("AB", false)]
    [InlineData("A ", false)]
    public void ComposedSubstitutionsKeepLogicalSourceAndContinuationOwnership(string text, bool continuationOnly) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithComposedMultipleLigature(text.Length > 1 ? text[1] : 'B', continuationOnly);
        var run = PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(text, string.Concat(run.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal(6, run.Glyphs.Last().GlyphId);
        Assert.Equal(0, run.Glyphs.Last().LogicalClusterStart);
        Assert.Equal(text.Length > 1 ? 1 : 0, run.Glyphs.Last().TextIndex);
        var output = new System.Text.StringBuilder();
        new ContentStreamBuilder(output).BeginText().TextMatrix(40, 400).ShowText(run.ToTextShowCommand(), 12).EndText();
        Assert.Equal(1, output.ToString().Split(new[] { "/ActualText" }, System.StringSplitOptions.None).Length - 1);
    }

    [Theory]
    [InlineData("A", true)]
    [InlineData("AB", false)]
    public void RedactionRemovesComposedContinuationPaint(string text, bool continuationOnly) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithComposedMultipleLigature('B', continuationOnly);
        var run = PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Paragraph(paragraph => paragraph.Text(text)).ToBytes();
        var span = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().First();
        string painted = "<" + string.Concat(run.Glyphs.Select(glyph => glyph.GlyphId.ToString("X4"))) + "> Tj";
        Assert.Contains(painted, PdfEncoding.Latin1GetString(pdf));
        byte[] redacted = PdfRedactionApplier.Apply(pdf, new[] { new PdfRedactionArea(1, span.X - 1, span.Y - span.FontSize * 1.5, span.FontSize, span.FontSize * 2, "cluster") });
        Assert.DoesNotContain(painted, PdfEncoding.Latin1GetString(redacted));
    }

    [Fact]
    public void RedactionRemovesEveryGlyphInDefaultMultipleSubstitutionCluster() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithMultipleSubstitution('A', scriptTag: "latn", featureTag: "liga");
        var font = PdfTrueTypeFontProgram.Parse(data, "Test");
        var run = font.ShapeText("A", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(2, run.Glyphs.Count);
        Assert.Equal("", run.Glyphs[1].UnicodeText);
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Paragraph(paragraph => paragraph.Text("A")).ToBytes();
        var span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans(), item => item.Text == "A");
        Assert.Equal(run.TotalAdvanceWidth1000 * span.FontSize / 1000D, span.Advance, 3);
        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { new PdfRedactionArea(1, span.X - 1, span.Y - span.FontSize * 1.5, span.Advance + 2, span.FontSize * 2, "cluster") });
        Assert.DoesNotContain("A", PdfTextExtractor.ExtractAllText(redacted));
        Assert.DoesNotContain("<00030004> Tj", PdfEncoding.Latin1GetString(redacted));
    }

}
