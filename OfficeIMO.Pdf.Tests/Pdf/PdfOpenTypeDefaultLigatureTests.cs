using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOpenTypeDefaultLigatureTests {
    [Fact]
    public void DeclinedComplexScriptRetainsFallbackDiagnostic() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(0x20, 0x0915, 0x094D, 0x0937);
        var report = new PdfConversionReport();
        PdfDocument.Create(new PdfOptions().ReportDiagnosticsTo(report).UseManagedTextShaping()
            .EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Paragraph(p => p.Text("\u0915\u094D\u0937")).ToBytes();
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmbeddedCompositeJustificationPositionsWordsAtMeasuredGap(bool cff) {
        byte[] data = File.ReadAllBytes((cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont())!);
        var options = new PdfOptions { PageWidth = 250, MarginLeft = 20, MarginRight = 20 }
            .EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test");
        byte[] pdf = PdfDocument.Create(options).Paragraph(p => p.Text("office affinity fine flow efficient office affinity fine flow efficient"), PdfAlign.Justify).ToBytes();
        using var reader = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var firstLine = reader.GetPage(1).GetWords().GroupBy(word => System.Math.Round(word.BoundingBox.Bottom, 1))
            .OrderByDescending(group => group.Key).First().OrderBy(word => word.BoundingBox.Left).ToArray();
        Assert.True(firstLine.Length > 1);
        Assert.InRange(firstLine.Last().BoundingBox.Right, 225, 232);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DecliningConfiguredProviderRetainsAutomaticLatinLigatures(bool cff) {
        byte[] data = File.ReadAllBytes((cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont())!);
        var options = PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
            shapingProvider: OfficeManagedTextShapingProvider.Instance);
        var run = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Test").ShapeText("office", options)
            : PdfTrueTypeFontProgram.Parse(data, "Test").ShapeText("office", options);
        Assert.True(run.Glyphs.Count < 6);
        Assert.Equal("office", string.Concat(run.Glyphs.Select(g => g.UnicodeText)));
    }

    [Fact]
    public void AutomaticLatinHonorsExplicitRtlNeutralMirroring() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('(', ')', '[', ']');
        var font = PdfTrueTypeFontProgram.Parse(data, "Test");
        var run = font.ShapeText("([)]", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
            direction: OfficeTextDirection.RightToLeft));
        // Reverse visual order with mirrored parentheses/brackets: [(])
        Assert.Equal(new[] { 3, 1, 4, 2 }, run.Glyphs.Select(g => g.GlyphId));
        Assert.Equal("([)]", run.ActualText);
    }

    [Fact]
    public void TwoByteLogicalGlyphCodesIgnorePdfWordSpacing() {
        var glyphs = new[] { new PdfGlyphInfo(32, " ", 0, 600, 600, 0, 0, 0), new PdfGlyphInfo(33, "A", 1, 600, 600, 0, 0, 0) };
        var output = new System.Text.StringBuilder();
        new ContentStreamBuilder(output).BeginText().TextMatrix(40, 400).WordSpacing(20)
            .ShowText(new PdfGlyphRun(glyphs, System.Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true).ToTextShowCommand(), 10).EndText();
        Assert.Contains("1 0 0 1 46 400 Tm", output.ToString());
        Assert.DoesNotContain("1 0 0 1 66 400 Tm", output.ToString());
    }

    [Fact]
    public void RedactingOneOrdinaryCharacterPreservesOtherCharactersInWord() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'a', 'h', 'l', 'p');
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Canvas(c => c.Text("Alpha", 40, 40, 200, 30, fontSize: 12)).ToBytes();
        var target = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().Single(s => s.Text == "p");
        byte[] redacted = PdfRedactionApplier.Apply(pdf, new[] { new PdfRedactionArea(1,
            target.X + 0.1, target.Y - target.FontSize, target.Advance - 0.2, target.FontSize * 1.5, "character") });
        Assert.Equal("Alha", string.Concat(PdfReadDocument.Open(redacted).Pages[0].GetTextSpans().Select(span => span.Text)));
        Assert.Equal("Al ha", PdfTextExtractor.ExtractAllText(redacted).Trim());
    }
    [Theory]
    [InlineData(0x0259)]
    [InlineData(0x1D00)]
    [InlineData(0x02E0)]
    [InlineData(0x2071)]
    [InlineData(0x2160)]
    [InlineData(0xFF21)]
    [InlineData(0x1DF25)]
    public void LatinRepertoireUsesRequiredFontFeatures(int first) {
        string text = char.ConvertFromUtf32(first) + "A";
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithRequiredLigature("ccmp", first, 'A');
        var font = PdfTrueTypeFontProgram.Parse(data, "Test");
        var run = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(3, Assert.Single(run.Glyphs).GlyphId);
        Assert.Equal(text, run.Glyphs[0].UnicodeText);
        byte[] unsupported = ManagedTextShapingTestAssets.CreateFontWithRequiredLigature("ccmp", first, 'A', 8);
        var report = new PdfConversionReport();
        PdfDocument.Create(new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, unsupported, "Test"))
            .Paragraph(p => p.Text(text)).ToBytes();
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Theory]
    [InlineData("\u03B1\u00B5\u03B2", false)]
    [InlineData("\u00B5", false)]
    [InlineData("A\u00B5B", true)]
    public void CommonMicroSignDoesNotEstablishLatinScript(string text, bool substituted) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(false, true), "Test");
        var shaped = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(text, string.Concat(shaped.Glyphs.Select(g => g.UnicodeText)));
        Assert.Equal(text.Length + (substituted ? 1 : 0), shaped.Glyphs.Count);
    }

    [Theory]
    [InlineData("\u200E")]
    [InlineData("\u200F")]
    [InlineData("\u202A\u202C")]
    public void RemovedDirectionalControlsRemainLigatureBoundaries(string controls) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn"), "Test");
        var shaped = font.ShapeText("f" + controls + "i", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(new[] { 1, 2 }, shaped.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new[] { 0, controls.Length + 1 }, shaped.Glyphs.Select(g => g.TextIndex));
    }

    [Theory]
    [InlineData("\u200E,A")]
    [InlineData("A,\u200E")]
    [InlineData("A\u200E,B")]
    public void RemovedControlsPreserveCommonScriptContextAtRunEdges(string text) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(false), "Test");
        var shaped = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(2 + (text.Contains("B") ? 1 : 0), shaped.Glyphs.Count);
        Assert.Contains(shaped.Glyphs, g => g.GlyphId == 2 && g.UnicodeText == ",");
    }
    [Theory]
    [InlineData("\u03B1,\u03B2", false)]
    [InlineData("\u0410,\u0411", false)]
    [InlineData(",\u03B1", false)]
    [InlineData("\u03B1,", false)]
    [InlineData("\u03B1,A", false)]
    [InlineData("A,\u03B1", false)]
    [InlineData(",", false)]
    [InlineData("A,B", true)]
    [InlineData(",A", true)]
    [InlineData("A,", true)]
    public void CommonCharactersInheritAdjacentStrongScript(string text, bool substituted) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(false), "Test");
        var scalar = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar));
        var shaped = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(text, string.Concat(shaped.Glyphs.Select(g => g.UnicodeText)));
        if (substituted) Assert.Equal(scalar.Glyphs.Count + 1, shaped.Glyphs.Count);
        else Assert.Equal(scalar.Glyphs.Select(g => g.GlyphId), shaped.Glyphs.Select(g => g.GlyphId));
    }

    [Theory]
    [InlineData("\u03B1,\u03B2", false)]
    [InlineData("\u0410,\u0411", false)]
    [InlineData(",", false)]
    [InlineData("A,B", true)]
    [InlineData("\u03B1,A", true)]
    public void UnsupportedLatinFeaturesWarnOnlyForEligibleInput(string text, bool warningExpected) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(true);
        var report = new PdfConversionReport();
        PdfDocument.Create(new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Paragraph(p => p.Text(text)).ToBytes();
        Assert.Equal(warningExpected, report.Warnings.Any(w => w.Code == "unsupported-font-ligature-substitution"));
    }
    [Theory]
    [InlineData("\u03B1\u0301", false)]
    [InlineData("\u0410\u0301", false)]
    [InlineData("\u0301", false)]
    [InlineData("\u03B1,\u0301", false)]
    [InlineData("1\u0301", false)]
    [InlineData("\u00AD\u0301", false)]
    [InlineData("A\u0301", true)]
    public void CombiningMarkInheritsItsBaseScript(string text, bool substituted) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithInheritedMarkSubstitution(), "Test");
        var scalar = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar));
        var shaped = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(text, string.Concat(shaped.Glyphs.Select(g => g.UnicodeText)));
        if (substituted) Assert.Equal(scalar.Glyphs.Count + 1, shaped.Glyphs.Count);
        else Assert.Equal(scalar.Glyphs.Select(g => g.GlyphId), shaped.Glyphs.Select(g => g.GlyphId));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LigatureSetUsesFontPreferenceOrder(bool longestFirst) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithOrderedLigatures(longestFirst), "Test");
        var shaped = font.ShapeText("fin", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(longestFirst ? new[] { 5 } : new[] { 4, 3 }, shaped.Glyphs.Select(g => g.GlyphId));
        Assert.Equal("fin", string.Concat(shaped.Glyphs.Select(g => g.UnicodeText)));
    }
    [Theory]
    [InlineData("\u0628\u062A")]
    [InlineData("\u05D0\u05D1")]
    public void AutomaticLatinModePreservesNonLatinScalarSequence(string text) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(0x0628, 0x062A, 0xFE91, 0xFE96, 0x05D0, 0x05D1);
        var font = PdfTrueTypeFontProgram.Parse(data, "Test");
        var scalar = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar));
        var automatic = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(scalar.Glyphs.Select(glyph => glyph.GlyphId), automatic.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(text, string.Concat(automatic.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal(scalar.TotalAdvanceWidth1000, automatic.TotalAdvanceWidth1000);
        var explicitManaged = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
            shapingProvider: OfficeManagedTextShapingProvider.Instance));
        Assert.NotEqual(automatic.Glyphs.Select(glyph => glyph.GlyphId).ToArray(), explicitManaged.Glyphs.Select(glyph => glyph.GlyphId).ToArray());
    }
    [Fact]
    public void RedactingMixedWhitespaceLigaturePreservesFollowingWord() {
        byte[] data = ManagedTextShapingTestAssets.CreateMixedWhitespaceLigatureFont();
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Canvas(canvas => canvas.Text("A B", 40, 40, 200, 30, fontSize: 12)).ToBytes();
        var spans = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans();
        var first = spans.First(span => span.Text.Contains("A"));
        Assert.Contains("B", PdfTextExtractor.ExtractAllText(pdf));
        byte[] redacted = PdfRedactionApplier.Apply(pdf, new[] { new PdfRedactionArea(1,
            first.X, first.Y - first.FontSize * 1.5, first.Advance - 0.1, first.FontSize * 2, "cluster") });
        string extracted = PdfTextExtractor.ExtractAllText(redacted);
        Assert.DoesNotContain("A", extracted);
        Assert.Contains("B", extracted);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContextScanningConsumesInputButNotLookahead(bool lookahead) {
        var request = new OfficeTextShapingRequest("AAA", "Test",
            ManagedTextShapingTestAssets.CreateFontWithContextInputOrLookahead(lookahead), false, 1000,
            direction: OfficeTextDirection.LeftToRight, language: "en", featureSettings: OfficeTextFeatureSettings.Default.With("calt", 1));
        var result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
        Assert.Equal(lookahead ? new[] { 2, 2, 1 } : new[] { 2, 1, 1 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal("AAA", string.Concat(result.Glyphs.Select(glyph => glyph.UnicodeText)));
    }
    [Theory]
    [InlineData(0x05D0)]
    [InlineData(0x0627)]
    [InlineData(0x0915)]
    public void MixedComplexTextStillShapesLatinSegment(int neighbor) {
        string text = "fi " + char.ConvertFromUtf32(neighbor);
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateMixedLatinLigatureFont(neighbor), "Test");
        var run = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(3, run.Glyphs.Count);
        Assert.Contains(run.Glyphs, glyph => glyph.GlyphId == 5 && glyph.UnicodeText == "fi");
        Assert.Equal(text, run.ActualText);
    }
    [Fact]
    public void ExplicitKerningWorksWithoutGsubInDefaultMode() {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -100), "Test");
        var scalar = font.ShapeText("AV", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar,
            featureSettings: OfficeTextFeatureSettings.Default.With("kern", 1)));
        var automatic = font.ShapeText("AV", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures,
            featureSettings: OfficeTextFeatureSettings.Default.With("kern", 1)));
        Assert.Equal(scalar.TotalAdvanceWidth1000, automatic.TotalAdvanceWidth1000);
        Assert.Contains(automatic.Glyphs, glyph => glyph.HasPositioning);
        Assert.Equal(font.ShapeText("AV", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.UnicodeScalar)).TotalAdvanceWidth1000 - 100,
            automatic.TotalAdvanceWidth1000);
    }

    [Theory]
    [InlineData(0x1AB0)]
    [InlineData(0x1DC0)]
    [InlineData(0x20D0)]
    [InlineData(0xFE20)]
    public void RequiredCompositionKeepsSupplementaryCombiningMarks(int mark) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithRequiredLigature("ccmp", 'A', mark), "Test");
        string text = "A" + char.ConvertFromUtf32(mark);
        var run = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(3, Assert.Single(run.Glyphs).GlyphId);
        Assert.Equal(text, run.Glyphs[0].UnicodeText);
    }

    [Theory]
    [InlineData("ccmp", 'A', 'B')]
    [InlineData("liga", 's', 't')]
    public void UnsupportedSelectedFeaturesWarnWithoutPresentationSequence(string tag, char first, char second) {
        byte[] data = tag == "ccmp" ? ManagedTextShapingTestAssets.CreateFontWithRequiredLigature(tag, first, second, 8)
            : ManagedTextShapingTestAssets.CreateFontWithLigature(first, second, scriptTag: "latn", lookupFlags: 8);
        var report = new PdfConversionReport();
        PdfDocument.Create(new PdfOptions().ReportDiagnosticsTo(report).EmbedStandardFont(PdfStandardFont.Helvetica, data, "Test"))
            .Paragraph(p => p.Text(new string(new[] { first, second }))).ToBytes();
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Theory]
    [InlineData("A ")]
    [InlineData("A B")]
    [InlineData(" A")]
    public void MixedWhitespaceClusterDoesNotOwnFollowingWord(string source) {
        var glyphs = new[] { new PdfGlyphInfo(1, source, 0, 600, 600, 0, 0, 0), new PdfGlyphInfo(2, "B", source.Length, 600, 600, 0, 0, 0) };
        var output = new System.Text.StringBuilder();
        new ContentStreamBuilder(output).BeginText().TextMatrix(40, 400)
            .ShowText(new PdfGlyphRun(glyphs, System.Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true).ToTextShowCommand(), 12).EndText();
        Assert.Equal(2, output.ToString().Split(new[] { "/ActualText" }, System.StringSplitOptions.None).Length - 1);
    }

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

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    public void RunLayoutDiagnosticsHonorEachRunsOptionalFeatures(int enabled, bool configuredRunFirst) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', featureTag: "dlig", scriptTag: "latn", lookupFlags: 8);
        var configured = new PdfTextRun("fi").WithFeatureSettings(OfficeTextFeatureSettings.Default.With("dlig", enabled));
        var ordinary = new PdfTextRun("fi");
        var runs = configuredRunFirst ? new[] { configured, ordinary } : new[] { ordinary, configured };

        var diagnostics = PdfTextDiagnostics.AnalyzeAdvancedTextLayoutRuns(runs, data, source: "Run preflight");
        var substitutions = diagnostics.Where(diagnostic => diagnostic.Code == "unsupported-font-ligature-substitution").ToList();
        if (enabled == 1) {
            Assert.Single(substitutions);
            Assert.Equal("Run preflight", substitutions[0].Source);
        } else {
            Assert.Empty(substitutions);
        }
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
