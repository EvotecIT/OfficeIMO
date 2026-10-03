using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfShortTextReuseTests {
    [Theory]
    [InlineData(false, "office affinity fine flow")]
    [InlineData(true, "office affinity fine flow")]
    [InlineData(false, "123.45 ([)]")]
    [InlineData(true, "123.45 ([)]")]
    [InlineData(false, "R09999 Value 42.00")]
    [InlineData(true, "R09999 Value 42.00")]
    public void DefaultPdfRunPreservesTheManagedGlyphAndLogicalClusterContract(bool cff, string text) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        Verify(File.ReadAllBytes(path!), cff, text);
    }

    [Fact]
    public void DefaultPdfRunPreservesLigatureSourceClusters() =>
        Verify(ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn"), false, "fi fi");

    [Fact]
    public void DefaultPdfRunPreservesMultipleSubstitutionContinuationGlyphs() =>
        Verify(ManagedTextShapingTestAssets.CreateFontWithMultipleSubstitution('A', scriptTag: "latn"), false, "A A");

    [Fact]
    public void DefaultPdfNumericRunDoesNotSelectUnsupportedLatinLookups() =>
        Verify(ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(true), false, "11,1");

    private static void Verify(byte[] data, bool cff, string text) {
        int notifications = 0;
        var options = PdfTextShapingOptions.ForRendering("Projection", PdfTextShapingMode.OpenTypeLigatures,
            providerShapedTextRecorder: (source, fontName, isCff, automatic) => {
                Assert.Equal(text, source);
                Assert.Equal("Projection", fontName);
                Assert.Equal(cff, isCff);
                Assert.True(automatic);
                notifications++;
            });
        PdfTrueTypeFontProgram? trueType = cff ? null : PdfTrueTypeFontProgram.Parse(data, "Projection");
        PdfOpenTypeCffFontProgram? cffFont = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Projection") : null;
        int unitsPerEm = cffFont?.UnitsPerEm ?? trueType!.UnitsPerEm;
        Func<int, int> getWidth = cff ? cffFont!.GetGlyphWidth1000 : trueType!.GetGlyphWidth1000;
        PdfGlyphRun first = cff ? cffFont!.ShapeText(text, options) : trueType!.ShapeText(text, options);
        var mappings = cff ? cffFont!.GetGlyphToUnicodeMappings() : trueType!.GetGlyphToUnicodeMappings();
        if (cff) cffFont!.ResetGlyphUsage(); else trueType!.ResetGlyphUsage();
        PdfGlyphRun actual = cff ? cffFont!.ShapeText(text, options) : trueType!.ShapeText(text, options);
        Assert.Equal(first.Glyphs, actual.Glyphs);
        Assert.Equal(mappings, cff ? cffFont!.GetGlyphToUnicodeMappings() : trueType!.GetGlyphToUnicodeMappings());
        var request = new OfficeTextShapingRequest(text, "Projection", data, cff, unitsPerEm,
            OfficeTextDirection.Auto, null, default, fontCollectionIndex: null, variationCoordinates: null,
            cloneFontData: true, applyDefaultLatinLigatures: true);
        OfficeTextShapingResult? expected = OfficeManagedTextShapingProvider.Instance.ShapeText(request);
        Assert.NotNull(expected);
        Assert.Equal(expected!.Direction, actual.Direction);
        Assert.Equal(expected.Glyphs.Count, actual.Glyphs.Count);
        for (int index = 0; index < actual.Glyphs.Count; index++) {
            OfficeShapedGlyph source = expected.Glyphs[index];
            PdfGlyphInfo projected = actual.Glyphs[index];
            Assert.Equal(source.GlyphId, projected.GlyphId);
            Assert.Equal(source.UnicodeText, projected.UnicodeText);
            Assert.Equal(source.TextIndex, projected.TextIndex);
            Assert.Equal(expected.GetLogicalClusterStart(index), projected.LogicalClusterStart);
            Assert.Equal(getWidth(source.GlyphId), projected.NominalWidth1000);
            Assert.Equal(projected.NominalWidth1000, projected.AdvanceWidth1000);
            Assert.False(projected.HasPositioning);
        }
        Assert.Equal(expected.Direction == OfficeTextDirection.LeftToRight ? null : text, actual.ActualText);
        Assert.Equal(expected.Direction == OfficeTextDirection.LeftToRight, actual.PreserveGlyphUnicode);
        Assert.Equal(expected.Glyphs.Select(glyph => glyph.GlyphId).Distinct().OrderBy(id => id),
            cff ? cffFont!.GetUsedGlyphIds() : trueType!.GetUsedGlyphIds());
        Assert.Equal(2, notifications);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedMeasurementRestoresSubsetUsageAfterResetAndKeepsFontSizeScaling(bool cff) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        PdfTrueTypeFontProgram? tt = cff ? null : PdfTrueTypeFontProgram.Parse(data);
        PdfOpenTypeCffFontProgram? ot = cff ? PdfOpenTypeCffFontProgram.Parse(data) : null;
        string text = "office affinity 123.45";
        double width = cff ? ot!.MeasureTextWidth(text, 10, PdfTextShapingMode.OpenTypeLigatures)
            : tt!.MeasureTextWidth(text, 10, PdfTextShapingMode.OpenTypeLigatures);
        var mappings = cff ? ot!.GetGlyphToUnicodeMappings() : tt!.GetGlyphToUnicodeMappings();
        var glyphs = cff ? ot!.GetUsedGlyphIds() : tt!.GetUsedGlyphIds();
        if (cff) ot!.ResetGlyphUsage(); else tt!.ResetGlyphUsage();
        double repeated = cff ? ot!.MeasureTextWidth(text, 20, PdfTextShapingMode.OpenTypeLigatures)
            : tt!.MeasureTextWidth(text, 20, PdfTextShapingMode.OpenTypeLigatures);
        Assert.Equal(width * 2, repeated);
        Assert.Equal(glyphs, cff ? ot!.GetUsedGlyphIds() : tt!.GetUsedGlyphIds());
        Assert.Equal(mappings, cff ? ot!.GetGlyphToUnicodeMappings() : tt!.GetGlyphToUnicodeMappings());
        Assert.Empty(cff ? ot!.ForkForDocument().GetUsedGlyphIds() : tt!.ForkForDocument().GetUsedGlyphIds());
    }

    [Fact]
    public void RepeatedMultipleSubstitutionMeasurementKeepsContinuationMappings() {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithMultipleSubstitution('A', scriptTag: "latn"));
        double first = font.MeasureTextWidth("A A", 10, PdfTextShapingMode.OpenTypeLigatures);
        var mappings = font.GetGlyphToUnicodeMappings();
        var glyphs = font.GetUsedGlyphIds();
        font.ResetGlyphUsage();
        Assert.Equal(first, font.MeasureTextWidth("A A", 10, PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(glyphs, font.GetUsedGlyphIds());
        Assert.Equal(mappings, font.GetGlyphToUnicodeMappings());
    }

    [Fact]
    public void CachedDefaultTextDoesNotSuppressProvidersOrExplicitFeatures() {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn"));
        var defaults = PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures);
        Assert.Single(font.ShapeText("fi", defaults).Glyphs);
        var disabled = PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures,
            featureSettings: OfficeTextFeatureSettings.Default.With("liga", 0));
        Assert.Equal(2, font.ShapeText("fi", disabled).Glyphs.Count);
        var provider = new DecliningProvider();
        var custom = PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures, provider);
        Assert.Single(font.ShapeText("fi", custom).Glyphs);
        Assert.Single(font.ShapeText("fi", custom).Glyphs);
        Assert.Equal(2, provider.Calls);
    }

    private sealed class DecliningProvider : IOfficeTextShapingProvider {
        internal int Calls { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Calls++; return null;
        }
    }

    [Fact]
    public void ReusedScalarFallbackRetainsUsageWithoutClaimingProviderCoverage() {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(true));
        int notifications = 0;
        var options = PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures,
            providerShapedTextRecorder: (_, _, _, _) => notifications++);
        PdfGlyphRun first = font.ShapeText("A,B", options);
        Assert.Null(first.SourceShapingResult);
        var mappings = font.GetGlyphToUnicodeMappings();
        var glyphs = font.GetUsedGlyphIds();
        font.ResetGlyphUsage();
        PdfGlyphRun repeated = font.ShapeText("A,B", options);
        Assert.Equal(first.Glyphs, repeated.Glyphs);
        Assert.Equal(mappings, font.GetGlyphToUnicodeMappings());
        Assert.Equal(glyphs, font.GetUsedGlyphIds());
        Assert.Equal(0, notifications);
        double width = font.MeasureTextWidth("A,B", 10, PdfTextShapingMode.OpenTypeLigatures);
        font.ResetGlyphUsage();
        Assert.Equal(width, font.MeasureTextWidth("A,B", 10, PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(mappings, font.GetGlyphToUnicodeMappings());
        Assert.Equal(glyphs, font.GetUsedGlyphIds());
    }

    [Fact]
    public void DiagnosticMissingGlyphsCannotSuppressRenderingOrMeasurementFailures() {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn"));
        const string text = "fi~";
        Assert.False(font.TryGetGlyphId('~', out _));
        PdfGlyphRun diagnostic = font.ShapeText(text, PdfTextShapingOptions.ForDiagnostics("test", font.FontName, PdfTextShapingMode.OpenTypeLigatures));
        Assert.True(diagnostic.HasMissingGlyphs);
        for (int attempt = 0; attempt < 2; attempt++) {
            Assert.Throws<ArgumentException>(() => font.ShapeText(text, PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures)));
            Assert.Throws<ArgumentException>(() => font.MeasureTextWidth(text, 10, PdfTextShapingMode.OpenTypeLigatures));
        }
    }
}
