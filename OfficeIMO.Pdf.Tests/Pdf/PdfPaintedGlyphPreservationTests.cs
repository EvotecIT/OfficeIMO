using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPaintedGlyphPreservationTests {
    private const string ArabicVisual = "\uFEE1\uFE8E\uFEE0\uFEB3";

    [Theory]
    [InlineData("גבא")]
    [InlineData(ArabicVisual)]
    public void PaintedGlyphsKeepContextualFormsAndVisualOrder(string visualText) {
        var font = PdfTrueTypeFontProgram.Parse(CreateDirectionalFont(), "Painted");
        var options = PdfTextShapingOptions.ForRendering("Painted", PdfTextShapingMode.OpenTypeLigatures,
            shapingProvider: new RejectUnexpectedShapingProvider(), preservesPaintedGlyphs: true);

        PdfGlyphRun run = font.ShapeText(visualText, options);

        Assert.Equal(GlyphIds(font, visualText), run.Glyphs.Select(static glyph => glyph.GlyphId));
        Assert.Equal(visualText, string.Concat(run.Glyphs.Select(static glyph => glyph.UnicodeText)));
        Assert.Null(run.ActualText);
    }

    [Fact]
    public void LogicalArabicStillUsesTheSharedManagedShaper() {
        var font = PdfTrueTypeFontProgram.Parse(CreateDirectionalFont(), "Logical");
        PdfGlyphRun run = font.ShapeText("سلام", PdfTextShapingOptions.ForRendering("Logical",
            PdfTextShapingMode.OpenTypeLigatures, shapingProvider: OfficeManagedTextShapingProvider.Instance));

        Assert.Equal(GlyphIds(font, ArabicVisual), run.Glyphs.Select(static glyph => glyph.GlyphId));
    }

    [Fact]
    public void PaintedGlyphsDoNotReuseCachedLatinSubstitutions() {
        var font = PdfTrueTypeFontProgram.Parse(
            ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn"), "Ligature");
        var logical = PdfTextShapingOptions.ForRendering("Ligature", PdfTextShapingMode.OpenTypeLigatures);
        Assert.Single(font.ShapeText("fi", logical).Glyphs);

        PdfGlyphRun painted = font.ShapeText("fi", PdfTextShapingOptions.ForRendering("Ligature",
            PdfTextShapingMode.OpenTypeLigatures, preservesPaintedGlyphs: true));

        Assert.Equal(2, painted.Glyphs.Count);
        Assert.Single(font.ShapeText("fi", logical).Glyphs);
    }

    [Fact]
    public void PaintedCffGlyphsBypassContextualSubstitution() {
        string? path = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(path);
        var font = PdfOpenTypeCffFontProgram.Parse(File.ReadAllBytes(path!), "Painted CFF");

        PdfGlyphRun run = font.ShapeText("fi", PdfTextShapingOptions.ForRendering("Painted CFF",
            PdfTextShapingMode.OpenTypeLigatures, shapingProvider: new RejectUnexpectedShapingProvider(),
            preservesPaintedGlyphs: true));

        Assert.Equal(2, run.Glyphs.Count);
        Assert.Equal("fi", string.Concat(run.Glyphs.Select(static glyph => glyph.UnicodeText)));
    }

    [Theory]
    [InlineData(false, "גבא")]
    [InlineData(true, "גבא")]
    [InlineData(false, ArabicVisual)]
    [InlineData(true, ArabicVisual)]
    public void PositionedPdfOutputPreservesPaintedGlyphsAcrossFormattingCopies(bool named, string visualText) {
        byte[] data = CreateDirectionalFont();
        var options = new PdfCreateOptions().UseManagedTextShaping();
        if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Painted", data));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Painted");
        PdfTextRun run = new PdfTextRun(visualText, fontSize: global::OfficeIMO.OfficeLength.Points(12D), fontFamily: named ? "Painted" : null)
            .WithPaintedGlyphPreservation()
            .WithFeatureSettings(OfficeTextFeatureSettings.Default)
            .WithHorizontalOffset(2D)
            .WithCharacterSpacing(global::OfficeIMO.OfficeLength.Points(.1D))
            .WithHorizontalTextScaling(110D)
            .WithTextDirection(OfficeTextDirection.Auto);
        double? width = PdfWriter.MeasurePositionedText(run, options);
        Assert.True(width > 0D);

        byte[] bytes = PdfDocument.Create(options).Canvas(canvas => canvas.PositionedText(
            new[] { run }, PdfCanvasTextStructureRole.Paragraph, 10D, 10D, 200D, 30D,
            fontSize: 12D)).ToBytes();

        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        string painted = string.Concat(independent.GetPage(1).Letters.OrderBy(static letter => letter.Location.X)
            .Select(static letter => letter.Value));
        Assert.Equal(OfficeArabicTextShaper.ToLogicalText(visualText), painted);
    }

    [Theory]
    [InlineData(false, "גבא", false)]
    [InlineData(true, "גבא", false)]
    [InlineData(false, ArabicVisual, false)]
    [InlineData(true, ArabicVisual, false)]
    [InlineData(false, "גבא", true)]
    [InlineData(true, "גבא", true)]
    [InlineData(false, ArabicVisual, true)]
    [InlineData(true, ArabicVisual, true)]
    public void PreflightHonorsPaintedGlyphOwnershipAndRetainsLogicalWarnings(bool named, string text, bool preservesPaintedGlyphs) {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions().SetTextShapingMode(PdfTextShapingMode.LatinLigatures)
            .ReportDiagnosticsTo(report);
        byte[] data = CreateDirectionalFont();
        if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Painted", data));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Painted");
        var run = new PdfTextRun(text, fontSize: global::OfficeIMO.OfficeLength.Points(12D), fontFamily: named ? "Painted" : null)
            .WithPaintedGlyphPreservation(preservesPaintedGlyphs);
        PdfDocument document = PdfDocument.Create(options).Canvas(canvas => canvas.PositionedText(
            new[] { run }, PdfCanvasTextStructureRole.Paragraph, 10D, 10D, 200D, 30D,
            fontSize: 12D));

        Assert.Empty(document.AnalyzeTextEncoding());
        if (preservesPaintedGlyphs) report.RequireNoLoss();
        else Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");

        Assert.NotEmpty(document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        if (preservesPaintedGlyphs) report.RequireNoLoss();
        else Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PaintedGlyphPreflightStillRejectsMissingFontGlyphs(bool named) {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions().SetTextShapingMode(PdfTextShapingMode.LatinLigatures)
            .ReportDiagnosticsTo(report);
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A');
        if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Painted", data));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, data, "Painted");
        var run = new PdfTextRun("א", fontFamily: named ? "Painted" : null)
            .WithPaintedGlyphPreservation();
        PdfDocument document = PdfDocument.Create(options).Canvas(canvas => canvas.PositionedText(
            new[] { run }, PdfCanvasTextStructureRole.Paragraph, 10D, 10D, 200D, 30D));

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding());
        Assert.Equal("U+05D0", diagnostic.CodePoint);
        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
        Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-text-glyph");
    }

    private static byte[] CreateDirectionalFont() => ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
        ' ', 'א', 'ב', 'ג', 'س', 'ل', 'ا', 'م', '\uFEB3', '\uFEE0', '\uFE8E', '\uFEE1');

    private static int[] GlyphIds(PdfTrueTypeFontProgram font, string text) => text.Select(character => {
        Assert.True(font.TryGetGlyphId(character, out int glyph));
        return glyph;
    }).ToArray();

    private sealed class RejectUnexpectedShapingProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) =>
            throw new InvalidOperationException("Already-painted glyphs must not be shaped a second time.");
    }
}
