using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOpenTypeDefaultLigatureTests {
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

}
