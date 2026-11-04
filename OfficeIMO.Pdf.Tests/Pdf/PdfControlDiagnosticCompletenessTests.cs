using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfControlDiagnosticCompletenessTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeclinedCompleteControlsRetainEveryControlAndRawMissingGlyphWithoutSubsetUsage(bool cff) {
        byte[] data = cff ? File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!)
            : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'A');
        const string source = "Mixed raw source";
        const string fontName = "Diagnostic completeness";
        const string text = "\U0010FFFD\u200EA\u200E\U0010FFFD";
        PdfTrueTypeFontProgram? tt = cff ? null : PdfTrueTypeFontProgram.Parse(data, fontName);
        PdfOpenTypeCffFontProgram? ot = cff ? PdfOpenTypeCffFontProgram.Parse(data, fontName) : null;
        var mappingsBefore = cff ? ot!.GetGlyphToUnicodeMappings().ToArray() : tt!.GetGlyphToUnicodeMappings().ToArray();
        // The raw API retains scalar coverage semantics, including the control's
        // missing cmap entry, and never records source mappings or glyph usage.
        // Compare the same diagnostic display name, not the font program's
        // sanitized PDF resource name against the configured display name.
        var raw = cff ? PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, ot!, source, fontName)
            : PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, tt!, source, fontName);
        PdfCreateOptions options = new PdfCreateOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, fontName);
        PdfTextShapingOptions descriptor = PdfTextDiagnostics.ForEncodingDiagnostics(options, fontName, source);
        var complete = cff ? PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, ot!, source, fontName, shapingOptions: descriptor)
            : PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, tt!, source, fontName, shapingOptions: descriptor);

        Assert.Equal(new[] { 0, 2, 4, 5 }, complete.Select(diagnostic => diagnostic.Index));
        Assert.Equal(new[] { "U+10FFFD", "U+200E", "U+200E", "U+10FFFD" },
            complete.Select(diagnostic => diagnostic.CodePoint));
        Assert.All(complete.Where(diagnostic => diagnostic.CodePoint == "U+200E"), diagnostic => {
            Assert.Contains("verified blank glyphs", diagnostic.Remediation, StringComparison.Ordinal);
            Assert.Contains("bounded OpenType shaping", diagnostic.Encoding, StringComparison.Ordinal);
        });
        foreach (PdfTextEncodingDiagnostic diagnostic in complete.Where(diagnostic => diagnostic.CodePoint == "U+10FFFD")) {
            PdfTextEncodingDiagnostic original = Assert.Single(raw.Where(candidate => candidate.Index == diagnostic.Index));
            Assert.Equal(original.Code, diagnostic.Code);
            Assert.Equal(original.Message, diagnostic.Message);
            Assert.Contains("font '" + fontName + "'", diagnostic.Message, StringComparison.Ordinal);
            Assert.Equal(original.Source, diagnostic.Source);
            Assert.Equal(original.Text, diagnostic.Text);
        }
        Assert.Empty(cff ? ot!.GetUsedGlyphIds() : tt!.GetUsedGlyphIds());
        Assert.Equal(mappingsBefore, cff ? ot!.GetGlyphToUnicodeMappings().ToArray() : tt!.GetGlyphToUnicodeMappings().ToArray());
    }
}
