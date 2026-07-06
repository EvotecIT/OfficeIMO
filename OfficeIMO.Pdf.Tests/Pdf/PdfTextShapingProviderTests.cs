using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextShapingProviderTests {
    [Fact]
    public void TextShapingProvider_ShapesEmbeddedTrueTypeComplexScriptWithoutUnsupportedWarnings() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Font");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var provider = new MappingTextShapingProvider(text, isOpenTypeCff: false, CreateGlyphMap(text, fontProgram));
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Provider Font")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Load(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_DoesNotSuppressWarningsWhenProviderDeclinesRun() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Font");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var provider = new DecliningTextShapingProvider();
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Provider Font")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Load(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_SuppressesWarningsForAutomaticallyPlannedFallbackRun() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Fallback");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("OfficeIMO Provider Fallback", fontData) },
            new[] { PdfStandardFont.TimesRoman });
        var provider = new MappingTextShapingProvider(text, isOpenTypeCff: false, CreateGlyphMap(text, fontProgram));
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Load(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_DoesNotSuppressFallbackWarningsWhenProviderDeclinesRun() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Fallback");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("OfficeIMO Provider Fallback", fontData) },
            new[] { PdfStandardFont.TimesRoman });
        var provider = new DecliningTextShapingProvider();
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Load(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_MapsOpenTypeCffLigatureGlyphBackToSourceText() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontData = File.ReadAllBytes(fontPath!);
        PdfOpenTypeCffFontProgram fontProgram = PdfOpenTypeCffFontProgram.Parse(fontData, "OfficeIMO Source Serif CFF");
        Assert.True(fontProgram.TryGetGlyphId('o', out int oGlyphId));
        Assert.True(fontProgram.TryGetGlyphId(0xFB03, out int ffiGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('c', out int cGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('e', out int eGlyphId));

        var provider = new MappingTextShapingProvider(
            "office",
            isOpenTypeCff: true,
            new[] {
                new PdfShapedGlyph(oGlyphId, "o", 0),
                new PdfShapedGlyph(ffiGlyphId, "ffi", 1),
                new PdfShapedGlyph(cGlyphId, "c", 4),
                new PdfShapedGlyph(eGlyphId, "e", 5)
            });
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif CFF")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("office"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Load(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains("office", extracted, StringComparison.Ordinal);
        Assert.Contains("<" + ffiGlyphId.ToString("X4", CultureInfo.InvariantCulture) + "> <006600660069>", raw, StringComparison.Ordinal);
    }

    private static IReadOnlyList<PdfShapedGlyph> CreateGlyphMap(string text, PdfTrueTypeFontProgram fontProgram) {
        var glyphs = new List<PdfShapedGlyph>();
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            Assert.True(fontProgram.TryGetGlyphId(scalar, out int glyphId));
            glyphs.Add(new PdfShapedGlyph(glyphId, char.ConvertFromUtf32(scalar), scalarStart));
        }

        return glyphs;
    }

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
    }

    private sealed class MappingTextShapingProvider : IPdfTextShapingProvider {
        private readonly string _text;
        private readonly bool _isOpenTypeCff;
        private readonly IReadOnlyList<PdfShapedGlyph> _glyphs;

        public MappingTextShapingProvider(string text, bool isOpenTypeCff, IReadOnlyList<PdfShapedGlyph> glyphs) {
            _text = text;
            _isOpenTypeCff = isOpenTypeCff;
            _glyphs = glyphs;
        }

        public int CallCount { get; private set; }

        public PdfTextShapingResult? ShapeText(PdfTextShapingRequest request) {
            if (!string.Equals(request.Text, _text, StringComparison.Ordinal)) {
                return null;
            }

            Assert.Equal(_isOpenTypeCff, request.IsOpenTypeCff);
            Assert.NotEmpty(request.FontData);
            Assert.False(string.IsNullOrWhiteSpace(request.FontName));
            CallCount++;
            return new PdfTextShapingResult(_glyphs);
        }
    }

    private sealed class DecliningTextShapingProvider : IPdfTextShapingProvider {
        public int CallCount { get; private set; }

        public PdfTextShapingResult? ShapeText(PdfTextShapingRequest request) {
            CallCount++;
            return null;
        }
    }
}
