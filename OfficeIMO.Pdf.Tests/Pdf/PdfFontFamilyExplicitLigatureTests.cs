using System;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontFamilyExplicitLigatureTests {
    [Fact]
    public void OpenTypeCffEmbeddedFont_WritesExplicitUnicodeLigatureScalarsWithExtractableText() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontBytes = File.ReadAllBytes(fontPath!);
        PdfOpenTypeFontInfo info = PdfOpenTypeFontInspector.Inspect(fontBytes, "OfficeIMO Source Serif CFF");

        Assert.True(info.ContainsUnicodeScalar(0xFB00));
        Assert.True(info.ContainsUnicodeScalar(0xFB01));
        Assert.True(info.ContainsUnicodeScalar(0xFB02));
        Assert.True(info.ContainsUnicodeScalar(0xFB03));
        Assert.True(info.ContainsUnicodeScalar(0xFB04));

        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO Source Serif CFF");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Explicit ligatures:\uFB00|\uFB01|\uFB02|\uFB03|\uFB04"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Contains("Explicit ligatures:ff|fi|fl|ffi|ffl", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Fact]
    public void TextShapingModeLatinLigatures_WritesCoveredOpenTypeCffLigatureGlyphsWithExtractableText() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontBytes = File.ReadAllBytes(fontPath!);
        PdfOpenTypeCffFontProgram fontProgram = PdfOpenTypeCffFontProgram.Parse(fontBytes, "OfficeIMO Source Serif CFF");
        Assert.True(fontProgram.TryGetGlyphId(0xFB00, out int ffGlyphId));
        Assert.True(fontProgram.TryGetGlyphId(0xFB01, out int fiGlyphId));
        Assert.True(fontProgram.TryGetGlyphId(0xFB02, out int flGlyphId));
        Assert.True(fontProgram.TryGetGlyphId(0xFB03, out int ffiGlyphId));
        Assert.True(fontProgram.TryGetGlyphId(0xFB04, out int fflGlyphId));

        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .SetTextShapingMode(PdfTextShapingMode.LatinLigatures)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO Source Serif CFF");

        const string text = "Ligature shaping: staff fine flow affinity waffle";
        Assert.Empty(PdfTextDiagnostics.AnalyzeGeneratedText(text, options, PdfStandardFont.Helvetica));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", raw, StringComparison.Ordinal);
        Assert.Contains(ffGlyphId.ToString("X4", CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains(fiGlyphId.ToString("X4", CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains(flGlyphId.ToString("X4", CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains(ffiGlyphId.ToString("X4", CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains(fflGlyphId.ToString("X4", CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains(BuildToUnicodeEntry(ffGlyphId, "ff"), raw, StringComparison.Ordinal);
        Assert.Contains(BuildToUnicodeEntry(fiGlyphId, "fi"), raw, StringComparison.Ordinal);
        Assert.Contains(BuildToUnicodeEntry(flGlyphId, "fl"), raw, StringComparison.Ordinal);
        Assert.Contains(BuildToUnicodeEntry(ffiGlyphId, "ffi"), raw, StringComparison.Ordinal);
        Assert.Contains(BuildToUnicodeEntry(fflGlyphId, "ffl"), raw, StringComparison.Ordinal);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-font-ligature-substitution");
    }

    [Fact]
    public void TextShapingModeLatinLigatures_PreservesCoveredLigatureSpanWhenFallbackSplitsAdjacentRun() {
        string? primaryFontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        string? fallbackFontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (primaryFontPath == null || fallbackFontPath == null) {
            return;
        }

        byte[] primaryFontBytes = File.ReadAllBytes(primaryFontPath);
        byte[] fallbackFontBytes = File.ReadAllBytes(fallbackFontPath);
        PdfOpenTypeCffFontProgram primaryFont = PdfOpenTypeCffFontProgram.Parse(primaryFontBytes, "OfficeIMO Source Serif CFF");
        PdfTrueTypeFontProgram fallbackFont = PdfTrueTypeFontProgram.Parse(fallbackFontBytes, "OfficeIMO Fallback Font");
        Assert.True(primaryFont.TryGetGlyphId(0xFB03, out int ffiGlyphId));

        int? fallbackScalar = FindFallbackOnlyScalar(primaryFont, fallbackFont);
        if (!fallbackScalar.HasValue) {
            return;
        }

        string fallbackText = char.ConvertFromUtf32(fallbackScalar.Value);
        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("OfficeIMO Fallback Font", fallbackFontBytes) },
            new[] { PdfStandardFont.TimesRoman });
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .SetTextShapingMode(PdfTextShapingMode.LatinLigatures)
            .EmbedStandardFont(PdfStandardFont.Helvetica, primaryFontBytes, "OfficeIMO Source Serif CFF")
            .RegisterEmbeddedFontFallbacks(fallbackSet);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("office" + fallbackText))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(ffiGlyphId.ToString("X4", CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains("office" + fallbackText, extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void TextShapingModeLatinLigatures_StillReportsUncoveredOpenTypeLigatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontBytes = File.ReadAllBytes(fontPath!);
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .SetTextShapingMode(PdfTextShapingMode.LatinLigatures)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO Source Serif CFF");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Uppercase ligature probe: FINE"))
            .ToBytes();

        Assert.NotEmpty(bytes);
        PdfConversionWarning warning = Assert.Single(report.Warnings, item => item.Code == "unsupported-font-ligature-substitution");
        Assert.Equal("OpenType GSUB ligature", warning.Details["script"]);
        Assert.Equal("U+0046", warning.Details["codePoint"]);
    }

    private static string BuildToUnicodeEntry(int glyphId, string text) {
        var sb = new StringBuilder()
            .Append('<')
            .Append(glyphId.ToString("X4", CultureInfo.InvariantCulture))
            .Append("> <");
        for (int index = 0; index < text.Length; index++) {
            sb.Append(((int)text[index]).ToString("X4", CultureInfo.InvariantCulture));
        }

        return sb.Append('>').ToString();
    }

    private static int? FindFallbackOnlyScalar(PdfOpenTypeCffFontProgram primaryFont, PdfTrueTypeFontProgram fallbackFont) {
        int[] candidates = {
            0x0416,
            0x042F,
            0x05D0,
            0x03A9,
            0x0141,
            0x0119,
            0x20AC,
            0x2192,
            0x4E00
        };

        foreach (int scalar in candidates) {
            if ((!primaryFont.TryGetGlyphId(scalar, out int primaryGlyphId) || primaryGlyphId <= 0) &&
                fallbackFont.TryGetGlyphId(scalar, out int fallbackGlyphId) &&
                fallbackGlyphId > 0) {
                return scalar;
            }
        }

        return null;
    }
}
