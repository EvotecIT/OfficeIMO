using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable font preflight helpers for generated PDF output.
/// </summary>
public static class PdfFontDiagnostics {
    private const uint TrueTypeScalerType = 0x00010000;
    private const uint AppleTrueTypeScalerType = 0x74727565;
    private const uint OpenTypeCffScalerType = 0x4F54544F;
    private const uint TrueTypeCollectionScalerType = 0x74746366;

    /// <summary>
    /// Finds caller-supplied embedded font data that cannot be used by the current generated font paths.
    /// </summary>
    /// <param name="fontData">Font bytes to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a generated font slot, field, sheet, slide, or converter area.</param>
    /// <param name="fontName">Optional configured font name used in diagnostic messages.</param>
    /// <returns>Font diagnostics in source order. An empty result means the font parsed through a current generated font path.</returns>
    public static IReadOnlyList<PdfFontEmbeddingDiagnostic> AnalyzeEmbeddedFont(byte[] fontData, string source = "", string? fontName = null) {
        Guard.NotNull(fontData, nameof(fontData));
        return AnalyzeEmbeddedFontFailure(fontData, source, fontName, exception: null);
    }

    internal static IReadOnlyList<PdfFontEmbeddingDiagnostic> AnalyzeEmbeddedFontFailure(byte[] fontData, string source, string? fontName, Exception? exception) {
        Guard.NotNull(fontData, nameof(fontData));
        string resolvedFontName = string.IsNullOrWhiteSpace(fontName) ? "embedded font" : fontName!;

        if (fontData.Length < 12) {
            return Single(source, resolvedFontName, "Unknown", "unsupported-embedded-font-data", "Embedded font data is too small to parse as a supported TrueType font.");
        }

        uint scalerType = ReadUInt32(fontData, 0);
        if (scalerType == OpenTypeCffScalerType) {
            try {
                _ = PdfOpenTypeCffFontProgram.Parse(fontData, resolvedFontName);
                return Array.Empty<PdfFontEmbeddingDiagnostic>();
            } catch (Exception parseException) when (IsFontProgramException(parseException)) {
                Exception effectiveException = exception ?? parseException;
                return Single(source, resolvedFontName, "OpenType/CFF", "unsupported-opentype-cff-font", "OpenType/CFF font data could not be parsed by OfficeIMO.Pdf: " + effectiveException.Message);
            }
        }

        if (scalerType == TrueTypeCollectionScalerType) {
            return Single(source, resolvedFontName, "TrueType Collection", "unsupported-truetype-collection-font", "TrueType collection files cannot be embedded directly by this generated font path. Load or extract a single TrueType face before embedding.");
        }

        if (scalerType != TrueTypeScalerType && scalerType != AppleTrueTypeScalerType) {
            string format = "0x" + scalerType.ToString("X8", CultureInfo.InvariantCulture);
            return Single(source, resolvedFontName, format, "unsupported-embedded-font-format", "Embedded font data uses unsupported scaler type " + format + ". Use a single-face TrueType font with glyf outlines or OpenType/CFF font with an OTTO scaler.");
        }

        try {
            _ = PdfTrueTypeFontProgram.Parse(fontData, resolvedFontName);
            return Array.Empty<PdfFontEmbeddingDiagnostic>();
        } catch (Exception parseException) when (IsFontProgramException(parseException)) {
            Exception effectiveException = exception ?? parseException;
            return Single(source, resolvedFontName, "TrueType", "unsupported-truetype-font", effectiveException.Message);
        }
    }

    internal static IReadOnlyList<PdfFontEmbeddingDiagnostic> AnalyzeOpenTypeCffFullFontEmbedding(PdfOpenTypeCffFontProgram font, string source) {
        Guard.NotNull(font, nameof(font));
        IReadOnlyList<int> usedGlyphIds = font.GetUsedGlyphIds();
        var details = new Dictionary<string, string> {
            ["glyphCount"] = font.GlyphCount.ToString(CultureInfo.InvariantCulture),
            ["usedGlyphCount"] = usedGlyphIds.Count.ToString(CultureInfo.InvariantCulture),
            ["fontFileLength"] = font.FontDataLength.ToString(CultureInfo.InvariantCulture),
            ["cffTableLength"] = font.CffTableLength.ToString(CultureInfo.InvariantCulture)
        };

        string message = "OpenType/CFF font '" + font.FontName + "' is embedded as a full /FontFile3 stream because OfficeIMO.Pdf does not yet subset CFF charstrings. Used-glyph widths and /ToUnicode mappings remain limited to the generated glyph usage.";
        return new[] {
            new PdfFontEmbeddingDiagnostic(
                source,
                font.FontName,
                "OpenType/CFF",
                "opentype-cff-full-font-embedded",
                message,
                PdfConversionWarningSeverity.Warning,
                PdfLayoutDiagnosticKind.SimplifiedContent,
                details)
        };
    }

    internal static bool IsFontProgramException(Exception exception) =>
        exception is NotSupportedException ||
        exception is ArgumentException ||
        exception is ArithmeticException ||
        exception is FormatException ||
        exception is IndexOutOfRangeException ||
        exception is InvalidOperationException;

    private static PdfFontEmbeddingDiagnostic[] Single(string source, string fontName, string format, string code, string message) =>
        new[] { new PdfFontEmbeddingDiagnostic(source, fontName, format, code, message) };

    private static uint ReadUInt32(byte[] data, int offset) {
        return ((uint)data[offset] << 24) |
            ((uint)data[offset + 1] << 16) |
            ((uint)data[offset + 2] << 8) |
            data[offset + 3];
    }
}
