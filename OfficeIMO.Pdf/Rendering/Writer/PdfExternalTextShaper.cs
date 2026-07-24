using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfExternalTextShaper {
    internal static bool TryShapeText(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        if (options.ShapingProvider == null) {
            glyphRun = null!;
            return false;
        }

        OfficeTextShapingResult? result = options.ShapingProvider.ShapeText(new OfficeTextShapingRequest(
            text,
            font.FontName,
            font.FontDataForInspection,
            isOpenTypeCff: false,
            font.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(text),
            options.Language,
            default,
            fontCollectionIndex: null,
            cloneFontData: false));

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.UnitsPerEm, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null);
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, false);
        return true;
    }

    internal static bool TryShapeText(string text, PdfOpenTypeCffFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        if (options.ShapingProvider == null) {
            glyphRun = null!;
            return false;
        }

        OfficeTextShapingResult? result = options.ShapingProvider.ShapeText(new OfficeTextShapingRequest(
            text,
            font.FontName,
            font.FontDataForInspection,
            isOpenTypeCff: true,
            font.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(text),
            options.Language,
            default,
            fontCollectionIndex: null,
            cloneFontData: false));

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.UnitsPerEm, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null);
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, true);
        return true;
    }

    private static PdfGlyphRun BuildGlyphRun(
        string text,
        OfficeTextShapingResult result,
        int glyphCount,
        int unitsPerEm,
        Func<int, int> getGlyphWidth1000,
        Action<int, string>? recordGlyphUsage) {
        if (result.Glyphs.Count == 0) {
            throw new ArgumentException("PDF text shaping provider returned no glyphs for non-null text.", nameof(result));
        }

        var glyphs = new List<PdfGlyphInfo>(result.Glyphs.Count);
        foreach (OfficeShapedGlyph shapedGlyph in result.Glyphs) {
            if (shapedGlyph.GlyphId <= 0 || shapedGlyph.GlyphId >= glyphCount) {
                throw new ArgumentException("PDF text shaping provider returned glyph id " + shapedGlyph.GlyphId.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", which is outside the embedded font glyph range.", nameof(result));
            }

            if (shapedGlyph.TextIndex < 0 || shapedGlyph.TextIndex > text.Length) {
                throw new ArgumentException("PDF text shaping provider returned a text index outside the source text.", nameof(result));
            }

            if (string.IsNullOrEmpty(shapedGlyph.UnicodeText)) {
                throw new ArgumentException("PDF text shaping provider returned a glyph without Unicode extraction text.", nameof(result));
            }

            int nominalWidth1000 = getGlyphWidth1000(shapedGlyph.GlyphId);
            int advanceWidth1000 = shapedGlyph.AdvanceWidth.HasValue
                ? ScaleToPdfUnits(shapedGlyph.AdvanceWidth.Value, unitsPerEm)
                : nominalWidth1000;
            int offsetX1000 = ScaleToPdfUnits(shapedGlyph.OffsetX, unitsPerEm);
            int offsetY1000 = ScaleToPdfUnits(shapedGlyph.OffsetY, unitsPerEm);
            recordGlyphUsage?.Invoke(shapedGlyph.GlyphId, shapedGlyph.UnicodeText);
            glyphs.Add(new PdfGlyphInfo(
                shapedGlyph.GlyphId,
                shapedGlyph.UnicodeText,
                shapedGlyph.TextIndex,
                nominalWidth1000,
                advanceWidth1000,
                offsetX1000,
                offsetY1000));
        }

        return new PdfGlyphRun(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), actualText: text);
    }

    private static int ScaleToPdfUnits(int value, int unitsPerEm) =>
        checked((int)Math.Round(value * 1000D / unitsPerEm, MidpointRounding.AwayFromZero));
}
