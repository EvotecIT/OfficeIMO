namespace OfficeIMO.Pdf;

internal static class PdfExternalTextShaper {
    internal static bool TryShapeText(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        if (options.ShapingProvider == null) {
            glyphRun = null!;
            return false;
        }

        PdfTextShapingResult? result = options.ShapingProvider.ShapeText(new PdfTextShapingRequest(
            text,
            font.FontName,
            font.FontDataSnapshot,
            isOpenTypeCff: false,
            options.ShapingMode));

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null);
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

        PdfTextShapingResult? result = options.ShapingProvider.ShapeText(new PdfTextShapingRequest(
            text,
            font.FontName,
            font.FontDataSnapshot,
            isOpenTypeCff: true,
            options.ShapingMode));

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null);
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, true);
        return true;
    }

    private static PdfGlyphRun BuildGlyphRun(
        string text,
        PdfTextShapingResult result,
        int glyphCount,
        Func<int, int> getGlyphWidth1000,
        Action<int, string>? recordGlyphUsage) {
        if (result.Glyphs.Count == 0) {
            throw new ArgumentException("PDF text shaping provider returned no glyphs for non-null text.", nameof(result));
        }

        var glyphs = new List<PdfGlyphInfo>(result.Glyphs.Count);
        foreach (PdfShapedGlyph shapedGlyph in result.Glyphs) {
            if (shapedGlyph.GlyphId <= 0 || shapedGlyph.GlyphId >= glyphCount) {
                throw new ArgumentException("PDF text shaping provider returned glyph id " + shapedGlyph.GlyphId.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", which is outside the embedded font glyph range.", nameof(result));
            }

            if (shapedGlyph.TextIndex < 0 || shapedGlyph.TextIndex > text.Length) {
                throw new ArgumentException("PDF text shaping provider returned a text index outside the source text.", nameof(result));
            }

            if (string.IsNullOrEmpty(shapedGlyph.UnicodeText)) {
                throw new ArgumentException("PDF text shaping provider returned a glyph without Unicode extraction text.", nameof(result));
            }

            recordGlyphUsage?.Invoke(shapedGlyph.GlyphId, shapedGlyph.UnicodeText);
            glyphs.Add(new PdfGlyphInfo(
                shapedGlyph.GlyphId,
                shapedGlyph.UnicodeText,
                shapedGlyph.TextIndex,
                getGlyphWidth1000(shapedGlyph.GlyphId)));
        }

        return new PdfGlyphRun(glyphs);
    }
}
