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
            options.ShapingMode,
            font.UnitsPerEm,
            PdfTextDirectionResolver.Resolve(text),
            options.Language));

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

        PdfTextShapingResult? result = options.ShapingProvider.ShapeText(new PdfTextShapingRequest(
            text,
            font.FontName,
            font.FontDataSnapshot,
            isOpenTypeCff: true,
            options.ShapingMode,
            font.UnitsPerEm,
            PdfTextDirectionResolver.Resolve(text),
            options.Language));

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
        PdfTextShapingResult result,
        int glyphCount,
        int unitsPerEm,
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

internal static class PdfTextDirectionResolver {
    internal static PdfTextDirection Resolve(string text) {
        for (int index = 0; index < text.Length;) {
            int scalar = ReadScalar(text, ref index);
            if (IsRightToLeft(scalar)) {
                return PdfTextDirection.RightToLeft;
            }

            if (IsLeftToRight(scalar)) {
                return PdfTextDirection.LeftToRight;
            }
        }

        return PdfTextDirection.Auto;
    }

    private static bool IsRightToLeft(int scalar) =>
        (scalar >= 0x0590 && scalar <= 0x08FF) ||
        (scalar >= 0xFB1D && scalar <= 0xFDFF) ||
        (scalar >= 0xFE70 && scalar <= 0xFEFF) ||
        (scalar >= 0x10800 && scalar <= 0x10FFF) ||
        (scalar >= 0x1E800 && scalar <= 0x1EEFF);

    private static bool IsLeftToRight(int scalar) =>
        (scalar >= 'A' && scalar <= 'Z') ||
        (scalar >= 'a' && scalar <= 'z') ||
        (scalar >= 0x00C0 && scalar <= 0x058F) ||
        (scalar >= 0x0900 && scalar <= 0x1FFF) ||
        (scalar >= 0x2C00 && scalar <= 0xA7FF) ||
        (scalar >= 0x10000 && scalar <= 0x107FF);

    private static int ReadScalar(string text, ref int index) {
        char first = text[index++];
        return char.IsHighSurrogate(first) && index < text.Length && char.IsLowSurrogate(text[index])
            ? char.ConvertToUtf32(first, text[index++])
            : first;
    }
}
