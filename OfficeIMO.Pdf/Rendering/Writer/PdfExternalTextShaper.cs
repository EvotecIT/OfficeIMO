using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfExternalTextShaper {
    internal static bool TryShapeText(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        bool automaticLatin = options.ShapingProvider == null && options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures &&
            !OfficeManagedTextShaper.RequiresComplexLayout(text);
        IOfficeTextShapingProvider? provider = options.ShapingProvider;
        if (provider == null && (options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures && !OfficeManagedTextShaper.RequiresComplexLayout(text) || !options.FeatureSettings.IsDefault || options.Direction != OfficeTextDirection.Auto)) provider = OfficeManagedTextShapingProvider.Instance;
        if (provider == null) {
            glyphRun = null!;
            return false;
        }

        OfficeTextShapingResult? result = provider.ShapeText(new OfficeTextShapingRequest(
            text,
            font.FontName,
            font.FontDataForInspection,
            isOpenTypeCff: false,
            font.UnitsPerEm,
            options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
            options.Language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            featureSettings: options.FeatureSettings,
            applyDefaultLatinLigatures: automaticLatin));

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.UnitsPerEm, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null,
            includeActualText: !automaticLatin);
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, false);
        return true;
    }

    internal static bool TryShapeText(string text, PdfOpenTypeCffFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        bool automaticLatin = options.ShapingProvider == null && options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures &&
            !OfficeManagedTextShaper.RequiresComplexLayout(text);
        IOfficeTextShapingProvider? provider = options.ShapingProvider;
        if (provider == null && (options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures && !OfficeManagedTextShaper.RequiresComplexLayout(text) || !options.FeatureSettings.IsDefault || options.Direction != OfficeTextDirection.Auto)) provider = OfficeManagedTextShapingProvider.Instance;
        if (provider == null) {
            glyphRun = null!;
            return false;
        }

        OfficeTextShapingResult? result = provider.ShapeText(new OfficeTextShapingRequest(
            text,
            font.FontName,
            font.FontDataForInspection,
            isOpenTypeCff: true,
            font.UnitsPerEm,
            options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
            options.Language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            featureSettings: options.FeatureSettings,
            applyDefaultLatinLigatures: automaticLatin));

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.UnitsPerEm, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null,
            includeActualText: !automaticLatin);
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, true);
        return true;
    }

    private static PdfGlyphRun BuildGlyphRun(
        string text,
        OfficeTextShapingResult result,
        int glyphCount,
        int unitsPerEm,
        Func<int, int> getGlyphWidth1000,
        Action<int, string>? recordGlyphUsage,
        bool includeActualText) {
        if (result.Glyphs.Count == 0) {
            throw new ArgumentException("PDF text shaping provider returned no glyphs for non-null text.", nameof(result));
        }

        var glyphs = new List<PdfGlyphInfo>(result.Glyphs.Count);
        bool hasCompleteVerticalAdvances = result.Direction == OfficeTextDirection.TopToBottom;
        foreach (OfficeShapedGlyph shapedGlyph in result.Glyphs) {
            if (shapedGlyph.GlyphId <= 0 || shapedGlyph.GlyphId >= glyphCount) {
                throw new ArgumentException("PDF text shaping provider returned glyph id " + shapedGlyph.GlyphId.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", which is outside the embedded font glyph range.", nameof(result));
            }

            if (shapedGlyph.TextIndex < 0 || shapedGlyph.TextIndex > text.Length) {
                throw new ArgumentException("PDF text shaping provider returned a text index outside the source text.", nameof(result));
            }

            int nominalWidth1000 = getGlyphWidth1000(shapedGlyph.GlyphId);
            int advanceWidth1000 = shapedGlyph.AdvanceWidth.HasValue
                ? ScaleToPdfUnits(shapedGlyph.AdvanceWidth.Value, unitsPerEm)
                : nominalWidth1000;
            int advanceHeight1000 = shapedGlyph.AdvanceHeight.HasValue
                ? ScaleToPdfUnits(shapedGlyph.AdvanceHeight.Value, unitsPerEm)
                : 0;
            hasCompleteVerticalAdvances &= shapedGlyph.AdvanceHeight.HasValue;
            int offsetX1000 = ScaleToPdfUnits(shapedGlyph.OffsetX, unitsPerEm);
            int offsetY1000 = ScaleToPdfUnits(shapedGlyph.OffsetY, unitsPerEm);
            recordGlyphUsage?.Invoke(shapedGlyph.GlyphId, shapedGlyph.UnicodeText);
            glyphs.Add(new PdfGlyphInfo(
                shapedGlyph.GlyphId,
                shapedGlyph.UnicodeText,
                shapedGlyph.TextIndex,
                nominalWidth1000,
                advanceWidth1000,
                advanceHeight1000,
                offsetX1000,
                offsetY1000));
        }

        // Automatic Latin runs retain logical clusters in ToUnicode. Avoid broad ActualText
        // scopes so partial redaction can preserve neighboring scalar and ligature glyphs.
        return new PdfGlyphRun(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), actualText: includeActualText ? text : null, result.Direction, hasCompleteVerticalAdvances, result, preserveGlyphUnicode: !includeActualText);
    }

    private static int ScaleToPdfUnits(int value, int unitsPerEm) =>
        checked((int)Math.Round(value * 1000D / unitsPerEm, MidpointRounding.AwayFromZero));
}
