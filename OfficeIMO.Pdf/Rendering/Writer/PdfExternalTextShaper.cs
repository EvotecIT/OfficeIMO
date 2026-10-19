using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfExternalTextShaper {
    internal static bool TryShapeText(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));
        if (options.PreservesPaintedGlyphs) {
            glyphRun = null!;
            return false;
        }

        bool automaticLatin = UsesDefaultLatinShaping(text, options);
        if (automaticLatin && options.FeatureSettings.IsDefault) return TryShapeDefaultLatinText(
            text, font.FontName, font.FontDataForInspection, false, font.UnitsPerEm, font.GlyphCount,
            font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null, options, out glyphRun);
        IOfficeTextShapingProvider? provider = options.ShapingProvider;
        // Right-to-left and joining scripts are never written in raw logical order: the managed
        // provider applies Arabic joining and bidi ordering, and declines scripts it cannot shape.
        // Right-to-left and joining text selected automatically keeps the default Latin ligatures of its Latin segments.
        bool automaticDirectional = provider == null && RequiresManagedDirectionalShaping(text, options);
        if (provider == null && (automaticLatin || !options.FeatureSettings.IsDefault || options.Direction != OfficeTextDirection.Auto ||
            automaticDirectional)) provider = OfficeManagedTextShapingProvider.Instance;
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
            automaticLatin ? options.Direction : options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
            options.Language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            featureSettings: automaticDirectional && options.FeatureSettings.IsDefault ? AutomaticDirectionalFeatures : options.FeatureSettings,
            applyDefaultLatinLigatures: automaticLatin));
        // Fonts whose substitutions the managed engine cannot apply in full still get joining and ordering.
        if (result == null && automaticDirectional && options.FeatureSettings.IsDefault) result = provider.ShapeText(new OfficeTextShapingRequest(
            text, font.FontName, font.FontDataForInspection, isOpenTypeCff: false, font.UnitsPerEm,
            options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
            options.Language, default, fontCollectionIndex: null, variationCoordinates: null, cloneFontData: false));

        if (result == null && options.ShapingProvider != null && options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures) {
            automaticLatin = true;
            result = OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
                text, font.FontName, font.FontDataForInspection, isOpenTypeCff: false, font.UnitsPerEm,
                automaticLatin ? options.Direction : options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
                options.Language, default, fontCollectionIndex: null, variationCoordinates: null,
                cloneFontData: false, featureSettings: options.FeatureSettings, applyDefaultLatinLigatures: true));
        }

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.UnitsPerEm, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null,
            includeActualText: !automaticLatin || result.Direction is not (OfficeTextDirection.LeftToRight or OfficeTextDirection.Auto) || OfficeManagedTextShaper.RequiresComplexLayout(text));
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, false, automaticLatin);
        return true;
    }

    internal static bool TryShapeText(string text, PdfOpenTypeCffFontProgram font, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));
        if (options.PreservesPaintedGlyphs) {
            glyphRun = null!;
            return false;
        }

        bool automaticLatin = UsesDefaultLatinShaping(text, options);
        if (automaticLatin && options.FeatureSettings.IsDefault) return TryShapeDefaultLatinText(
            text, font.FontName, font.FontDataForInspection, true, font.UnitsPerEm, font.GlyphCount,
            font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null, options, out glyphRun);
        IOfficeTextShapingProvider? provider = options.ShapingProvider;
        // Right-to-left and joining scripts are never written in raw logical order: the managed
        // provider applies Arabic joining and bidi ordering, and declines scripts it cannot shape.
        // Right-to-left and joining text selected automatically keeps the default Latin ligatures of its Latin segments.
        bool automaticDirectional = provider == null && RequiresManagedDirectionalShaping(text, options);
        if (provider == null && (automaticLatin || !options.FeatureSettings.IsDefault || options.Direction != OfficeTextDirection.Auto ||
            automaticDirectional)) provider = OfficeManagedTextShapingProvider.Instance;
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
            automaticLatin ? options.Direction : options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
            options.Language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            featureSettings: automaticDirectional && options.FeatureSettings.IsDefault ? AutomaticDirectionalFeatures : options.FeatureSettings,
            applyDefaultLatinLigatures: automaticLatin));
        // Fonts whose substitutions the managed engine cannot apply in full still get joining and ordering.
        if (result == null && automaticDirectional && options.FeatureSettings.IsDefault) result = provider.ShapeText(new OfficeTextShapingRequest(
            text, font.FontName, font.FontDataForInspection, isOpenTypeCff: true, font.UnitsPerEm,
            options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
            options.Language, default, fontCollectionIndex: null, variationCoordinates: null, cloneFontData: false));

        if (result == null && options.ShapingProvider != null && options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures) {
            automaticLatin = true;
            result = OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
                text, font.FontName, font.FontDataForInspection, isOpenTypeCff: true, font.UnitsPerEm,
                automaticLatin ? options.Direction : options.Direction == OfficeTextDirection.Auto ? OfficeTextElements.ResolveBaseDirection(text) : options.Direction,
                options.Language, default, fontCollectionIndex: null, variationCoordinates: null,
                cloneFontData: false, featureSettings: options.FeatureSettings, applyDefaultLatinLigatures: true));
        }

        if (result == null) {
            glyphRun = null!;
            return false;
        }

        glyphRun = BuildGlyphRun(text, result, font.GlyphCount, font.UnitsPerEm, font.GetGlyphWidth1000, options.RecordGlyphUsage ? font.RecordGlyphUsage : null,
            includeActualText: !automaticLatin || result.Direction is not (OfficeTextDirection.LeftToRight or OfficeTextDirection.Auto) || OfficeManagedTextShaper.RequiresComplexLayout(text));
        options.ProviderShapedTextRecorder?.Invoke(text, font.FontName, true, automaticLatin);
        return true;
    }

    // Default ligatures stay on for the Latin segments of automatically shaped right-to-left text.
    private static readonly OfficeTextFeatureSettings AutomaticDirectionalFeatures = OfficeTextFeatureSettings.Default.With("liga", 1);

    internal static bool RequiresManagedDirectionalShaping(string text, PdfTextShapingOptions options) =>
        options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures &&
        (ContainsRightToLeftCharacter(text) || OfficeTextElements.ContainsJoiningScript(text));

    // Right-to-left letters and marks, not invisible directional controls such as U+200F.
    private static bool ContainsRightToLeftCharacter(string text) {
        for (int index = 0; index < text.Length; index++) {
            int scalar = char.IsSurrogatePair(text, index) ? char.ConvertToUtf32(text[index], text[++index]) : text[index];
            if (OfficeUnicodeTextData.GetBidiClass(scalar) is OfficeBidiClass.R or OfficeBidiClass.AL &&
                System.Globalization.CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(scalar), 0) != System.Globalization.UnicodeCategory.Format) return true;
        }
        return false;
    }

    // Selecting our managed provider must not replace default Latin ligatures
    // with nominal glyphs when an explicit feature (such as kerning) is present.
    // Other providers retain ownership of their substitutions and positioning.
    private static bool UsesDefaultLatinShaping(string text, PdfTextShapingOptions options) =>
        options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures &&
        (options.ShapingProvider == null && !RequiresManagedDirectionalShaping(text, options) ||
         ReferenceEquals(options.ShapingProvider, OfficeManagedTextShapingProvider.Instance) &&
         (options.Direction == OfficeTextDirection.Auto || options.Direction == OfficeTextDirection.LeftToRight) &&
         !OfficeManagedTextShaper.RequiresComplexLayout(text));

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
            advanceWidth1000 = checked(advanceWidth1000 + ScaleToPdfUnits(result.GetAdvanceAdjustment(glyphs.Count), unitsPerEm));
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
                offsetY1000, result.GetLogicalClusterStart(glyphs.Count)));
        }

        // Automatic Latin runs retain logical clusters in ToUnicode. Avoid broad ActualText
        // scopes so partial redaction can preserve neighboring scalar and ligature glyphs.
        // A purely right-to-left run in exact reverse cluster order is what conforming readers
        // reorder themselves; a run-wide ActualText would make them reverse it a second time.
        if (includeActualText && result.Direction == OfficeTextDirection.RightToLeft && IsReversedLogicalRun(text, glyphs)) includeActualText = false;
        return new PdfGlyphRun(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), actualText: includeActualText ? text : null, result.Direction, hasCompleteVerticalAdvances, result, preserveGlyphUnicode: !includeActualText);
    }

    internal static bool IsReversedLogicalRun(string text, IReadOnlyList<PdfGlyphInfo> glyphs) {
        int expectedEnd = text.Length;
        foreach (PdfGlyphInfo glyph in glyphs) {
            if (string.IsNullOrEmpty(glyph.UnicodeText)) continue;
            int start = glyph.TextIndex;
            if (start + glyph.UnicodeText.Length != expectedEnd ||
                string.CompareOrdinal(text, start, glyph.UnicodeText, 0, glyph.UnicodeText.Length) != 0) return false;
            expectedEnd = start;
        }
        if (expectedEnd != 0) return false;
        int[] scalars = OfficeUnicodeBidi.ToScalars(text, out _);
        foreach (int scalar in scalars) {
            if (OfficeUnicodeTextData.GetBidiClass(scalar) is OfficeBidiClass.L or OfficeBidiClass.EN or OfficeBidiClass.AN or
                OfficeBidiClass.LRE or OfficeBidiClass.RLE or OfficeBidiClass.LRO or OfficeBidiClass.RLO or OfficeBidiClass.PDF or
                OfficeBidiClass.LRI or OfficeBidiClass.RLI or OfficeBidiClass.FSI or OfficeBidiClass.PDI) return false;
        }
        return true;
    }

    internal static int ScaleToPdfUnits(int value, int unitsPerEm) =>
        checked((int)Math.Round(value * 1000D / unitsPerEm, MidpointRounding.AwayFromZero));
}
