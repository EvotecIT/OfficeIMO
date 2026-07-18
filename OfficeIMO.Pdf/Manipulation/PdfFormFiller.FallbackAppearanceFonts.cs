namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static bool TryCreateFallbackTextAppearanceFontPlan(
        PdfFormFillerOptions? options,
        string displayValue,
        string diagnosticSource,
        ref int nextObjectNumber,
        out TextAppearanceFontPlan? fontPlan,
        out string? failureMessage) {
        fontPlan = null;
        failureMessage = null;
        if (string.IsNullOrEmpty(displayValue) || options?.AppearanceFontFallbacksSnapshot is not PdfEmbeddedFontFallbackSet fallbackSet) {
            return false;
        }

        PdfTextFallbackPlan textPlan;
        try {
            textPlan = fallbackSet.PlanText(displayValue, diagnosticSource);
        } catch (Exception exception) when (
            exception is NotSupportedException ||
            exception is ArgumentException ||
            exception is ArithmeticException ||
            exception is FormatException ||
            exception is IndexOutOfRangeException ||
            exception is InvalidOperationException) {
            failureMessage = "The configured appearance font fallback set could not be parsed as supported embedded fonts.";
            return false;
        }

        options.AddTextShapingDiagnostics(fallbackSet.AnalyzeAdvancedTextLayout(textPlan.OriginalText, diagnosticSource));

        if (!textPlan.IsFullyCovered) {
            options.AddTextFallbackPlanDiagnostics(textPlan);
            failureMessage = CreateFallbackAppearanceFailureMessage(textPlan);
            return false;
        }

        int[] usedFontIndexes = textPlan.Segments
            .Select(segment => segment.FontIndex)
            .Distinct()
            .OrderBy(index => index)
            .ToArray();
        if (usedFontIndexes.Length == 0) {
            return false;
        }

        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> candidates = fallbackSet.Candidates;
        var programs = new Dictionary<int, FallbackAppearanceFontProgram>();
        var resourceNames = new Dictionary<int, string>();
        var objectNumbers = new Dictionary<int, EmbeddedAppearanceFontObjectNumbers>();
        var appearanceFonts = new PdfDictionary();

        foreach (int fontIndex in usedFontIndexes) {
            if (fontIndex < 0 || fontIndex >= candidates.Count) {
                failureMessage = "The configured appearance font fallback set returned an invalid font candidate index.";
                return false;
            }

            PdfEmbeddedFontFallbackCandidate candidate = candidates[fontIndex];
            FallbackAppearanceFontProgram fontProgram;
            try {
                fontProgram = FallbackAppearanceFontProgram.Parse(candidate);
            } catch (Exception exception) when (
                exception is NotSupportedException ||
                exception is ArgumentException ||
                exception is ArithmeticException ||
                exception is FormatException ||
                exception is IndexOutOfRangeException ||
                exception is InvalidOperationException) {
                failureMessage = $"The configured appearance font fallback candidate '{candidate.FontName}' could not be parsed as a supported embedded font.";
                return false;
            }
            options?.AddFontDiagnostics(fontProgram.AnalyzeFullFontEmbedding(diagnosticSource));

            string resourceName = DefaultAppearanceFontName + fontIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var numbers = new EmbeddedAppearanceFontObjectNumbers(
                nextObjectNumber++,
                nextObjectNumber++,
                nextObjectNumber++,
                nextObjectNumber++,
                nextObjectNumber++);

            programs[fontIndex] = fontProgram;
            resourceNames[fontIndex] = resourceName;
            objectNumbers[fontIndex] = numbers;
            appearanceFonts.Items[resourceName] = new PdfReference(numbers.Type0FontObjectNumber, 0);
        }

        var resources = new PdfDictionary();
        resources.Items["Font"] = appearanceFonts;

        fontPlan = new TextAppearanceFontPlan(
            resourceNames[usedFontIndexes[0]],
            resources,
            encodedTextHex: null,
            encodeTextSegmentHex: null,
            (text, fontSize) => MeasureFallbackAppearanceText(fallbackSet, programs, text, fontSize, diagnosticSource),
            text => EncodeFallbackAppearanceText(fallbackSet, programs, resourceNames, text, diagnosticSource),
            (objects, _) => MaterializeFallbackAppearanceFonts(objects, programs, objectNumbers));
        return true;
    }

    private static List<PdfTextAppearanceSegment> EncodeFallbackAppearanceText(
        PdfEmbeddedFontFallbackSet fallbackSet,
        Dictionary<int, FallbackAppearanceFontProgram> programs,
        Dictionary<int, string> resourceNames,
        string text,
        string diagnosticSource) {
        PdfTextFallbackPlan plan = fallbackSet.PlanText(text, diagnosticSource);
        if (!plan.IsFullyCovered) {
            throw new InvalidOperationException(CreateFallbackAppearanceFailureMessage(plan));
        }

        var encodedSegments = new List<PdfTextAppearanceSegment>();
        foreach (PdfTextFallbackSegment segment in plan.Segments) {
            if (segment.Text.Length == 0) {
                continue;
            }

            if (!programs.TryGetValue(segment.FontIndex, out FallbackAppearanceFontProgram? fontProgram) ||
                !resourceNames.TryGetValue(segment.FontIndex, out string? resourceName)) {
                throw new InvalidOperationException("The configured appearance font fallback plan referenced a font that was not materialized.");
            }

            encodedSegments.Add(new PdfTextAppearanceSegment(resourceName, fontProgram.EncodeTextAsGlyphHex(segment.Text)));
        }

        return encodedSegments;
    }

    private static double MeasureFallbackAppearanceText(
        PdfEmbeddedFontFallbackSet fallbackSet,
        Dictionary<int, FallbackAppearanceFontProgram> programs,
        string text,
        double fontSize,
        string diagnosticSource) {
        PdfTextFallbackPlan plan = fallbackSet.PlanText(text, diagnosticSource);
        if (!plan.IsFullyCovered) {
            throw new InvalidOperationException(CreateFallbackAppearanceFailureMessage(plan));
        }

        double width = 0D;
        foreach (PdfTextFallbackSegment segment in plan.Segments) {
            if (!programs.TryGetValue(segment.FontIndex, out FallbackAppearanceFontProgram? fontProgram)) {
                throw new InvalidOperationException("The configured appearance font fallback plan referenced a font that was not materialized.");
            }

            width += fontProgram.MeasureTextWidth(segment.Text, fontSize);
        }

        return width;
    }

    private static void MaterializeFallbackAppearanceFonts(
        Dictionary<int, PdfIndirectObject> objects,
        Dictionary<int, FallbackAppearanceFontProgram> programs,
        Dictionary<int, EmbeddedAppearanceFontObjectNumbers> objectNumbers) {
        foreach (KeyValuePair<int, EmbeddedAppearanceFontObjectNumbers> entry in objectNumbers.OrderBy(entry => entry.Key)) {
            FallbackAppearanceFontProgram fontProgram = programs[entry.Key];
            EmbeddedAppearanceFontObjectNumbers numbers = entry.Value;
            fontProgram.Materialize(
                objects,
                numbers.FontFileObjectNumber,
                numbers.DescriptorObjectNumber,
                numbers.DescendantFontObjectNumber,
                numbers.ToUnicodeObjectNumber,
                numbers.Type0FontObjectNumber);
        }
    }

    private static string CreateFallbackAppearanceFailureMessage(PdfTextFallbackPlan plan) {
        string details = string.Join(" ", plan.Diagnostics.Select(diagnostic =>
            string.IsNullOrWhiteSpace(diagnostic.Code)
                ? diagnostic.Message
                : diagnostic.Code + ": " + diagnostic.Message));
        return string.IsNullOrWhiteSpace(details)
            ? "The configured appearance font fallback set cannot encode every character required by the form field appearance."
            : details;
    }

    private static bool IsOpenTypeCffFontData(byte[] fontData) =>
        fontData.Length >= 4 &&
        fontData[0] == 0x4F &&
        fontData[1] == 0x54 &&
        fontData[2] == 0x54 &&
        fontData[3] == 0x4F;

    private sealed class FallbackAppearanceFontProgram {
        private readonly PdfTrueTypeFontProgram? _trueTypeFont;
        private readonly PdfOpenTypeCffFontProgram? _cffFont;

        private FallbackAppearanceFontProgram(PdfTrueTypeFontProgram font) {
            _trueTypeFont = font;
            _cffFont = null;
        }

        private FallbackAppearanceFontProgram(PdfOpenTypeCffFontProgram font) {
            _trueTypeFont = null;
            _cffFont = font;
        }

        public static FallbackAppearanceFontProgram Parse(PdfEmbeddedFontFallbackCandidate candidate) {
            byte[] fontData = candidate.DataSnapshot;
            return IsOpenTypeCffFontData(fontData)
                ? new FallbackAppearanceFontProgram(PdfOpenTypeCffFontProgram.Parse(fontData, candidate.FontName))
                : new FallbackAppearanceFontProgram(PdfTrueTypeFontProgram.Parse(fontData, candidate.FontName));
        }

        public string EncodeTextAsGlyphHex(string text) {
            if (_trueTypeFont != null) {
                return _trueTypeFont.EncodeTextAsGlyphHex(text);
            }

            if (_cffFont != null) {
                return _cffFont.EncodeTextAsGlyphHex(text);
            }

            throw new InvalidOperationException("The configured appearance font fallback candidate was not parsed.");
        }

        public double MeasureTextWidth(string text, double fontSize) {
            if (_trueTypeFont != null) {
                return _trueTypeFont.MeasureTextWidth(text, fontSize);
            }

            if (_cffFont != null) {
                return _cffFont.MeasureTextWidth(text, fontSize);
            }

            throw new InvalidOperationException("The configured appearance font fallback candidate was not parsed.");
        }

        public IReadOnlyList<PdfFontEmbeddingDiagnostic> AnalyzeFullFontEmbedding(string source) {
            if (_cffFont != null) {
                return PdfFontDiagnostics.AnalyzeOpenTypeCffCompactEmbedding(_cffFont, source);
            }

            return Array.Empty<PdfFontEmbeddingDiagnostic>();
        }

        public void Materialize(
            Dictionary<int, PdfIndirectObject> objects,
            int fontFileObjectNumber,
            int descriptorObjectNumber,
            int descendantFontObjectNumber,
            int toUnicodeObjectNumber,
            int type0FontObjectNumber) {
            if (_trueTypeFont != null) {
                MaterializeEmbeddedTextAppearanceFont(
                    objects,
                    _trueTypeFont,
                    fontFileObjectNumber,
                    descriptorObjectNumber,
                    descendantFontObjectNumber,
                    toUnicodeObjectNumber,
                    type0FontObjectNumber);
                return;
            }

            if (_cffFont != null) {
                MaterializeEmbeddedTextAppearanceFont(
                    objects,
                    _cffFont,
                    fontFileObjectNumber,
                    descriptorObjectNumber,
                    descendantFontObjectNumber,
                    toUnicodeObjectNumber,
                    type0FontObjectNumber);
                return;
            }

            throw new InvalidOperationException("The configured appearance font fallback candidate was not parsed.");
        }
    }

    private sealed class EmbeddedAppearanceFontObjectNumbers {
        public EmbeddedAppearanceFontObjectNumbers(
            int fontFileObjectNumber,
            int descriptorObjectNumber,
            int descendantFontObjectNumber,
            int toUnicodeObjectNumber,
            int type0FontObjectNumber) {
            FontFileObjectNumber = fontFileObjectNumber;
            DescriptorObjectNumber = descriptorObjectNumber;
            DescendantFontObjectNumber = descendantFontObjectNumber;
            ToUnicodeObjectNumber = toUnicodeObjectNumber;
            Type0FontObjectNumber = type0FontObjectNumber;
        }

        public int FontFileObjectNumber { get; }

        public int DescriptorObjectNumber { get; }

        public int DescendantFontObjectNumber { get; }

        public int ToUnicodeObjectNumber { get; }

        public int Type0FontObjectNumber { get; }
    }
}
