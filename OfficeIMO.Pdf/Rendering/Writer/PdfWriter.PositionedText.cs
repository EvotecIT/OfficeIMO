namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    internal static double? MeasurePositionedText(PdfTextRun run, PdfOptions options) {
        PdfStandardFont baseFont = ChooseNormal(options.DefaultFont);
        var parts = NormalizeFallbackRuns(new[] { run }, baseFont, options).ToList();
        // Automatic fallback fonts may only be registered once the full scene is known.
        if (parts.Any(part => !CanWriteRunWithSelectedFont(part, baseFont, options))) return null;
        return parts.Sum(part => GetRichSegmentWidth(CreatePositionedTextSegment(part, options.DefaultFontSize, options)));
    }

    private static RichSeg CreatePositionedTextSegment(PdfTextRun run, double fontSize, PdfOptions options, double fontMetricScale = 1D) {
        PdfStandardFont font = ResolveFontForRun(run, ChooseNormal(options.DefaultFont));
        PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolved)
            ? resolved : null;
        double effectiveFontSize = run.FontSize ?? fontSize;
        double measuredWidth = MeasurePositionedTextWidth(
            run.Text, font, namedFont, effectiveFontSize, run.Baseline, options,
            run.FeatureSettings, run.TextDirection, fontMetricScale);
        return new RichSeg(run.Text, run.Bold, run.Italic, run.Underline, run.Strike,
            run.Color, run.BackgroundColor, run.LinkUri, run.LinkDestinationName, run.LinkContents,
            font, effectiveFontSize, run.Baseline, measuredWidth, namedFont: namedFont,
            underlineStyle: run.UnderlineStyle, strikeStyle: run.StrikeStyle,
            decorationColor: run.DecorationColor, featureSettings: run.FeatureSettings,
            textDirection: run.TextDirection, fontMetricScale: fontMetricScale);
    }

    private static double MeasurePositionedTextWidth(
        string text,
        PdfStandardFont font,
        PdfNamedFontFace? namedFont,
        double fontSize,
        PdfTextBaseline baseline,
        PdfOptions options,
        OfficeIMO.Drawing.OfficeTextFeatureSettings featureSettings,
        OfficeIMO.Drawing.OfficeTextDirection direction, double fontMetricScale) {
        if (direction == OfficeIMO.Drawing.OfficeTextDirection.Auto && fontMetricScale == 1D) {
            return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
        }

        double effectiveFontSize = EffectiveRichFontSize(fontSize, baseline);
        PdfTextShapingOptions shapingOptions;
        if (namedFont.HasValue &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? namedTrueType) &&
            namedTrueType != null) {
            if (direction == OfficeIMO.Drawing.OfficeTextDirection.Auto && !namedTrueType.HasTracking)
                return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
            shapingOptions = PdfTextShapingOptions.ForRendering(
                namedTrueType.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return namedTrueType.MeasureShapedTextWidth(text, namedTrueType.ShapeText(text, shapingOptions), effectiveFontSize, fontMetricScale);
        }
        if (namedFont.HasValue &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? namedCff) &&
            namedCff != null) {
            if (direction == OfficeIMO.Drawing.OfficeTextDirection.Auto)
                return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
            shapingOptions = PdfTextShapingOptions.ForRendering(
                namedCff.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return namedCff.ShapeText(text, shapingOptions).TotalAdvanceWidth1000 * effectiveFontSize / 1000D;
        }
        if (options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? standardTrueType) &&
            standardTrueType != null) {
            if (direction == OfficeIMO.Drawing.OfficeTextDirection.Auto && !standardTrueType.HasTracking)
                return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
            shapingOptions = PdfTextShapingOptions.ForRendering(
                standardTrueType.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return standardTrueType.MeasureShapedTextWidth(text, standardTrueType.ShapeText(text, shapingOptions), effectiveFontSize, fontMetricScale);
        }
        if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? standardCff) &&
            standardCff != null) {
            if (direction == OfficeIMO.Drawing.OfficeTextDirection.Auto)
                return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
            shapingOptions = PdfTextShapingOptions.ForRendering(
                standardCff.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return standardCff.ShapeText(text, shapingOptions).TotalAdvanceWidth1000 * effectiveFontSize / 1000D;
        }
        return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
    }

    private static (List<List<RichSeg>> Lines, List<double> LineHeights) CreatePositionedTextLine(
        IEnumerable<PdfTextRun> runs, double fontSize, double leading, PdfOptions options, double fontMetricScale = 1D) {
        var segments = NormalizeFallbackRuns(runs, ChooseNormal(options.DefaultFont), options)
            .Select(run => CreatePositionedTextSegment(run, fontSize, options, fontMetricScale)).ToList();
        return (new List<List<RichSeg>> { segments }, new List<double> { leading });
    }
}
