namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    internal static double? MeasurePositionedText(PdfTextRun run, PdfOptions options) {
        PdfStandardFont baseFont = ChooseNormal(options.DefaultFont);
        var parts = NormalizeFallbackRuns(new[] { run }, baseFont, options).ToList();
        // Automatic fallback fonts may only be registered once the full scene is known.
        if (parts.Any(part => !CanWriteRunWithSelectedFont(part, baseFont, options))) return null;
        return parts.Sum(part => GetRichSegmentWidth(CreatePositionedTextSegment(part, options.DefaultFontSize, options)));
    }

    private static RichSeg CreatePositionedTextSegment(PdfTextRun run, double fontSize, PdfOptions options) {
        PdfStandardFont font = ResolveFontForRun(run, ChooseNormal(options.DefaultFont));
        PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolved)
            ? resolved : null;
        double effectiveFontSize = run.FontSize ?? fontSize;
        double measuredWidth = MeasurePositionedTextWidth(
            run.Text, font, namedFont, effectiveFontSize, run.Baseline, options,
            run.FeatureSettings, run.TextDirection);
        return new RichSeg(run.Text, run.Bold, run.Italic, run.Underline, run.Strike,
            run.Color, run.BackgroundColor, run.LinkUri, run.LinkDestinationName, run.LinkContents,
            font, effectiveFontSize, run.Baseline, measuredWidth, namedFont: namedFont,
            underlineStyle: run.UnderlineStyle, strikeStyle: run.StrikeStyle,
            decorationColor: run.DecorationColor, featureSettings: run.FeatureSettings,
            textDirection: run.TextDirection);
    }

    private static double MeasurePositionedTextWidth(
        string text,
        PdfStandardFont font,
        PdfNamedFontFace? namedFont,
        double fontSize,
        PdfTextBaseline baseline,
        PdfOptions options,
        OfficeIMO.Drawing.OfficeTextFeatureSettings featureSettings,
        OfficeIMO.Drawing.OfficeTextDirection direction) {
        if (direction == OfficeIMO.Drawing.OfficeTextDirection.Auto) {
            return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
        }

        double effectiveFontSize = EffectiveRichFontSize(fontSize, baseline);
        PdfTextShapingOptions shapingOptions;
        if (namedFont.HasValue &&
            options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? namedTrueType) &&
            namedTrueType != null) {
            shapingOptions = PdfTextShapingOptions.ForRendering(
                namedTrueType.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return namedTrueType.ShapeText(text, shapingOptions).TotalAdvanceWidth1000 * effectiveFontSize / 1000D;
        }
        if (namedFont.HasValue &&
            options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? namedCff) &&
            namedCff != null) {
            shapingOptions = PdfTextShapingOptions.ForRendering(
                namedCff.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return namedCff.ShapeText(text, shapingOptions).TotalAdvanceWidth1000 * effectiveFontSize / 1000D;
        }
        if (options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? standardTrueType) &&
            standardTrueType != null) {
            shapingOptions = PdfTextShapingOptions.ForRendering(
                standardTrueType.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return standardTrueType.ShapeText(text, shapingOptions).TotalAdvanceWidth1000 * effectiveFontSize / 1000D;
        }
        if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? standardCff) &&
            standardCff != null) {
            shapingOptions = PdfTextShapingOptions.ForRendering(
                standardCff.FontName, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot,
                options.RecordProviderShapedTextRunDelegate, options.Language, featureSettings, direction);
            return standardCff.ShapeText(text, shapingOptions).TotalAdvanceWidth1000 * effectiveFontSize / 1000D;
        }
        return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings);
    }

    private static (List<List<RichSeg>> Lines, List<double> LineHeights) CreatePositionedTextLine(
        IEnumerable<PdfTextRun> runs, double fontSize, double leading, PdfOptions options) {
        var segments = NormalizeFallbackRuns(runs, ChooseNormal(options.DefaultFont), options)
            .Select(run => CreatePositionedTextSegment(run, fontSize, options)).ToList();
        return (new List<List<RichSeg>> { segments }, new List<double> { leading });
    }
}
