namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    internal static double MeasurePositionedText(PdfTextRun run, PdfOptions options) =>
        MeasureRichSegment(CreatePositionedTextSegment(run, options.DefaultFontSize, options), options);

    private static RichSeg CreatePositionedTextSegment(PdfTextRun run, double fontSize, PdfOptions options) {
        PdfStandardFont font = ResolveFontForRun(run, ChooseNormal(options.DefaultFont));
        PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolved)
            ? resolved : null;
        return new RichSeg(run.Text, run.Bold, run.Italic, run.Underline, run.Strike,
            run.Color, run.BackgroundColor, run.LinkUri, run.LinkDestinationName, run.LinkContents,
            font, run.FontSize ?? fontSize, run.Baseline, namedFont: namedFont,
            underlineStyle: run.UnderlineStyle, strikeStyle: run.StrikeStyle,
            decorationColor: run.DecorationColor, featureSettings: run.FeatureSettings);
    }

    private static (List<List<RichSeg>> Lines, List<double> LineHeights) CreatePositionedTextLine(
        IEnumerable<PdfTextRun> runs, double fontSize, double leading, PdfOptions options) {
        var segments = NormalizeFallbackRuns(runs, ChooseNormal(options.DefaultFont), options)
            .Select(run => CreatePositionedTextSegment(run, fontSize, options)).ToList();
        return (new List<List<RichSeg>> { segments }, new List<double> { leading });
    }
}
