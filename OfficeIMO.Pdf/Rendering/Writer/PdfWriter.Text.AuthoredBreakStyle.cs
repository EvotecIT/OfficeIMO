namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // An authored break can carry the only font metrics on an otherwise empty
    // line. Retain its style when a wrapped paragraph is split into columns.
    private static RichSeg CreateRichLineBreakSegment(PdfTextRun run, PdfStandardFont font, double fontSize, PdfNamedFontFace? namedFont) =>
        new(string.Empty, run.Bold, run.Italic, run.Underline, run.Strike, run.Color, run.BackgroundColor,
            null, null, null, font, fontSize, run.Baseline, 0D, endsWithHardBreak: true,
            namedFont: namedFont, underlineStyle: run.UnderlineStyle, strikeStyle: run.StrikeStyle,
            decorationColor: run.DecorationColor, featureSettings: run.FeatureSettings, textDirection: run.TextDirection);
}
