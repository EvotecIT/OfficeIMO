namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double ResolveRichRunLineAlignmentRise(RichSeg segment, PdfOptions? options, PdfLineSpacing? spacing) {
        if (spacing == null || spacing.FontAlignment == PdfFontLineAlignment.Baseline ||
            segment.InlineElement != null && !UsesInlineRunTextMetrics(segment.InlineElement)) return 0D;
        return ResolveFontLineAlignmentRise(segment.Font, segment.NamedFont,
            EffectiveRichFontSize(segment.FontSize, segment.Baseline), options, spacing);
    }

    private static double ResolveFontLineAlignmentRise(PdfStandardFont font, PdfNamedFontFace? namedFont,
        double size, PdfOptions? options, PdfLineSpacing? spacing) {
        if (spacing == null || spacing.FontAlignment == PdfFontLineAlignment.Baseline) return 0D;
        double unit = spacing.FontAlignmentPointsPerPixel;
        double ascent;
        double descent;
        var metrics = ResolveFontLineMetrics(font, namedFont, options);
        if (metrics.HasValue) metrics.Value.GetDeviceMetrics(size, unit, out ascent, out descent);
        else {
            ascent = Math.Round(GetAscenderForOptions(font, namedFont, size, options) / unit, MidpointRounding.AwayFromZero);
            descent = Math.Round(GetDescenderForOptions(font, namedFont, size, options) / unit, MidpointRounding.AwayFromZero);
        }
        double offset = spacing.FontAlignment switch {
            PdfFontLineAlignment.Top => Math.Max(0D, ascent - 1D),
            PdfFontLineAlignment.Center => ascent - Math.Floor((ascent + descent) / 2D),
            PdfFontLineAlignment.Bottom => -Math.Floor(descent / 2D),
            _ => 0D
        };
        return -offset * unit;
    }
}
