using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    internal sealed record TextDecorationMeasurement(double Position, double Thickness);

    /// <summary>Measures an explicit strike band with the same selected font face and baseline as text painting.</summary>
    internal static TextDecorationMeasurement? MeasureStrikethrough(PdfTextRun run, PdfOptions options,
        double pointsPerPixel) => MeasureFontDecoration(run, options, pointsPerPixel, underline: false);

    /// <summary>Measures an explicit underline band with the same selected font face and baseline as text painting.</summary>
    internal static TextDecorationMeasurement? MeasureUnderline(PdfTextRun run, PdfOptions options,
        double pointsPerPixel) => MeasureFontDecoration(run, options, pointsPerPixel, underline: true);

    private static TextDecorationMeasurement? MeasureFontDecoration(PdfTextRun run, PdfOptions options,
        double pointsPerPixel, bool underline) {
        PdfStandardFont font = ResolveFontForRun(run, ChooseNormal(options.DefaultFont));
        PdfNamedFontFace? namedFont = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic,
            out PdfNamedFontFace resolved) ? resolved : null;
        OfficeOpenTypeDecorationMetrics? metrics = ResolveFontDecorationMetrics(font, namedFont, options);
        OfficeOpenTypeDecorationMetrics.Metric? selected = underline ? metrics?.Underline : metrics?.Strikethrough;
        if (selected is not { } metric) return null;
        double size = run.FontSize ?? options.DefaultFontSize;
        metric.GetDeviceMetrics(EffectiveRichFontSize(size, run.Baseline), pointsPerPixel,
            out double position, out double thickness);
        return new TextDecorationMeasurement(position + TextRiseForBaseline(size, run.Baseline), thickness);
    }

    private static OfficeOpenTypeDecorationMetrics? ResolveFontDecorationMetrics(
        PdfStandardFont font, PdfNamedFontFace? namedFont, PdfOptions options) {
        if (namedFont.HasValue) {
            if (options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? namedTrueType) && namedTrueType != null)
                return namedTrueType.DecorationMetrics;
            if (options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? namedCff) && namedCff != null)
                return namedCff.DecorationMetrics;
        }
        if (options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? trueType) && trueType != null)
            return trueType.DecorationMetrics;
        if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cff) && cff != null)
            return cff.DecorationMetrics;
        return null;
    }
}
