using System;

namespace OfficeIMO.Drawing;

// Paired callbacks keep width and ink geometry on the same selected font/shaping path.
internal sealed class OfficeDrawingTextMetrics {
    internal OfficeDrawingTextMetrics(Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> measurePaint,
        Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> measureHorizontalPaint,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextDirection, double>? measurePositioned = null)
        : this(measure, measurePaint, measureHorizontalPaint, measurePositioned,
            (segment, size) => OfficeDrawingTextLayout.MeasureSegmentPaintBounds(segment, size, measurePaint)) {
    }

    // Preserve the original friend-callable constructor; adapters may supply their
    // renderer's complete segment paint without changing public layout contracts.
    internal OfficeDrawingTextMetrics(Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> measurePaint,
        Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> measureHorizontalPaint,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextDirection, double>? measurePositioned,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?> measureSegmentPaint)
        : this(measure, measurePaint, measureHorizontalPaint, measurePositioned, measureSegmentPaint, null) {
    }

    internal OfficeDrawingTextMetrics(Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> measurePaint,
        Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> measureHorizontalPaint,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextDirection, double>? measurePositioned,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?> measureSegmentPaint,
        Func<string?, double, string?, OfficeFontStyle, OfficePublisherFontMetrics?>? measurePublisherFont)
        : this(measure, measurePaint, measureHorizontalPaint, measurePositioned, measureSegmentPaint, measurePublisherFont, null, null) {
    }

    internal OfficeDrawingTextMetrics(Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> measurePaint,
        Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> measureHorizontalPaint,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextDirection, double>? measurePositioned,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?> measureSegmentPaint,
        Func<string?, double, string?, OfficeFontStyle, OfficePublisherFontMetrics?>? measurePublisherFont,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextFeatureSettings?, double>? measureFeatures,
        Func<OfficeRichTextSegment, double, (double Left, double Right)>? measureSegmentHorizontalPaint) {
        RichTextMeasurement = new OfficeRichTextMeasurement(measure, measureFeatures);
        MeasureSegmentHorizontalPaintBounds = measureSegmentHorizontalPaint ??
            ((segment, size) => measureHorizontalPaint(segment.Text, size, segment.FontFamily, segment.FontStyle));
        MeasurePublisherFont = measurePublisherFont;
        MeasureSegmentPaintBounds = measureSegmentPaint;
        MeasureText = measure;
        MeasurePaintBounds = measurePaint;
        MeasureHorizontalPaintBounds = measureHorizontalPaint;
        MeasurePositionedText = measurePositioned ?? ((text, size, family, style, _) => measure(text, size, family, style));
    }
    internal OfficeRichTextMeasurement RichTextMeasurement { get; }
    internal Func<OfficeRichTextSegment, double, (double Left, double Right)> MeasureSegmentHorizontalPaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, OfficePublisherFontMetrics?>? MeasurePublisherFont { get; }
    internal Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?> MeasureSegmentPaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, double> MeasureText { get; }
    internal Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> MeasurePaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> MeasureHorizontalPaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, OfficeTextDirection, double> MeasurePositionedText { get; }
}
