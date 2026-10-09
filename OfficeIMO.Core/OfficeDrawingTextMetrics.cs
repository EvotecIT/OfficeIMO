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
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?> measureSegmentPaint) {
        MeasureSegmentPaintBounds = measureSegmentPaint;
        MeasureText = measure;
        MeasurePaintBounds = measurePaint;
        MeasureHorizontalPaintBounds = measureHorizontalPaint;
        MeasurePositionedText = measurePositioned ?? ((text, size, family, style, _) => measure(text, size, family, style));
    }
    internal Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?> MeasureSegmentPaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, double> MeasureText { get; }
    internal Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> MeasurePaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> MeasureHorizontalPaintBounds { get; }
    internal Func<string?, double, string?, OfficeFontStyle, OfficeTextDirection, double> MeasurePositionedText { get; }
}
