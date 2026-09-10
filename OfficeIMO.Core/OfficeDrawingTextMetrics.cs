using System;

namespace OfficeIMO.Drawing;

// Paired callbacks keep width and ink geometry on the same selected font/shaping path.
internal sealed class OfficeDrawingTextMetrics {
    internal OfficeDrawingTextMetrics(Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> measurePaint) {
        MeasureText = measure; MeasurePaintBounds = measurePaint;
    }
    internal Func<string?, double, string?, OfficeFontStyle, double> MeasureText { get; }
    internal Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds> MeasurePaintBounds { get; }
}
