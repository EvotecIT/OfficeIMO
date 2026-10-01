using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        private LayoutBox Text(string text, double scale) {
            double size = FontSize(scale);
            OfficeTextMeasurementStyle style = _measurer.CreateStyle(_options.Font.WithSize(size), 72D);
            double width = Math.Max(size * 0.2D, _measurer.MeasureWidth(text, style));
            double height = Math.Max(size, _measurer.MeasureLineHeight(style));
            // Positioned drawing text paints its baseline one source em below the frame top.
            double baseline = size;
            if (_options.Fonts.TryResolveFaceForText(text, _options.Font.FamilyName,
                _options.Font.Style, size, out OfficeFontFace? face)) {
                IOfficeFontProgram font = face!.Program;
                width = Math.Max(0.01D, font.Measure(text, size));
                double fontBaseline = font is IOfficeFontBaselineMetrics metrics
                    ? metrics.BaselineOffset(size) : font.LineHeight(size) * 0.8D;
                var contours = font is IOfficeBoundedFontProgram bounded
                    ? bounded.GetTextContoursBounded(text, 0D, -fontBaseline, size, 1_000_000, _cancellationToken)
                    : font.GetTextContours(text, 0D, -fontBaseline, size);
                double top = double.PositiveInfinity, bottom = double.NegativeInfinity;
                foreach (var contour in contours) {
                    foreach (OfficePoint point in contour) {
                        _cancellationToken.ThrowIfCancellationRequested();
                        top = Math.Min(top, point.Y);
                        bottom = Math.Max(bottom, point.Y);
                    }
                }
                if (!double.IsInfinity(top) && !double.IsInfinity(bottom) && bottom > top) {
                    baseline = Math.Max(0D, -top);
                    height = Math.Max(0.01D, bottom - Math.Min(0D, top));
                }
            }
            var box = new LayoutBox(width, height, baseline);
            if (!string.IsNullOrEmpty(text)) {
                box.Commands.Add(LayoutCommand.TextCommand(text, 0D, 0D, width, height, size, baseline));
            }
            return box;
        }
    }
}
