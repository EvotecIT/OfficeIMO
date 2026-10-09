using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    // Bounds are relative to the already resolved normal/sub/superscript baseline.
    // Empty glyphs can still paint a background or a continuous decoration.
    internal static OfficeTextPaintBounds? MeasureSegmentPaintBounds(OfficeRichTextSegment segment, double size,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint) {
        double top = double.PositiveInfinity, bottom = double.NegativeInfinity;
        if (segment.Color.A > 0 && !string.IsNullOrWhiteSpace(segment.Text)) {
            OfficeTextPaintBounds glyphs = measurePaint?.Invoke(segment.Text, size, segment.FontFamily, segment.FontStyle)
                ?? new OfficeTextPaintBounds(-size * .84D, size * .16D);
            Include(glyphs.Top, glyphs.Bottom);
        }
        if (segment.BackgroundColor.HasValue && segment.BackgroundColor.Value.A > 0 && segment.Width > 0D && segment.FontSize > 0D) {
            double backgroundTop = -size * .84D;
            Include(backgroundTop, backgroundTop + OfficeTextBlockRenderer.ResolveRichTextSegmentBackgroundHeight(segment));
        }
        if (segment.Color.A > 0 && segment.Width > 0D && segment.Text.Length > 0) {
            Decoration(segment.UnderlineStyle, size * .02D);
            Decoration(segment.StrikethroughStyle, -size * .32D);
        }
        return double.IsPositiveInfinity(top) ? null : new OfficeTextPaintBounds(top, bottom);

        void Include(double start, double end) { top = Math.Min(top, start); bottom = Math.Max(bottom, end); }
        void Decoration(OfficeTextDecorationStyle style, double center) {
            if (style == OfficeTextDecorationStyle.None || style == OfficeTextDecorationStyle.Words && string.IsNullOrWhiteSpace(segment.Text)) return;
            var bounds = OfficeRasterCanvas.TextLineDecorationBounds(segment.Width, size, style, center);
            Include(bounds.Top, bounds.Bottom);
        }
    }
}
