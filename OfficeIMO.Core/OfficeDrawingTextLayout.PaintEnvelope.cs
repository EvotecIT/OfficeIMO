using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    /// <summary>
    /// Checks glyph, decoration, background and tab-leader paint against the content rectangle after the
    /// same vertical alignment and content offset used by drawing renderers.
    /// </summary>
    internal static bool IsVerticalPaintClipped(OfficeRichTextBlockLayout layout, double height,
        OfficeTextVerticalAlignment verticalAlignment,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null) =>
        IsVerticalPaintClipped(layout, height, verticalAlignment, measurePaint, null);

    internal static bool IsVerticalPaintClipped(OfficeRichTextBlockLayout layout, double height,
        OfficeTextVerticalAlignment verticalAlignment,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint, measureSegmentPaint);
        double top = OfficeTextPlacement.ResolveTop(0D, height, layout.Height, verticalAlignment) + layout.ContentOffsetY;
        return top + bounds.Top < -.01D || top + bounds.Bottom > height + .01D;
    }

    // Paragraph alignment, insets, list markers and tabs are already materialized in
    // the layout. Preserve that placement while measuring a separate horizontal canvas.
    internal static (double Left, double Right) HorizontalPaintEnvelope(
        OfficeRichTextBlockLayout layout, OfficeRasterCanvas metrics) =>
        HorizontalPaintEnvelope(layout, (text, size, family, style) =>
            metrics.MeasureTextLineHorizontalPaintBounds(text ?? string.Empty, size, family, style));

    internal static (double Left, double Right) HorizontalPaintEnvelope(OfficeRichTextBlockLayout layout,
        Func<string?, double, string?, OfficeFontStyle, (double Left, double Right)> measurePaint) {
        double left = 0D, right = layout.Width;
        foreach (OfficeRichTextLine line in layout.Lines) {
            double cursor = line.OffsetX;
            foreach (OfficeRichTextSegment segment in line.Segments) {
                double size = OfficeTextBlockRenderer.ResolveRichTextRenderedFontSize(segment);
                var bounds = measurePaint(segment.Text, size, segment.FontFamily, segment.FontStyle);
                left = Math.Min(left, cursor + bounds.Left); right = Math.Max(right, cursor + bounds.Right);
                if (segment.Underline || segment.Strikethrough) {
                    double radius = Math.Max(1D, size / 16D) / 2D;
                    left = Math.Min(left, cursor - radius); right = Math.Max(right, cursor + segment.Width + radius);
                }
                cursor += segment.Width;
            }
        }
        return (left, right);
    }
}
