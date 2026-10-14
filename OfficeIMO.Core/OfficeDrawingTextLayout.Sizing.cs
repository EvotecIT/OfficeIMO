using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    // Reflow can expose different glyph bearings. Bound repeated measurement;
    // any remaining paint overflow is still visible to projection loss checks.
    private const int MaximumHorizontalPaintPasses = 8;

    /// <summary>
    /// Measures the complete unwrapped paragraph body's natural width, including
    /// paragraph insets, labels, tabs and horizontal ink/decorations. Alignment
    /// uses that finite intrinsic width, never an effectively infinite canvas.
    /// </summary>
    internal static double RequiredParagraphFrameWidth(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        OfficeDrawingTextMetrics metrics, CancellationToken cancellationToken) {
        var natural = MeasureNaturalHorizontalPaint(paragraphs, metrics, 1D, cancellationToken);
        return natural.Width + natural.Left + natural.Right;
    }

    private static (double Width, double Left, double Right) MeasureNaturalHorizontalPaint(
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, OfficeDrawingTextMetrics metrics,
        double scale, CancellationToken cancellationToken = default) {
        OfficeRichTextBlockLayout natural = CreateParagraphsWithArea(paragraphs, double.MaxValue, double.MaxValue,
            metrics.MeasureText, OfficeTextAreaAlignment.Left, wrap: false, scale: scale,
            measurePaint: metrics.MeasurePaintBounds, cancellationToken: cancellationToken,
            measureSegmentPaint: metrics.MeasureSegmentPaintBounds);
        cancellationToken.ThrowIfCancellationRequested();
        var paint = HorizontalPaintEnvelope(natural, metrics.MeasureHorizontalPaintBounds);
        return (natural.Width, Math.Max(0D, -paint.Left), Math.Max(0D, paint.Right - natural.Width));
    }

    // Reserve bearing/decorative insets before alignment and wrapping. A later
    // global translation would move intentional no-wrap overhang away from its
    // center/right anchor and cannot contain right-aligned or justified ink.
    private static OfficeRichTextBlockLayout CreateWithHorizontalPaintInsets(OfficeDrawingRichText text,
        double width, double height, OfficeDrawingTextMetrics metrics, double scale, bool projection = false) {
        if (!text.NormalizeHorizontalPaint || text.Paragraphs.Count == 0) return Layout(width, height);
        var natural = MeasureNaturalHorizontalPaint(text.Paragraphs, metrics, scale);
        double leftInset = natural.Left, rightInset = natural.Right;
        if (text.WrapText) {
            for (int pass = 0; pass < MaximumHorizontalPaintPasses; pass++) {
                double probeWidth = width - leftInset - rightInset;
                if (probeWidth <= 0D) break;
                // Probe the complete wrapped body, including lines hidden by a
                // fixed height. These same insets drive the height-growth pass.
                var paint = HorizontalPaintEnvelope(Layout(probeWidth, double.MaxValue), metrics.MeasureHorizontalPaintBounds);
                double nextLeft = Math.Max(leftInset, -paint.Left);
                double nextRight = Math.Max(rightInset, paint.Right - probeWidth);
                if (nextLeft <= leftInset + .000001D && nextRight <= rightInset + .000001D) break;
                leftInset = nextLeft; rightInset = nextRight;
            }
        }
        double innerWidth = width - leftInset - rightInset;
        // A cap smaller than the ink insets cannot contain this body. Retain the
        // ordinary anchor rather than inventing a positive semantic rectangle.
        if (innerWidth <= 0D) return Layout(width, height);
        OfficeRichTextBlockLayout layout = Layout(innerWidth, height);
        if (leftInset == 0D && rightInset == 0D) return layout;
        var lines = new List<OfficeRichTextLine>(layout.Lines.Count);
        foreach (OfficeRichTextLine line in layout.Lines)
            lines.Add(new OfficeRichTextLine(line.Segments, line.LineHeight, line.OffsetX + leftInset));
        return new OfficeRichTextBlockLayout(lines, layout.LineHeight, layout.Width + leftInset + rightInset,
            layout.Height, layout.Clipped) {
            ContentOffsetY = layout.ContentOffsetY, OnlyUnwrappedWidthOverflow = layout.OnlyUnwrappedWidthOverflow
        };

        OfficeRichTextBlockLayout Layout(double availableWidth, double availableHeight) => projection
            ? CreateForProjectionCore(text, availableWidth, availableHeight, metrics.MeasureText, scale,
                metrics.MeasurePaintBounds, metrics.MeasureSegmentPaintBounds)
            : CreateCore(text, availableWidth, availableHeight, metrics.MeasureText, scale,
                metrics.MeasurePaintBounds, metrics.MeasureSegmentPaintBounds);
    }

    internal static OfficeRichTextBlockLayout CreateWithRasterMetrics(OfficeDrawingRichText text,
        double width, double height, OfficeRasterCanvas metrics, double scale = 1D) => text.NormalizeHorizontalPaint
        ? CreateWithHorizontalPaintInsets(text, width, height,
            new OfficeDrawingTextMetrics(metrics.MeasureText, metrics.MeasureTextPaintBounds,
                (value, size, family, style) => metrics.MeasureTextLineHorizontalPaintBounds(value ?? string.Empty, size, family, style)), scale)
        : Create(text, width, height, metrics.MeasureText, scale, metrics.MeasureTextPaintBounds);
}
