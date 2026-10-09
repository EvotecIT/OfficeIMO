using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    // Paragraph alignment/justification is already materialized in line offsets
    // and segment advances. Preserve separate paint rectangles until ancestor clips
    // and transforms have been applied; a hidden wide line must not size a surface.
    internal static IEnumerable<(double Left, double Top, double Right, double Bottom)> PlacedRasterParagraphPaintBounds(
        OfficeRichTextBlockLayout layout, OfficeRasterCanvas canvas, double left, double top,
        double height, OfficeTextVerticalAlignment verticalAlignment) {
        double lineTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment) + layout.ContentOffsetY;
        foreach (OfficeRichTextLine line in layout.Lines) {
            canvas.CancellationToken.ThrowIfCancellationRequested();
            double lineHeight = OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, layout.LineHeight);
            double baseline = OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, lineTop, lineHeight, true);
            double cursor = left + line.OffsetX;
            foreach (OfficeRichTextSegment segment in line.Segments) {
                double size = OfficeTextBlockRenderer.ResolveRichTextRenderedFontSize(segment);
                double segmentTop = OfficeTextBlockRenderer.ResolveRichTextRenderedBaseline(segment, baseline) - size * .84D;
                if (segment.TabLinePaint is { } leader && leader.Color.A > 0) {
                    double leaderBaseline = OfficeTextBlockRenderer.ResolveRichTextRenderedBaseline(segment, baseline);
                    yield return (cursor + leader.Left, leaderBaseline + leader.Top, cursor + leader.Right, leaderBaseline + leader.Bottom);
                }
                if (segment.BackgroundColor.HasValue && segment.BackgroundColor.Value.A > 0 && segment.Width > 0D && segment.FontSize > 0D)
                    yield return (cursor, segmentTop, cursor + segment.Width,
                        segmentTop + OfficeTextBlockRenderer.ResolveRichTextSegmentBackgroundHeight(segment));
                if (segment.Color.A > 0) {
                    foreach (var ink in canvas.TextLinePaintBounds(segment.Text, size, segment.FontFamily, segment.FontStyle,
                        segment.UnderlineStyle, segment.StrikethroughStyle))
                        yield return (cursor + ink.Left, segmentTop + ink.Top, cursor + ink.Right, segmentTop + ink.Bottom);
                }
                cursor += segment.Width;
            }
            lineTop += lineHeight;
        }
    }
}
