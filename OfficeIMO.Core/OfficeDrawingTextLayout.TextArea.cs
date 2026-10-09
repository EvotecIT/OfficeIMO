using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    // Intrinsic sizing precedes vertical clipping: hidden wider paragraphs still anchor the whole area.
    // Retained measurements keep the authored frame's wrapping, tab advances and font-fit result stable.
    private static OfficeRichTextBlockLayout CreateIntrinsicParagraphs(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double width, double height, Func<string?, double, string?, OfficeFontStyle, double> measure,
        bool wrap, double scale, double fontScale,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        CancellationToken cancellationToken, bool reportLeaderPaintClipping, OfficeTextAreaAlignment areaAlignment,
        bool reportUnwrappedWidthOverflow) {
        var measuredParagraphs = new List<(OfficeRichTextParagraph Paragraph, MeasuredParagraph Measured)>(paragraphs.Count);
        var leaderBudget = new OfficeTextTabLeaderLayout.Budget(reportLeaderPaintClipping);
        double areaWidth = 0D;
        bool clipped = false;
        foreach (OfficeRichTextParagraph paragraph in paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            MeasuredParagraph? measured = MeasureParagraph(paragraph, width, measure, wrap, scale, fontScale,
                measurePaint, cancellationToken, leaderBudget, ref clipped, reportUnwrappedWidthOverflow);
            if (measured == null) continue;
            measuredParagraphs.Add((paragraph, measured));
            for (int i = 0; i < measured.LineCount; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeRichTextLine natural = PlaceParagraphLine(paragraph, measured, measured.Line(i), i, 0D,
                    width, wrap, OfficeTextAlignment.Left, measure);
                areaWidth = Math.Max(areaWidth, natural.OffsetX + natural.Width + measured.Margins.Right);
            }
        }

        double areaLeft = (width - areaWidth) * (areaAlignment == OfficeTextAreaAlignment.Center ? .5D :
            areaAlignment == OfficeTextAreaAlignment.Right ? 1D : 0D);
        var lines = new List<OfficeRichTextLine>();
        double usedHeight = 0D, maximumLineHeight = 1D;
        foreach (var entry in measuredParagraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRichTextParagraph paragraph = entry.Paragraph;
            MeasuredParagraph measured = entry.Measured;
            if (!AddGap(measured.Margins.Top)) { clipped = true; break; }
            bool paragraphTruncated = false;
            for (int i = 0; i < measured.LineCount; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeRichTextLine line = measured.Line(i);
                double lineHeight = paragraph.LineHeight * scale ??
                    OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, measured.Layout.LineHeight);
                if (i == 0 && measured.Label != null && !paragraph.LineHeight.HasValue)
                    lineHeight = Math.Max(lineHeight, measured.Label.Height);
                if (lines.Count >= OfficeTextLayoutEngine.MaximumLayoutLines || usedHeight + lineHeight > height + .000001D) {
                    clipped = true; paragraphTruncated = true; break;
                }
                OfficeRichTextLine placed = PlaceParagraphLine(paragraph, measured, line, i, lineHeight, areaWidth,
                    wrap, paragraph.Alignment, measure);
                lines.Add(new OfficeRichTextLine(placed.Segments, lineHeight, placed.OffsetX + areaLeft));
                usedHeight += lineHeight;
                maximumLineHeight = Math.Max(maximumLineHeight, lineHeight);
            }
            if (paragraphTruncated || !AddGap(measured.Margins.Bottom)) { clipped = true; break; }
        }
        return new OfficeRichTextBlockLayout(lines, maximumLineHeight, areaLeft + areaWidth, usedHeight, clipped);

        bool AddGap(double gap) {
            if (gap == 0D) return true;
            if (usedHeight + gap > height || lines.Count >= OfficeTextLayoutEngine.MaximumLayoutLines) return false;
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), gap));
            usedHeight += gap; return true;
        }
    }
}
