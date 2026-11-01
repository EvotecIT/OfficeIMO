using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    private static OfficeRichTextLine TrimRichTextLineToWidthWithEllipsis(OfficeRichTextLine line, double maxWidth, OfficeRichTextMeasurement measure) {
        if (line.Segments.Count == 0) {
            return line;
        }

        var segments = new List<OfficeRichTextSegment>(line.Segments.Count);
        for (int i = 0; i < line.Segments.Count; i++) {
            segments.Add(line.Segments[i]);
        }

        OfficeRichTextSegment ellipsisStyle = segments[segments.Count - 1];
        double width = NormalizeNonNegative(maxWidth);
        while (segments.Count > 0) {
            OfficeRichTextLine candidate = CreateRichTextLineWithEllipsis(segments, ellipsisStyle, measure, line.LineHeight);
            if (candidate.Width <= width) {
                return new OfficeRichTextLine(candidate.Segments, candidate.LineHeight, line.OffsetX);
            }

            int last = segments.Count - 1;
            OfficeRichTextSegment segment = segments[last];
            string text = OfficeTextElements.RemoveLast(segment.Text);
            if (text.Length == 0) {
                segments.RemoveAt(last);
            } else {
                segments[last] = CopyRichTextSegment(segment, text, measure.Measure(text, segment));
            }
        }

        const string ellipsis = "...";
        return measure.Measure(ellipsis, ellipsisStyle) <= width
            ? new OfficeRichTextLine(new[] { CreateRichTextSegment(ellipsis, ellipsisStyle, measure) }, line.LineHeight, line.OffsetX)
            : new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), line.LineHeight, line.OffsetX);
    }

    private static OfficeRichTextLine CreateRichTextLineWithEllipsis(List<OfficeRichTextSegment> segments, OfficeRichTextSegment ellipsisStyle, OfficeRichTextMeasurement measure, double lineHeight) {
        var measured = new List<OfficeRichTextSegment>(segments.Count);
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            string text = i == segments.Count - 1 ? segment.Text + "..." : segment.Text;
            measured.Add(CopyRichTextSegment(segment, text, measure.Measure(text, segment)));
        }

        if (measured.Count == 0) {
            const string ellipsis = "...";
            measured.Add(CreateRichTextSegment(ellipsis, ellipsisStyle, measure));
        }

        return new OfficeRichTextLine(measured, lineHeight);
    }

    private static OfficeRichTextLine CreateRichTextLine(List<OfficeRichTextSegment> segments, OfficeRichTextMeasurement measure) {
        var measured = new List<OfficeRichTextSegment>(segments.Count);
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            measured.Add(CopyRichTextSegment(segment, segment.Text, measure.Measure(segment.Text, segment)));
        }

        return new OfficeRichTextLine(measured);
    }

    private static OfficeRichTextSegment CreateRichTextSegment(string text, OfficeRichTextSegment style, OfficeRichTextMeasurement measure) =>
        CopyRichTextSegment(style, text, measure.Measure(text, style));

    private static OfficeRichTextSegment CopyRichTextSegment(OfficeRichTextSegment style, string text, double width) =>
        style.WithLayout(text, width, text == style.Text ? style.PreparedShaping : null);

    private static double EffectiveFontSize(OfficeRichTextSegment segment) =>
        segment.Baseline == OfficeTextBaseline.Normal ? segment.FontSize : segment.FontSize * 0.65D;

    private static bool ClipRichTextLinesToHeight(
        List<OfficeRichTextLine> lines,
        double maxHeight,
        double maxWidth,
        OfficeRichTextMeasurement measure,
        OfficeTextOverflowBehavior overflowBehavior) {
        if (lines.Count == 0) {
            return false;
        }

        double height = NormalizeNonNegative(maxHeight);
        double used = 0D;
        int visibleCount = 0;
        for (int i = 0; i < lines.Count; i++) {
            double lineHeight = Math.Max(1D, lines[i].LineHeight);
            if (visibleCount > 0 && used + lineHeight > height + 0.01D) {
                break;
            }

            used += lineHeight;
            visibleCount++;
        }

        visibleCount = Math.Max(1, visibleCount);
        if (visibleCount >= lines.Count) {
            return false;
        }

        lines.RemoveRange(visibleCount, lines.Count - visibleCount);
        if (overflowBehavior == OfficeTextOverflowBehavior.Ellipsis) {
            OfficeRichTextLine last = lines[lines.Count - 1];
            lines[lines.Count - 1] = TrimRichTextLineToWidthWithEllipsis(last, Math.Max(0D, maxWidth - last.OffsetX), measure);
        }

        return true;
    }

}
