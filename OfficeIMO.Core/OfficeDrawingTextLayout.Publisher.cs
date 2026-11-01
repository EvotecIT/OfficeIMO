using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    private static OfficeRichTextBlockLayout ApplyPublisherLineGeometry(OfficeRichTextBlockLayout layout,
        IReadOnlyList<OfficeRichTextRun> runs, OfficeRichTextParagraph paragraph,
        Func<string?, double, string?, OfficeFontStyle, OfficePublisherFontMetrics?> measure, ParagraphLabelLayout? label = null) {
        var lines = new List<OfficeRichTextLine>(layout.Lines.Count);
        double height = 0D, maximumAdvance = 1D;
        foreach (OfficeRichTextLine line in layout.Lines) {
            bool empty = true;
            foreach (OfficeRichTextSegment segment in line.Segments) empty &= segment.Text.Length == 0;
            OfficeRichTextLine measured = line;
            if (lines.Count == 0 && label != null) {
                var combined = new List<OfficeRichTextSegment>(line.Segments);
                combined.AddRange(label.Line.Segments);
                measured = line.WithPlacement(combined, line.LineHeight, line.OffsetX);
            }
            OfficeRichTextLine resolved = ResolvePublisherLineGeometry(measured, runs, paragraph, measure, empty);
            resolved = resolved.WithPlacement(line.Segments, resolved.LineHeight, line.OffsetX);
            lines.Add(resolved);
            height += resolved.LineHeight;
            maximumAdvance = Math.Max(maximumAdvance, resolved.LineHeight);
        }
        if (lines.Count > 0) height -= lines[lines.Count - 1].LineHeight -
            (lines[lines.Count - 1].ContentHeight ?? lines[lines.Count - 1].LineHeight);
        return new OfficeRichTextBlockLayout(lines, maximumAdvance, layout.Width, height, layout.Clipped) {
            OnlyUnwrappedWidthOverflow = layout.OnlyUnwrappedWidthOverflow
        };
    }

    private static OfficeRichTextLine ResolvePublisherLineGeometry(OfficeRichTextLine line,
        IReadOnlyList<OfficeRichTextRun> runs, OfficeRichTextParagraph paragraph,
        Func<string?, double, string?, OfficeFontStyle, OfficePublisherFontMetrics?> measure, bool includeTypingRuns) {
        double ascent = 0D, descent = 0D, baseline = 0D;
        bool resolved = false;
        foreach (OfficeRichTextSegment segment in line.Segments) {
            // Advance-only tab/label separators have no font strut of their own.
            if (segment.Text.Length == 0) continue;
            Include(segment.Text, segment.FontSize, segment.FontFamily, segment.FontStyle, segment.Baseline);
        }
        if (includeTypingRuns) {
            foreach (OfficeRichTextRun run in runs)
                Include(string.Empty, run.FontSize, run.FontFamily,
                    (run.Bold ? OfficeFontStyle.Bold : 0) | (run.Italic ? OfficeFontStyle.Italic : 0), run.Baseline);
        }
        if (!resolved) return line;
        double content = ascent + descent;
        double advance = content * (paragraph.LineHeightFactor ?? 1.2D);
        var result = line.WithPlacement(line.Segments, advance, line.OffsetX);
        result.ContentHeight = content; result.BaselineOffset = baseline;
        return result;

        void Include(string text, double size, string? family, OfficeFontStyle style, OfficeTextBaseline position) {
            double effectiveSize = position == OfficeTextBaseline.Normal ? size : size * .65D;
            OfficePublisherFontMetrics? selected = measure(text, effectiveSize, family, style);
            OfficePublisherFontMetrics metrics = selected ?? new OfficePublisherFontMetrics(effectiveSize * .84D, effectiveSize * .16D, effectiveSize * .84D);
            double shift = position == OfficeTextBaseline.Superscript ? -size * .30D :
                position == OfficeTextBaseline.Subscript ? size * .15D : 0D;
            ascent = Math.Max(ascent, metrics.Ascent - shift);
            descent = Math.Max(descent, metrics.Descent + shift);
            baseline = Math.Max(baseline, metrics.Baseline - shift);
            resolved |= selected.HasValue;
        }
    }
}
