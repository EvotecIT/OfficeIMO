using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    private sealed class ParagraphLabelLayout {
        internal ParagraphLabelLayout(OfficeRichTextLine line, double left, double height, bool separator) { Line = line; Left = left; Height = height; Separator = separator; }
        internal OfficeRichTextLine Line { get; }
        internal double Left { get; }
        internal double Height { get; }
        internal bool Separator { get; }
    }

    private static ParagraphLabelLayout? ResolveParagraphLabel(OfficeRichTextParagraph paragraph, double contentWidth, double scale, double fontScale,
        Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        CancellationToken cancellationToken, ref OfficeTextPadding margin, ref OfficeTextParagraphIndent indent, ref bool clipped) {
        OfficeTextParagraphLabel? label = paragraph.Label;
        if (label == null) return null;
        OfficeRichTextRun original = label.Run;
        double size = fontScale == 1D ? original.FontSize * scale : Math.Max(1D, original.FontSize * scale * fontScale);
        var run = new OfficeRichTextRun(original.Text, size, original.Color, original.Bold, original.Italic,
            original.Underline, original.FontFamily, original.Strikethrough, original.BackgroundColor, original.UnderlineStyle,
            original.StrikethroughStyle, original.Baseline) { LinkUri = original.LinkUri };
        var layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(new[] { run }, double.MaxValue, double.MaxValue,
            paragraph.LineHeightFactor ?? 1.2, measure, false, measurePaint: measurePaint, cancellationToken: cancellationToken);
        OfficeRichTextLine line = layout.Lines.Count == 0 ? new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>()) : layout.Lines[0];
        double position = label.Position * scale, width = line.Width;
        double alignment = label.Alignment == OfficeTextAlignment.Right ? 1 : label.Alignment == OfficeTextAlignment.Center ? .5 : 0;
        double left = label.MinimumWidth.HasValue ? position + Math.Max(0, label.MinimumWidth.Value * scale - width) * alignment : position - width * alignment;
        double normalFirst = margin.Left + indent.FirstLineOffset, continuation = margin.Left + indent.ContinuationLineOffset;
        double distance = label.MinimumDistance * scale;
        if (label.MinimumWidth.HasValue) {
            // A minimum gap first moves the label inside its box, then moves the text.
            double required = left + width + distance - (label.TextPosition ?? normalFirst / scale) * scale;
            if (required > 0) left -= Math.Min(left - position, required);
        }
        if (left < 0) { left = 0; clipped = true; }
        // A partial number can identify a different item. Omit an overflowing label as a whole,
        // keeping the original body insets, rather than painting beyond the content rectangle.
        if (left + width > contentWidth - margin.Right + .000001D) { clipped = true; return null; }
        double first;
        if (label.FollowedBy == OfficeTextParagraphLabelFollowedBy.Position)
            first = Math.Max((label.TextPosition ?? normalFirst / scale) * scale, left + width + distance);
        else first = left + width + (label.FollowedBy == OfficeTextParagraphLabelFollowedBy.Space ?
            measure(" ", run.FontSize, run.FontFamily, (run.Bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular) | (run.Italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular)) : 0);
        double basis = Math.Min(first, continuation);
        margin = new OfficeTextPadding(basis, margin.Top, margin.Right, margin.Bottom);
        indent = new OfficeTextParagraphIndent(first - basis, continuation - basis);
        return new ParagraphLabelLayout(line, left, OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, layout.LineHeight),
            label.FollowedBy != OfficeTextParagraphLabelFollowedBy.Nothing);
    }

    private static OfficeRichTextLine AddParagraphLabel(OfficeRichTextLine body, ParagraphLabelLayout label, double height) {
        var segments = new List<OfficeRichTextSegment>(label.Line.Segments.Count + body.Segments.Count + 1);
        segments.AddRange(label.Line.Segments);
        double gap = Math.Max(0, body.OffsetX - label.Left - label.Line.Width);
        // Alignment may leave physical space even when the label has no semantic separator.
        if (label.Separator || gap > 0)
            segments.Add(new OfficeRichTextSegment(label.Separator ? " " : string.Empty, gap, Math.Max(1, label.Line.FontSize), OfficeColor.Black, false, false, false, "Arial", false, null));
        segments.AddRange(body.Segments);
        return new OfficeRichTextLine(segments, height, label.Left);
    }
}
