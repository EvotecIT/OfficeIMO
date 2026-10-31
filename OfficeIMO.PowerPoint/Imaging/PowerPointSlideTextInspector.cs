using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Reads text placement through the canonical Drawing layout without a PDF round trip.</summary>
internal static class PowerPointSlideTextInspector {
    internal static IReadOnlyList<PowerPointTextSpanInfo> Read(OfficeDrawing drawing) {
        var spans = new List<PowerPointTextSpanInfo>();
        Visit(drawing, OfficeTransform.Identity, drawing.Height, spans);
        return spans.AsReadOnly();
    }

    private static void Visit(OfficeDrawing drawing, OfficeTransform transform, double slideHeight,
        List<PowerPointTextSpanInfo> spans) {
        OfficeRasterCanvas metrics = OfficeDrawingTextLayout.CreateMetrics(drawing);
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingGroup group) {
                OfficeTransform nested = OfficeTransform.Translate(group.ContentOffsetX, group.ContentOffsetY)
                    .Then(OfficeTransform.Translate(group.X, group.Y))
                    .Then(group.FrameTransform?.CreateDestinationTransform() ?? OfficeTransform.Identity).Then(transform);
                Visit(group.InnerDrawing, nested, slideHeight, spans);
            } else if (element is OfficeDrawingRichText rich) {
                ReadRich(rich, metrics, transform, slideHeight, spans);
            } else if (element is OfficeDrawingText text) {
                ReadPlain(text, metrics, transform, slideHeight, spans);
            }
        }
    }

    private static void ReadRich(OfficeDrawingRichText text, OfficeRasterCanvas metrics, OfficeTransform parent,
        double slideHeight, List<PowerPointTextSpanInfo> spans) {
        double width = text.Width - text.Padding.Horizontal, height = text.Height - text.Padding.Vertical;
        if (width <= 0 || height <= 0) return;
        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, width, height, metrics);
        double left = text.X + text.Padding.Left;
        double top = OfficeTextPlacement.ResolveTop(text.Y + text.Padding.Top, height, layout.Height, text.VerticalAlignment) + layout.ContentOffsetY;
        OfficeTransform transform = text.CreateFrameTransform().CreateDestinationTransform().Then(parent);
        foreach (OfficeRichTextLine line in layout.Lines) {
            double lineHeight = OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, layout.LineHeight);
            double baseline = OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, top, lineHeight, true);
            double cursor = OfficeTextPlacement.ResolveLineLeft(left + line.OffsetX,
                Math.Max(0, width - line.OffsetX), line.Width, text.Alignment);
            foreach (OfficeRichTextSegment segment in line.Segments) {
                Add(segment.Text, segment.FontFamily, OfficeTextBlockRenderer.ResolveRichTextRenderedFontSize(segment),
                    cursor, OfficeTextBlockRenderer.ResolveRichTextRenderedBaseline(segment, baseline), segment.Width,
                    segment.Bold, segment.Italic, segment.Color, transform, slideHeight, spans);
                cursor += segment.Width;
            }
            top += lineHeight;
        }
    }

    private static void ReadPlain(OfficeDrawingText text, OfficeRasterCanvas metrics, OfficeTransform parent,
        double slideHeight, List<PowerPointTextSpanInfo> spans) {
        double width = text.Width - text.Padding.Horizontal, height = text.Height - text.Padding.Vertical;
        if (width <= 0 || height <= 0) return;
        OfficeTransform transform = text.CreateFrameTransform().CreateDestinationTransform().Then(parent);
        double left = text.X + text.Padding.Left, top = text.Y + text.Padding.Top;
        if (text.WrapText || text.ShrinkToFit || text.StackedText || text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.HasPadding) {
            int priorLevel = text.BaselineLevel == 0 ? 0 : text.BaselineLevel - Math.Sign(text.BaselineLevel);
            OfficeTextScriptGeometry prior = OfficeTextScriptGeometry.Resolve(Math.Max(.001D, text.Font.Size), priorLevel,
                superscriptOffsetFactor: .35D, subscriptOffsetFactor: .18D);
            var run = new OfficeRichTextRun(text.Text, prior.RenderedFontSize, text.Color ?? OfficeColor.Black,
                text.Font.IsBold, text.Font.IsItalic, text.Font.IsUnderline, text.Font.FamilyName, text.Font.IsStrikethrough,
                underlineStyle: text.UnderlineStyle, strikethroughStyle: text.StrikethroughStyle, baseline: text.Baseline);
            var rich = new OfficeDrawingRichText(new[] { run }, text.X, text.Y, text.Width, text.Height,
                text.Alignment, text.LineHeight, text.VerticalAlignment, text.RotationDegrees, text.RotationCenterX,
                text.RotationCenterY, text.WrapText || text.StackedText, text.ShrinkToFit, text.FlipHorizontal,
                text.FlipVertical, text.Padding, text.ParagraphIndent);
            ReadRich(rich, metrics, OfficeTransform.Translate(0, prior.BaselineOffset).Then(parent), slideHeight, spans);
        } else {
            double advance = text.TextAdvanceWidth ?? metrics.MeasureText(text.RasterText,
                text.Font.Size * text.BaselineScale, text.Font.FamilyName, text.Font.Style);
            double x = OfficeTextPlacement.ResolveLineLeft(left, width, advance, text.Alignment);
            Add(text.Text, text.Font.FamilyName, text.Font.Size * text.BaselineScale, x,
                top + text.Font.Size + text.BaselineOffset, advance, text.Font.IsBold, text.Font.IsItalic,
                text.Color, transform, slideHeight, spans);
        }
    }

    private static void Add(string text, string? family, double size, double x, double y, double advance,
        bool bold, bool italic, OfficeColor? color, OfficeTransform transform, double height,
        List<PowerPointTextSpanInfo> spans) {
        text = text.Trim();
        if (string.IsNullOrEmpty(text) || color?.A == 0) return;
        OfficePoint start = transform.TransformPoint(new OfficePoint(x, y));
        OfficePoint end = transform.TransformPoint(new OfficePoint(x + advance, y));
        double scale = Math.Sqrt(transform.M21 * transform.M21 + transform.M22 * transform.M22);
        spans.Add(new PowerPointTextSpanInfo(text, family, size * scale, start.X, height - start.Y,
            Math.Sqrt(Math.Pow(end.X - start.X, 2) + Math.Pow(end.Y - start.Y, 2)), bold, italic, color,
            -Math.Atan2(transform.M12, transform.M11) * 180D / Math.PI));
    }
}
