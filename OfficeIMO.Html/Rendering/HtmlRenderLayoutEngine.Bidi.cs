using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static string ResolveLogicalText(IEnumerable<HtmlRenderVisual> visuals, string fallback) {
        var text = new StringBuilder();
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderText renderedText) text.Append(renderedText.Text);
            else if (visual is HtmlRenderLogicalTextGroup logicalText) text.Append(logicalText.Text);
            else if (visual is HtmlRenderClipGroup clip) text.Append(ResolveLogicalText(clip.Visuals, string.Empty));
            else if (visual is HtmlRenderPathClipGroup pathClip) text.Append(ResolveLogicalText(pathClip.Visuals, string.Empty));
            else if (visual is HtmlRenderEffectGroup effect) text.Append(ResolveLogicalText(effect.Visuals, string.Empty));
            else if (visual is HtmlRenderSemanticGroup semantic) text.Append(ResolveLogicalText(semantic.Visuals, string.Empty));
        }
        return text.Length == 0 ? fallback : text.ToString();
    }

    private IReadOnlyList<InlinePaintSegment> ResolveInlinePaintSegments(InlineSegment segment, double x) {
        if (!OfficeTextElements.ContainsRightToLeft(segment.Text) && !OfficeTextElements.ContainsBidiControl(segment.Text)) {
            return new[] { new InlinePaintSegment(segment.Text, x, segment.Width, 0) };
        }

        var result = new List<InlinePaintSegment>();
        IReadOnlyList<InlineDirectionalGroup> groups = ResolveDirectionalGroups(segment);
        double cursor = x;
        foreach (InlineDirectionalGroup group in groups) {
            double groupX = cursor;
            if (group.RightToLeft) {
                AppendRightToLeftPaintSegments(result, group, groupX, segment.Run.Style.Font);
            } else {
                result.Add(new InlinePaintSegment(group.Text, groupX, Math.Max(0.01D, group.Width), group.LogicalOrder));
            }
            cursor += group.Width;
        }
        return result.OrderBy(static item => item.LogicalOrder).ToArray();
    }

    private IReadOnlyList<InlineDirectionalGroup> ResolveDirectionalGroups(InlineSegment segment) {
        var groups = new List<InlineDirectionalGroup>();
        OfficeTextDirection baseDirection = string.Equals(segment.Run.Style.Direction, "rtl", StringComparison.Ordinal)
            ? OfficeTextDirection.RightToLeft
            : OfficeTextDirection.LeftToRight;
        foreach (OfficeBidiTextRun run in OfficeBidiTextResolver.ResolveVisualRuns(segment.Text, baseDirection)) {
            groups.Add(new InlineDirectionalGroup(
                run.Text,
                run.Direction == OfficeTextDirection.RightToLeft,
                MeasureText(run.Text, segment.Run.Style.Font),
                run.LogicalOrder));
        }
        return groups;
    }

    private void AppendRightToLeftPaintSegments(List<InlinePaintSegment> result, InlineDirectionalGroup group, double x, OfficeFontInfo font) {
        IReadOnlyList<string> elements = OfficeTextElements.Enumerate(group.Text).ToList();
        bool hasContextualWidths = _fonts.TryMeasureTextElements(
            group.Text,
            elements,
            font.Size,
            font.FamilyName,
            font.Style,
            out IReadOnlyList<double> contextualWidths);
        double right = x + group.Width;
        for (int index = 0; index < elements.Count; index++) {
            CheckCancellation();
            string element = elements[index];
            double advance = hasContextualWidths ? contextualWidths[index] : MeasureText(element, font);
            right -= advance;
            result.Add(new InlinePaintSegment(element, right, Math.Max(0.01D, advance), group.LogicalOrder));
        }
    }

    private readonly struct InlinePaintSegment {
        internal InlinePaintSegment(string text, double x, double width, int logicalOrder) {
            Text = text;
            X = x;
            Width = width;
            LogicalOrder = logicalOrder;
        }

        internal string Text { get; }
        internal double X { get; }
        internal double Width { get; }
        internal int LogicalOrder { get; }
    }

    private readonly struct InlineDirectionalGroup {
        internal InlineDirectionalGroup(string text, bool rightToLeft, double width, int logicalOrder) {
            Text = text;
            RightToLeft = rightToLeft;
            Width = width;
            LogicalOrder = logicalOrder;
        }

        internal string Text { get; }
        internal bool RightToLeft { get; }
        internal double Width { get; }
        internal int LogicalOrder { get; }
    }
}
