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
        if (!OfficeTextElements.ContainsRightToLeft(segment.Text)) {
            return new[] { new InlinePaintSegment(segment.Text, x, segment.Width) };
        }

        var result = new List<InlinePaintSegment>();
        IReadOnlyList<InlineDirectionalGroup> groups = ResolveDirectionalGroups(segment);
        bool baseRightToLeft = string.Equals(segment.Run.Style.Direction, "rtl", StringComparison.Ordinal);
        double cursor = baseRightToLeft ? x + segment.Width : x;
        foreach (InlineDirectionalGroup group in groups) {
            double groupX = baseRightToLeft ? cursor - group.Width : cursor;
            if (group.RightToLeft) {
                AppendRightToLeftPaintSegments(result, group, groupX, segment.Run.Style.Font);
            } else {
                result.Add(new InlinePaintSegment(group.Text, groupX, Math.Max(0.01D, group.Width)));
            }
            cursor += baseRightToLeft ? -group.Width : group.Width;
        }
        return result;
    }

    private IReadOnlyList<InlineDirectionalGroup> ResolveDirectionalGroups(InlineSegment segment) {
        var groups = new List<InlineDirectionalGroup>();
        var text = new StringBuilder();
        bool? rightToLeft = null;
        foreach (string element in OfficeTextElements.Enumerate(segment.Text)) {
            bool elementRightToLeft = OfficeTextElements.ContainsRightToLeft(element);
            bool neutral = element.All(character => char.IsWhiteSpace(character) || char.IsPunctuation(character));
            bool resolvedDirection = neutral && rightToLeft.HasValue ? rightToLeft.Value : elementRightToLeft;
            if (rightToLeft.HasValue && rightToLeft.Value != resolvedDirection) {
                string value = text.ToString();
                groups.Add(new InlineDirectionalGroup(value, rightToLeft.Value, MeasureText(value, segment.Run.Style.Font)));
                text.Clear();
            }
            rightToLeft = resolvedDirection;
            text.Append(element);
        }
        if (text.Length > 0) {
            string value = text.ToString();
            groups.Add(new InlineDirectionalGroup(value, rightToLeft == true, MeasureText(value, segment.Run.Style.Font)));
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
            result.Add(new InlinePaintSegment(element, right, Math.Max(0.01D, advance)));
        }
    }

    private readonly struct InlinePaintSegment {
        internal InlinePaintSegment(string text, double x, double width) {
            Text = text;
            X = x;
            Width = width;
        }

        internal string Text { get; }
        internal double X { get; }
        internal double Width { get; }
    }

    private readonly struct InlineDirectionalGroup {
        internal InlineDirectionalGroup(string text, bool rightToLeft, double width) {
            Text = text;
            RightToLeft = rightToLeft;
            Width = width;
        }

        internal string Text { get; }
        internal bool RightToLeft { get; }
        internal double Width { get; }
    }
}
