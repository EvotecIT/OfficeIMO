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
        if (segment.BidiResolved ||
            !OfficeTextElements.ContainsRightToLeft(segment.Text) && !OfficeTextElements.ContainsBidiControl(segment.Text)) {
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

    private IReadOnlyList<InlineSegment> ResolveInlineLineSegments(
        IReadOnlyList<InlineSegment> segments,
        string paragraphDirection,
        out bool resolved) {
        resolved = false;
        if (segments.Count < 2 || segments.Any(static segment =>
                segment.Run.AtomicBlock != null ||
                segment.Run.RunningStringElement != null ||
                segment.Run.PositionedMarkerElement != null)) {
            return segments;
        }

        string directionalText = string.Concat(segments.Select(static segment => segment.Text));
        if (!OfficeTextElements.ContainsBidiControl(directionalText)) return segments;

        var visibleElements = new List<InlineBidiElement>();
        for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
            CheckCancellation();
            InlineSegment segment = segments[segmentIndex];
            IReadOnlyList<string> paintElements = OfficeTextElements.Split(segment.Text);
            IReadOnlyList<string> logicalElements = OfficeTextElements.Split(segment.LogicalText);
            bool hasContextualWidths = _fonts.TryMeasureTextElements(
                segment.Text,
                paintElements,
                segment.Run.Style.Font.Size,
                segment.Run.Style.Font.FamilyName,
                segment.Run.Style.Font.Style,
                out IReadOnlyList<double> contextualWidths);
            var widths = new double[paintElements.Count];
            double visibleWidth = 0D;
            for (int elementIndex = 0; elementIndex < paintElements.Count; elementIndex++) {
                string paintText = paintElements[elementIndex];
                if (OfficeTextElements.ContainsBidiControl(paintText)) continue;
                double elementWidth = hasContextualWidths
                    ? contextualWidths[elementIndex]
                    : MeasureText(paintText, segment.Run.Style.Font);
                widths[elementIndex] = elementWidth;
                visibleWidth += elementWidth;
            }

            double widthScale = visibleWidth > 0D ? segment.Width / visibleWidth : 1D;
            for (int elementIndex = 0; elementIndex < paintElements.Count; elementIndex++) {
                string paintText = paintElements[elementIndex];
                if (OfficeTextElements.ContainsBidiControl(paintText)) continue;
                string logicalText = elementIndex < logicalElements.Count
                    ? logicalElements[elementIndex]
                    : OfficeArabicTextShaper.ToLogicalText(paintText);
                visibleElements.Add(new InlineBidiElement(
                    paintText,
                    logicalText,
                    segmentIndex,
                    elementIndex,
                    Math.Max(0.01D, widths[elementIndex] * widthScale)));
            }
        }

        OfficeTextDirection baseDirection = string.Equals(paragraphDirection, "rtl", StringComparison.Ordinal)
            ? OfficeTextDirection.RightToLeft
            : OfficeTextDirection.LeftToRight;
        IReadOnlyList<InlineBidiElement> ordered = OfficeBidiTextResolver.ToVisualOrder(
            directionalText,
            visibleElements,
            baseDirection,
            _cancellationToken,
            static element => element.WithText(OfficeBidiTextResolver.MirrorText(element.Text)));
        if (visibleElements.Count > 0 && ordered.Count == 0) return segments;

        var result = new List<InlineSegment>();
        int start = 0;
        while (start < ordered.Count) {
            int end = start + 1;
            while (end < ordered.Count && ordered[end].SourceSegmentIndex == ordered[start].SourceSegmentIndex) end++;
            InlineSegment source = segments[ordered[start].SourceSegmentIndex];
            string text = string.Concat(ordered.Skip(start).Take(end - start).Select(static element => element.Text));
            string logicalText = string.Concat(ordered
                .Skip(start)
                .Take(end - start)
                .OrderBy(static element => element.SourceElementIndex)
                .Select(static element => element.LogicalText));
            double segmentWidth = ordered.Skip(start).Take(end - start).Sum(static element => element.Width);
            result.Add(new InlineSegment(text, segmentWidth, source.Run, logicalText, bidiResolved: true));
            start = end;
        }

        resolved = true;
        return result;
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

    private readonly struct InlineBidiElement {
        internal InlineBidiElement(
            string text,
            string logicalText,
            int sourceSegmentIndex,
            int sourceElementIndex,
            double width) {
            Text = text;
            LogicalText = logicalText;
            SourceSegmentIndex = sourceSegmentIndex;
            SourceElementIndex = sourceElementIndex;
            Width = width;
        }

        internal string Text { get; }
        internal string LogicalText { get; }
        internal int SourceSegmentIndex { get; }
        internal int SourceElementIndex { get; }
        internal double Width { get; }

        internal InlineBidiElement WithText(string text) =>
            new InlineBidiElement(text, LogicalText, SourceSegmentIndex, SourceElementIndex, Width);
    }
}
