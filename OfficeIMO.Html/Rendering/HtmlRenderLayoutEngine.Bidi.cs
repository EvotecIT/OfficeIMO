using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static string ResolveLogicalText(IEnumerable<HtmlRenderVisual> visuals, string fallback) {
        IReadOnlyList<HtmlRenderVisual> materialized = visuals as IReadOnlyList<HtmlRenderVisual> ?? visuals.ToArray();
        if (HtmlRenderLogicalText.TryResolveReorderedText(materialized, out string reorderedText)) return reorderedText;
        var text = new StringBuilder();
        foreach (HtmlRenderVisual visual in materialized.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) continue;
            if (visual is HtmlRenderText renderedText) text.Append(renderedText.Text);
            else if (visual is HtmlRenderLogicalTextGroup logicalText) text.Append(logicalText.Text);
            else if (visual is HtmlRenderClipGroup clip) text.Append(ResolveLogicalText(clip.Visuals, string.Empty));
            else if (visual is HtmlRenderPathClipGroup pathClip) text.Append(ResolveLogicalText(pathClip.Visuals, string.Empty));
            else if (visual is HtmlRenderEffectGroup effect) text.Append(ResolveLogicalText(effect.Visuals, string.Empty));
            else if (visual is HtmlRenderSemanticGroup semantic) text.Append(ResolveLogicalText(semantic.Visuals, string.Empty));
            else if (visual is HtmlRenderLayoutRegion region) text.Append(ResolveLogicalText(region.Visuals, string.Empty));
        }
        return text.Length == 0 ? fallback : text.ToString();
    }

    private IReadOnlyList<InlinePaintSegment> ResolveInlinePaintSegments(InlineSegment segment, double x) {
        if (segment.PreparedShaping != null) {
            return new[] { new InlinePaintSegment(segment.LogicalText, x, Math.Max(0.01D, segment.Width),
                segment.Width, 0, bidiResolved: true, preparedShaping: segment.PreparedShaping) };
        }
        bool ordinaryDirection = !OfficeTextElements.ContainsRightToLeft(segment.Text)
            && !OfficeTextElements.ContainsBidiControl(segment.Text)
            && segment.ResolvedDirection != OfficeTextDirection.RightToLeft;
        if (!segment.Run.PreserveContextualFallback && (segment.BidiResolved || ordinaryDirection)
            && (!segment.Run.Style.TextFeatureSettings.IsDefault
                || RequiresConfiguredTextShaping(segment.LogicalText, segment.Run.Style.TextFeatureSettings))
            && TryPrepareDirectionalProviderText(segment.LogicalText, segment.Run.Style,
                segment.ResolvedDirection ?? OfficeTextDirection.LeftToRight, segment.Width, out OfficePreparedTextShaping? retained)) {
            return new[] { new InlinePaintSegment(retained!.Text, x, Math.Max(0.01D, segment.Width),
                segment.Width, 0, bidiResolved: true, preparedShaping: retained) };
        }
        if (!segment.BidiResolved
            && (Math.Abs(segment.Run.Style.LetterSpacing) > 0.0001D || Math.Abs(segment.Run.Style.WordSpacing) > 0.0001D)) {
            if (!OfficeTextElements.ContainsRightToLeft(segment.Text) && !OfficeTextElements.ContainsBidiControl(segment.Text)
                && segment.ResolvedDirection != OfficeTextDirection.RightToLeft) {
                return ResolveSpacedPaintSegments(segment, x);
            }
        }
        if (segment.BidiResolved
            && (Math.Abs(segment.Run.Style.LetterSpacing) > 0.0001D || Math.Abs(segment.Run.Style.WordSpacing) > 0.0001D)) {
            return ResolveSpacedPaintSegments(segment, x);
        }
        if (segment.BidiResolved ||
            !OfficeTextElements.ContainsRightToLeft(segment.Text) && !OfficeTextElements.ContainsBidiControl(segment.Text)
                && segment.ResolvedDirection != OfficeTextDirection.RightToLeft) {
            ReportUnretainedProviderPaint(segment.Run, segment.LogicalText);
            return new[] { new InlinePaintSegment(segment.Text, x, Math.Max(0.01D, segment.Width), segment.Width, 0) };
        }

        var result = new List<InlinePaintSegment>();
        IReadOnlyList<InlineDirectionalGroup> groups = ResolveDirectionalGroups(segment);
        double cursor = x;
        foreach (InlineDirectionalGroup group in groups) {
            double groupX = cursor;
            if (group.RightToLeft) {
                if (!segment.Run.PreserveContextualFallback && TryPrepareDirectionalProviderText(group.Text, segment.Run.Style,
                    OfficeTextDirection.RightToLeft, group.Width, out OfficePreparedTextShaping? prepared)) {
                    result.Add(new InlinePaintSegment(
                        prepared!.Text,
                        groupX,
                        Math.Max(0.01D, group.Width),
                        group.Width,
                        group.LogicalOrder, bidiResolved: true, preparedShaping: prepared));
                } else if (!segment.Run.PreserveContextualFallback && TryMeasureWithConfiguredProvider(group.Text, segment.Run.Style, out _)) {
                    if (!segment.Run.Style.TextFeatureSettings.IsDefault || !OfficeTextElements.ContainsRightToLeft(group.Text)
                        || RequiresConfiguredTextShaping(group.Text, segment.Run.Style.TextFeatureSettings))
                        ReportProviderContextLoss(segment.Run, OfficeTextElements.ContainsRightToLeft(group.Text)
                            ? "directional-glyph-context-not-retained" : "neutral-directional-glyph-context-not-retained");
                    // A result for a fresh paragraph is not a result for this
                    // already resolved directional slice. Preserve its physical
                    // scalar order when the complete glyph context was declined.
                    AppendRightToLeftPaintSegments(result, group, groupX, segment.Run.Style);
                } else {
                    AppendRightToLeftPaintSegments(result, group, groupX, segment.Run.Style);
                }
            } else if (!segment.Run.PreserveContextualFallback && TryPrepareDirectionalProviderText(group.Text, segment.Run.Style,
                OfficeTextDirection.LeftToRight, group.Width, out OfficePreparedTextShaping? leftToRight)) {
                result.Add(new InlinePaintSegment(leftToRight!.Text, groupX, Math.Max(0.01D, group.Width),
                    group.Width, group.LogicalOrder, bidiResolved: true, preparedShaping: leftToRight));
            } else if (Math.Abs(segment.Run.Style.LetterSpacing) > 0.0001D || Math.Abs(segment.Run.Style.WordSpacing) > 0.0001D) {
                result.AddRange(ResolveSpacedPaintSegments(
                    new InlineSegment(group.Text, group.Width, segment.Run, group.Text),
                    groupX,
                    group.LogicalOrder,
                    bidiResolved: true));
            } else {
                ReportUnretainedProviderPaint(segment.Run, group.Text);
                result.Add(new InlinePaintSegment(group.Text, groupX, Math.Max(0.01D, group.Width), group.Width,
                    group.LogicalOrder, bidiResolved: true));
            }
            cursor += group.Width;
        }
        return result.OrderBy(static item => item.LogicalOrder).ToArray();
    }

    private IReadOnlyList<InlinePaintSegment> ResolveSpacedPaintSegments(InlineSegment segment, double x,
        int? logicalOrder = null, bool bidiResolved = false) {
        var result = new List<InlinePaintSegment>();
        if (RequiresConfiguredTextShaping(segment.LogicalText, segment.Run.Style.TextFeatureSettings))
            ReportProviderContextLoss(segment.Run, "spaced-provider-glyph-context-not-retained");
        double cursor = x;
        IReadOnlyList<string> elements = OfficeTextElements.Split(segment.Text);
        OfficeFontInfo effectiveFont = GetEffectiveTextFont(segment.Run.Style);
        if (Math.Abs(segment.Run.Style.LetterSpacing) > 0.0001D) {
            for (int index = 0; index < elements.Count; index++) {
                string element = elements[index];
                double glyphWidth = MeasureText(element, effectiveFont, segment.Run.Style.FontDescriptor);
                double advance = glyphWidth
                    + segment.Run.Style.LetterSpacing
                    + (IsWordSpacingSeparator(element) ? segment.Run.Style.WordSpacing : 0D);
                result.Add(new InlinePaintSegment(element, cursor, Math.Max(0.01D, glyphWidth), advance,
                    logicalOrder ?? index, bidiResolved));
                cursor += advance;
            }
            return result;
        }

        int start = 0;
        while (start < elements.Count) {
            bool whitespace = IsWordSpacingSeparator(elements[start]);
            int end = start + 1;
            while (end < elements.Count && IsWordSpacingSeparator(elements[end]) == whitespace) end++;
            string text = string.Concat(elements.Skip(start).Take(end - start));
            double glyphWidth = MeasureText(text, effectiveFont, segment.Run.Style.FontDescriptor);
            double advance = glyphWidth + (whitespace ? segment.Run.Style.WordSpacing * (end - start) : 0D);
            result.Add(new InlinePaintSegment(text, cursor, Math.Max(0.01D, glyphWidth), advance,
                logicalOrder ?? start, bidiResolved));
            cursor += advance;
            start = end;
        }
        return result;
    }

    private IReadOnlyList<InlineDirectionalGroup> ResolveDirectionalGroups(InlineSegment segment) {
        if (segment.ResolvedDirection.HasValue) {
            return new[] { new InlineDirectionalGroup(segment.Text,
                segment.ResolvedDirection == OfficeTextDirection.RightToLeft, segment.Width, 0,
                segment.ContextualFallbackPaint) };
        }
        var groups = new List<InlineDirectionalGroup>();
        OfficeTextDirection baseDirection = string.Equals(segment.Run.Style.Direction, "rtl", StringComparison.Ordinal)
            ? OfficeTextDirection.RightToLeft
            : OfficeTextDirection.LeftToRight;
        foreach (OfficeBidiTextRun run in OfficeBidiTextResolver.ResolveVisualRuns(segment.Text, baseDirection)) {
            groups.Add(new InlineDirectionalGroup(
                run.Text,
                run.Direction == OfficeTextDirection.RightToLeft,
                MeasureInlineText(run.Text, segment.Run.Style, segment.Run.PreserveContextualFallback),
                run.LogicalOrder));
        }
        return groups;
    }

    private IReadOnlyList<IReadOnlyList<InlineSegment>>? ResolveInlineParagraphSegments(
        IReadOnlyList<IReadOnlyList<InlineSegment>> lines,
        string paragraphDirection) {
        if (lines.Count == 1 && lines[0].Count < 2) return null;
        if (lines.SelectMany(static line => line).Any(static segment =>
                segment.Run.AtomicBlock != null ||
                segment.Run.IsEmptyInlineBox ||
                segment.Run.RunningStringElement != null ||
                segment.Run.PositionedMarkerElement != null)) {
            return null;
        }

        string directionalText = string.Concat(lines.SelectMany(static line => line).Select(static segment =>
            segment.Text.Length == 0 && segment.LogicalText == "\t" ? "\t" : segment.Text));
        // Ordinary RTL words participate in one paragraph even when inline styles or
        // font fallback split them into separate source runs without bidi controls.
        if (!OfficeTextElements.ContainsBidiControl(directionalText)
            && !OfficeTextElements.ContainsRightToLeft(directionalText)) return null;

        var visibleLines = new List<IReadOnlyList<InlineBidiElement>>(lines.Count);
        foreach (IReadOnlyList<InlineSegment> line in lines) {
            visibleLines.Add(BuildInlineBidiElements(line));
        }

        OfficeTextDirection baseDirection = string.Equals(paragraphDirection, "rtl", StringComparison.Ordinal)
            ? OfficeTextDirection.RightToLeft
            : OfficeTextDirection.LeftToRight;
        IReadOnlyList<IReadOnlyList<InlineBidiElement>> orderedLines = OfficeBidiTextResolver.ToVisualLineOrder(
            directionalText,
            visibleLines,
            baseDirection,
            _cancellationToken,
            static element => element.WithText(OfficeBidiTextResolver.MirrorText(element.Text)));
        if (orderedLines.Count != lines.Count) return null;

        var result = new List<IReadOnlyList<InlineSegment>>(lines.Count);
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            result.Add(BuildResolvedInlineSegments(lines[lineIndex], orderedLines[lineIndex]));
        }
        return result.AsReadOnly();
    }

    private List<InlineBidiElement> BuildInlineBidiElements(IReadOnlyList<InlineSegment> segments) {
        var visibleElements = new List<InlineBidiElement>();
        for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
            CheckCancellation();
            InlineSegment segment = segments[segmentIndex];
            if (segment.Text.Length == 0 && segment.LogicalText == "\t") {
                // A preserved tab is a width-bearing source segment even though it paints no glyph.
                visibleElements.Add(new InlineBidiElement(string.Empty, "\t", segmentIndex, 0, segment.Width));
                continue;
            }
            IReadOnlyList<string> paintElements = OfficeTextElements.Split(segment.Text);
            IReadOnlyList<string> logicalElements = OfficeTextElements.Split(segment.LogicalText);
            bool hasPreparedWidths = TryPartitionPreparedBidiWidths(segment, paintElements, logicalElements,
                out IReadOnlyList<double> preparedWidths);
            OfficeFontInfo effectiveFont = GetEffectiveTextFont(segment.Run.Style);
            IReadOnlyList<double> contextualWidths = Array.Empty<double>();
            bool hasContextualWidths = !hasPreparedWidths && _fonts.TryMeasureTextElements(
                segment.Text,
                paintElements,
                effectiveFont.Size,
                effectiveFont.FamilyName,
                effectiveFont.Style,
                out contextualWidths);
            var widths = new double[paintElements.Count];
            double visibleWidth = 0D;
            for (int elementIndex = 0; elementIndex < paintElements.Count; elementIndex++) {
                string paintText = paintElements[elementIndex];
                if (OfficeTextElements.ContainsBidiControl(paintText)) continue;
                double elementWidth = hasPreparedWidths ? preparedWidths[elementIndex] : hasContextualWidths
                    ? contextualWidths[elementIndex]
                    : MeasureText(paintText, effectiveFont, segment.Run.Style.FontDescriptor);
                widths[elementIndex] = elementWidth;
                visibleWidth += elementWidth;
            }

            double widthScale = !hasPreparedWidths && visibleWidth > 0D ? segment.Width / visibleWidth : 1D;
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
                    hasPreparedWidths ? widths[elementIndex] : Math.Max(0.01D, widths[elementIndex] * widthScale)));
            }
        }
        return visibleElements;
    }

    private IReadOnlyList<InlineSegment> BuildResolvedInlineSegments(
        IReadOnlyList<InlineSegment> sources,
        IReadOnlyList<InlineBidiElement> ordered) {
        var result = new List<InlineSegment>();
        var firstGroups = new InlineSegment?[sources.Count];
        var lastGroups = new InlineSegment?[sources.Count];
        int start = 0;
        while (start < ordered.Count) {
            int sourceIndex = ordered[start].SourceSegmentIndex;
            HtmlRenderBoxStyle style = sources[sourceIndex].Run.Style;
            bool providerContext = _options.TextShapingProvider != null || !style.TextFeatureSettings.IsDefault
                || OfficeManagedTextShaper.RequiresComplexLayout(sources[sourceIndex].LogicalText);
            int end = start + 1;
            while (end < ordered.Count && ordered[end].SourceSegmentIndex == sourceIndex
                && (!providerContext || ordered[end].RightToLeft == ordered[start].RightToLeft)) end++;
            string text = string.Concat(ordered.Skip(start).Take(end - start).Select(static element => element.Text));
            InlineBidiElement[] logicalOrder = ordered
                .Skip(start)
                .Take(end - start)
                .OrderBy(static element => element.SourceElementIndex)
                .ToArray();
            string logicalText = string.Concat(logicalOrder.Select(static element => element.LogicalText));
            double segmentWidth = ordered.Skip(start).Take(end - start).Sum(static element => element.Width);
            OfficePreparedTextShaping? retainedSlice = SlicePreparedBidiGroup(sources[sourceIndex], logicalOrder, logicalText);
            // Preserve prepared contextual paint when no provider owns this script.
            // Marks and required ligatures add fidelity loss without discarding those forms.
            // Unsupported script input and configured shapers still need logical source.
            bool needsShapingInput = providerContext && ordered[start].RightToLeft
                || RequiresConfiguredTextShaping(logicalText, style.TextFeatureSettings)
                || OfficeTextElements.ContainsShapingRequiredScript(logicalText)
                || OfficeTextElements.ContainsVariationSelector(logicalText)
                || OfficeTextElements.ContainsJoiningScript(logicalText)
                    && !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(logicalText)
                || OfficeTextElements.ContainsZeroWidthJoinerSequence(logicalText)
                || OfficeTextElements.ContainsRightToLeft(logicalText)
                    && (_options.TextShapingProvider != null || !style.TextFeatureSettings.IsDefault);
            string? contextualFallback = needsShapingInput && retainedSlice == null
                && OfficeTextElements.ContainsJoiningScript(logicalText)
                ? string.Concat(logicalOrder.Select(static element => element.Text)) : null;
            var group = new InlineSegment(
                retainedSlice != null || needsShapingInput ? logicalText : text,
                segmentWidth,
                sources[sourceIndex].Run,
                logicalText,
                bidiResolved: retainedSlice != null || !needsShapingInput,
                logicalEndProgress: sources[sourceIndex].LogicalEndProgress,
                resolvedDirection: !providerContext ? null
                    : ordered[start].RightToLeft ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight,
                contextualFallbackPaint: contextualFallback,
                preparedShaping: retainedSlice,
                prepareFromRun: false);
            result.Add(group);
            firstGroups[sourceIndex] ??= group;
            lastGroups[sourceIndex] = group;
            start = end;
        }
        // Reordering may split a source segment into several visual groups.
        // Its box edges belong to the two visual ends, never to every group.
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            if (firstGroups[sourceIndex] is not InlineSegment first || lastGroups[sourceIndex] is not InlineSegment last) continue;
            first.LeadingAdvance += sources[sourceIndex].LeadingAdvance;
            last.TrailingAdvance += sources[sourceIndex].TrailingAdvance;
        }
        return result.AsReadOnly();
    }

    private string ResolveRootInlineLineLogicalText(
        IReadOnlyList<InlineSegment> segments,
        AngleSharp.Dom.IElement? formattingContainer) {
        var text = new StringBuilder();
        foreach (InlineSegment segment in segments) {
            if (FindNearestInlineStackingElement(segment.Run.OwnerElement, formattingContainer) != null) continue;
            foreach (string element in OfficeTextElements.Enumerate(segment.LogicalText)) {
                if (!OfficeTextElements.ContainsBidiControl(element)) text.Append(element);
            }
        }
        return text.ToString();
    }

    private void AppendRightToLeftPaintSegments(List<InlinePaintSegment> result, InlineDirectionalGroup group, double x, HtmlRenderBoxStyle style) {
        string paint = group.ContextualFallbackPaint ?? group.Text;
        IReadOnlyList<string> elements = OfficeTextElements.Enumerate(paint).ToList();
        OfficeFontInfo effectiveFont = GetEffectiveTextFont(style);
        bool hasContextualWidths = _fonts.TryMeasureTextElements(
            paint,
            elements,
            effectiveFont.Size,
            effectiveFont.FamilyName,
            effectiveFont.Style,
            out IReadOnlyList<double> contextualWidths);
        double right = x + group.Width;
        for (int index = 0; index < elements.Count; index++) {
            CheckCancellation();
            string element = elements[index];
            double advance = (hasContextualWidths ? contextualWidths[index] : MeasureText(element, effectiveFont, style.FontDescriptor))
                + style.LetterSpacing
                + (IsWordSpacingSeparator(element) ? style.WordSpacing : 0D);
            right -= advance;
            double glyphWidth = hasContextualWidths ? contextualWidths[index] : MeasureText(element, effectiveFont, style.FontDescriptor);
            result.Add(new InlinePaintSegment(
                OfficeBidiTextResolver.MirrorText(element),
                right,
                Math.Max(0.01D, glyphWidth),
                advance,
                group.LogicalOrder,
                bidiResolved: true));
        }
    }

    private readonly struct InlinePaintSegment {
        internal InlinePaintSegment(string text, double x, double width, double advance, int logicalOrder,
            bool bidiResolved = false, OfficePreparedTextShaping? preparedShaping = null) {
            Text = text;
            X = x;
            Width = width;
            Advance = advance;
            LogicalOrder = logicalOrder;
            BidiResolved = bidiResolved;
            PreparedShaping = preparedShaping;
        }

        internal bool BidiResolved { get; }
        internal OfficePreparedTextShaping? PreparedShaping { get; }
        internal string Text { get; }
        internal double X { get; }
        internal double Width { get; }
        internal double Advance { get; }
        internal int LogicalOrder { get; }
    }

    private readonly struct InlineDirectionalGroup {
        internal InlineDirectionalGroup(string text, bool rightToLeft, double width, int logicalOrder,
            string? contextualFallbackPaint = null) {
            Text = text;
            RightToLeft = rightToLeft;
            Width = width;
            LogicalOrder = logicalOrder;
            ContextualFallbackPaint = contextualFallbackPaint;
        }

        internal string Text { get; }
        internal bool RightToLeft { get; }
        internal double Width { get; }
        internal int LogicalOrder { get; }
        internal string? ContextualFallbackPaint { get; }
    }

    private readonly struct InlineBidiElement {
        internal InlineBidiElement(
            string text,
            string logicalText,
            int sourceSegmentIndex,
            int sourceElementIndex,
            double width, bool rightToLeft = false) {
            Text = text;
            LogicalText = logicalText;
            SourceSegmentIndex = sourceSegmentIndex;
            SourceElementIndex = sourceElementIndex;
            Width = width;
            RightToLeft = rightToLeft;
        }

        internal string Text { get; }
        internal string LogicalText { get; }
        internal int SourceSegmentIndex { get; }
        internal int SourceElementIndex { get; }
        internal double Width { get; }
        internal bool RightToLeft { get; }

        internal InlineBidiElement WithText(string text) =>
            new InlineBidiElement(text, LogicalText, SourceSegmentIndex, SourceElementIndex, Width, rightToLeft: true);
    }
}
