using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool IsVerticalWritingMode(string writingMode) =>
        writingMode == "vertical-rl" || writingMode == "vertical-lr"
        || writingMode == "sideways-rl" || writingMode == "sideways-lr";

    private double ResolveVerticalInlineExtent(
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        double fallback) {
        if (style.ExplicitHeight.HasValue) {
            double height = style.ExplicitHeight.Value;
            return Math.Max(1D, style.BorderBox ? height - style.VerticalInsets : height);
        }

        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        if (containingHeight.HasValue) return Math.Max(1D, containingHeight.Value - style.VerticalInsets);
        double surfaceHeight = _options.Mode == HtmlRenderMode.Paged
            ? _activePageGeometry.ContentHeight
            : (_options.ViewportHeight ?? fallback) - _options.Margins.Top - _options.Margins.Bottom;
        return Math.Max(1D, surfaceHeight - style.VerticalInsets);
    }

    private static void ArrangeVerticalBlockChildren(
        IReadOnlyList<HtmlRenderFlowBlock> children,
        HtmlRenderBoxStyle style,
        double contentWidth,
        ICollection<FlowPaintLayer> paintLayers,
        ICollection<double> breakOffsets,
        ICollection<HtmlRenderForcedBreak> forcedBreaks,
        ICollection<HtmlRenderLineBreakGroup> lineBreakGroups,
        ICollection<HtmlRenderContinuationGroup> continuationGroups,
        ICollection<HtmlRenderTrailingGroup> trailingGroups,
        ICollection<HtmlCssRunningStringAssignment> runningStringAssignments,
        ICollection<HtmlInlineBreakProgress> continuationBreakProgress,
        out double contentHeight) {
        bool rightToLeft = style.WritingMode == "vertical-rl" || style.WritingMode == "sideways-rl";
        double cursor = rightToLeft ? contentWidth : 0D;
        contentHeight = 0D;
        string? childPageName = children.Count > 0 ? children[0].PageName : null;
        for (int childIndex = 0; childIndex < children.Count; childIndex++) {
            HtmlRenderFlowBlock child = children[childIndex];
            double advance = ResolveVerticalBlockAdvance(child, style.LineHeight);
            double childX = rightToLeft ? cursor - advance : cursor;
            if (rightToLeft) cursor = childX;
            else cursor += advance;

            if (childIndex > 0 && !string.Equals(childPageName, child.PageName, StringComparison.Ordinal)) {
                forcedBreaks.Add(new HtmlRenderForcedBreak(0D, HtmlPageBreakTarget.Page, child.PageName, changesPageName: true));
            }
            childPageName = child.PageName;
            if (child.BreakBefore != HtmlPageBreakTarget.None) forcedBreaks.Add(new HtmlRenderForcedBreak(0D, child.BreakBefore));
            foreach (HtmlRenderForcedBreak forcedBreak in child.ForcedBreaks) forcedBreaks.Add(forcedBreak);
            if (childIndex > 0 && child.OwnerElement != null) {
                continuationBreakProgress.Add(new HtmlInlineBreakProgress(0D, 0, child.OwnerElement));
            }

            paintLayers.Add(new FlowPaintLayer(child, childX, 0D, paintLayers.Count));
            contentHeight = Math.Max(contentHeight, child.Height);
            if (child.BreakAfter != HtmlPageBreakTarget.None) forcedBreaks.Add(new HtmlRenderForcedBreak(contentHeight, child.BreakAfter));
            foreach (double offset in child.BreakOffsets) breakOffsets.Add(offset);
            foreach (HtmlRenderLineBreakGroup group in child.LineBreakGroups) lineBreakGroups.Add(group);
            foreach (HtmlRenderContinuationGroup group in child.ContinuationGroups) continuationGroups.Add(group.Translate(childX, 0D));
            foreach (HtmlRenderTrailingGroup group in child.TrailingGroups) trailingGroups.Add(group.Translate(childX, 0D));
            foreach (HtmlCssRunningStringAssignment assignment in child.RunningStringAssignments) runningStringAssignments.Add(assignment);
            foreach (HtmlInlineBreakProgress progress in child.InlineBreakProgress) continuationBreakProgress.Add(progress);
        }
        if (contentHeight > 0D) breakOffsets.Add(contentHeight);
    }

    private static double ResolveVerticalBlockAdvance(HtmlRenderFlowBlock child, double fallback) {
        double maximum = 0D;
        foreach (HtmlRenderVisual visual in child.Visuals) maximum = Math.Max(maximum, ResolveVerticalWritingWidth(visual));
        return Math.Max(0.01D, maximum > 0D ? maximum : Math.Min(child.Width, Math.Max(fallback, child.Height)));
    }

    private static double ResolveVerticalWritingWidth(HtmlRenderVisual visual) {
        double width = visual.Source?.EndsWith(":vertical-writing", StringComparison.Ordinal) == true
            ? visual.Width
            : 0D;
        IReadOnlyList<HtmlRenderVisual>? children = visual switch {
            HtmlRenderEffectGroup effect => effect.Visuals,
            HtmlRenderSemanticGroup semantic => semantic.Visuals,
            HtmlRenderLayoutRegion region => region.Visuals,
            HtmlRenderLogicalTextGroup logical => logical.Visuals,
            HtmlRenderClipGroup clip => clip.Visuals,
            HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
            HtmlRenderFormField field => field.Visuals,
            _ => null
        };
        if (children == null) return width;
        foreach (HtmlRenderVisual child in children) width = Math.Max(width, ResolveVerticalWritingWidth(child));
        return width;
    }

    private HtmlInlineLayout TransformSidewaysVerticalInlineLayout(
        HtmlInlineLayout inline,
        HtmlRenderBoxStyle style,
        IElement element) {
        if (inline.Visuals.Count == 0) return inline;
        double sourceWidth = Math.Max(0.01D, MaximumVerticalSourceRight(inline.Visuals));
        double sourceHeight = Math.Max(0.01D, inline.Height);
        bool rightToLeftBlocks = style.WritingMode == "vertical-rl" || style.WritingMode == "sideways-rl";
        OfficeTransform transform = rightToLeftBlocks
            ? OfficeTransform.RotateDegrees(90D).Then(OfficeTransform.Translate(sourceHeight, 0D))
            : OfficeTransform.RotateDegrees(-90D).Then(OfficeTransform.Translate(0D, sourceWidth));
        string source = HtmlRenderStyleResolver.DescribeSource(element) + ":vertical-writing";
        bool sidewaysOnly = style.TextOrientation == "sideways"
            || style.WritingMode.StartsWith("sideways", StringComparison.Ordinal);
        IReadOnlyList<HtmlRenderVisual> verticalVisuals = sidewaysOnly
            ? new HtmlRenderVisual[] {
                new HtmlRenderEffectGroup(
                    0D,
                    0D,
                    sourceHeight,
                    sourceWidth,
                    transform,
                    1D,
                    inline.Visuals,
                    0,
                    source)
            }
            : TransformVerticalVisuals(
                inline.Visuals,
                transform,
                style.TextOrientation,
                rightToLeftBlocks,
                sourceHeight,
                sourceWidth,
                source);
        string logicalText = ResolveLogicalText(inline.Visuals, string.Empty);
        HtmlRenderVisual verticalVisual = logicalText.Length == 0
            ? new HtmlRenderEffectGroup(
                0D,
                0D,
                sourceHeight,
                sourceWidth,
                OfficeTransform.Identity,
                1D,
                verticalVisuals,
                0,
                source)
            : new HtmlRenderLogicalTextGroup(
                logicalText,
                0D,
                0D,
                sourceHeight,
                sourceWidth,
                verticalVisuals,
                0,
                source);
        return new HtmlInlineLayout(
            new[] { verticalVisual },
            sourceWidth,
            runningStringAssignments: inline.RunningStringAssignments);
    }

    private IReadOnlyList<HtmlRenderVisual> TransformVerticalVisuals(
        IReadOnlyList<HtmlRenderVisual> visuals,
        OfficeTransform axisTransform,
        string textOrientation,
        bool rightToLeftBlocks,
        double destinationWidth,
        double destinationHeight,
        string source) {
        var result = new List<HtmlRenderVisual>();
        for (int index = 0; index < visuals.Count; index++) {
            HtmlRenderVisual visual = visuals[index];
            if (visual is HtmlRenderText text) {
                AppendVerticalText(result, text, axisTransform, textOrientation, rightToLeftBlocks, source);
                continue;
            }
            if (visual is HtmlRenderSemanticGroup semantic) {
                IReadOnlyList<HtmlRenderVisual> children = TransformVerticalVisuals(
                    semantic.Visuals,
                    axisTransform,
                    textOrientation,
                    rightToLeftBlocks,
                    destinationWidth,
                    destinationHeight,
                    source);
                result.Add(new HtmlRenderSemanticGroup(
                    semantic.Role,
                    0D,
                    0D,
                    destinationWidth,
                    destinationHeight,
                    children,
                    result.Count,
                    semantic.Source,
                    semantic.ColumnSpan,
                    semantic.RowSpan,
                    semantic.HeaderScope,
                    structureElementKey: semantic.StructureElementKey));
                continue;
            }
            if (visual is HtmlRenderLogicalTextGroup logical) {
                IReadOnlyList<HtmlRenderVisual> children = TransformVerticalVisuals(
                    logical.Visuals,
                    axisTransform,
                    textOrientation,
                    rightToLeftBlocks,
                    destinationWidth,
                    destinationHeight,
                    source);
                result.Add(new HtmlRenderLogicalTextGroup(
                    logical.Text,
                    0D,
                    0D,
                    destinationWidth,
                    destinationHeight,
                    children,
                    result.Count,
                    logical.Source));
                continue;
            }

            result.Add(new HtmlRenderEffectGroup(
                0D,
                0D,
                destinationWidth,
                destinationHeight,
                axisTransform,
                1D,
                new[] { visual },
                result.Count,
                source));
        }
        return result;
    }

    private void AppendVerticalText(
        List<HtmlRenderVisual> destination,
        HtmlRenderText visual,
        OfficeTransform axisTransform,
        string textOrientation,
        bool rightToLeftBlocks,
        string source) {
        IReadOnlyList<string> elements = OfficeTextElements.Split(visual.Text);
        if (elements.Count == 0) return;

        var measured = new double[elements.Count];
        double measuredTotal = 0D;
        for (int index = 0; index < elements.Count; index++) {
            measured[index] = Math.Max(0.01D, MeasureText(elements[index], visual.Font));
            measuredTotal += measured[index];
        }

        double totalAdvance = Math.Abs(visual.TextAdvanceWidth ?? visual.Width);
        double scale = measuredTotal > 0D ? totalAdvance / measuredTotal : 1D;
        double cursor = visual.X;
        for (int index = 0; index < elements.Count; index++) {
            string textElement = elements[index];
            double glyphWidth = measured[index];
            double advance = measured[index] * scale;
            OfficePoint center = axisTransform.TransformPoint(new OfficePoint(
                cursor + advance / 2D,
                visual.Y + visual.Height / 2D));
            cursor += advance;

            bool upright = textOrientation == "upright" || OfficeTextElements.IsVerticalMixedOrientationUpright(textElement);
            double x = center.X - glyphWidth / 2D;
            double y = center.Y - visual.Height / 2D;
            var glyph = new HtmlRenderText(
                textElement,
                x,
                y,
                glyphWidth,
                visual.Height,
                visual.Font,
                visual.Color,
                visual.Alignment,
                visual.LineHeight,
                0,
                visual.LinkUri,
                visual.Source,
                visual.SemanticRole,
                layoutY: y,
                visual.SemanticNodeId,
                glyphWidth,
                visual.BidiVisualOrderResolved,
                visual.SemanticFragmentOrder,
                visual.LogicalTextOrder,
                visual.UnderlineStyle,
                visual.StrikethroughStyle,
                visual.Baseline,
                visual.BaselineLevel,
                visual.BaselineScale,
                visual.BaselineOffset,
                glyphWidth,
                visual.DecorationColor);
            if (upright) {
                destination.Add(glyph.Translate(0D, 0D, destination.Count));
                continue;
            }

            destination.Add(new HtmlRenderEffectGroup(
                center.X - visual.Height / 2D,
                center.Y - glyphWidth / 2D,
                visual.Height,
                glyphWidth,
                OfficeTransform.RotateDegrees(rightToLeftBlocks ? 90D : -90D, center.X, center.Y),
                1D,
                new[] { glyph },
                destination.Count,
                source));
        }
    }

    private static double MaximumVerticalSourceRight(IEnumerable<HtmlRenderVisual> visuals) {
        double maximum = 0.01D;
        foreach (HtmlRenderVisual visual in visuals) {
            maximum = Math.Max(maximum, visual.X + visual.Width);
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            if (children != null) maximum = Math.Max(maximum, MaximumVerticalSourceRight(children));
        }
        return maximum;
    }
}
