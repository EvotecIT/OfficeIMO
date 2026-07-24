using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryLayoutSpanningMultiColumnContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        double boxWidth,
        double contentWidth,
        double columnWidth,
        double gap,
        int requestedCount,
        string source,
        out HtmlRenderFlowBlock block) {
        block = null!;
        IReadOnlyList<MultiColumnSpanPartition> partitions = ResolveMultiColumnSpanPartitions(element, contentWidth, style);
        if (partitions.Count <= 1) return false;

        var contentVisuals = new List<HtmlRenderVisual>();
        var contentBreakOffsets = new SortedSet<double>();
        double contentHeight = 0D;
        for (int index = 0; index < partitions.Count; index++) {
            MultiColumnSpanPartition partition = partitions[index];
            IReadOnlyList<HtmlRenderFlowBlock> children = BuildChildBlocks(
                element,
                partition.Nodes,
                columnWidth,
                style,
                depth,
                includeGeneratedBefore: index == 0,
                includeGeneratedAfter: index == partitions.Count - 1);
            AppendMultiColumnPartition(
                children,
                requestedCount,
                columnWidth,
                gap,
                style,
                source,
                contentHeight,
                contentVisuals,
                contentBreakOffsets,
                out double partitionHeight);
            contentHeight += partitionHeight;

            if (partition.Spanner == null || partition.SpannerStyle == null) continue;
            RecordNormalFlowPlacement(partition.Spanner, element, 0D, contentHeight, partition.SpannerStyle);
            HtmlRenderFlowBlock spanner = LayoutElement(partition.Spanner, contentWidth, partition.SpannerStyle, style, depth + 1);
            foreach (HtmlRenderVisual visual in spanner.Visuals) {
                contentVisuals.Add(visual.Translate(0D, contentHeight, contentVisuals.Count));
            }
            foreach (double offset in spanner.BreakOffsets) {
                contentBreakOffsets.Add(contentHeight + offset);
            }
            contentHeight += spanner.Height;
        }

        double boxHeight = ResolveBoxHeight(contentHeight, style);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.Negative,
            visuals);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in contentVisuals) {
            visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        }
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.NonNegative,
            visuals);
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);

        IEnumerable<double> breakOffsets = contentBreakOffsets
            .Where(offset => offset > 0.0001D && offset < contentHeight - 0.0001D)
            .Select(offset => contentY + offset)
            .Concat(new[] { outerHeight });
        block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            breakOffsets,
            pageName: style.PageName);
        return true;
    }

    private IReadOnlyList<MultiColumnSpanPartition> ResolveMultiColumnSpanPartitions(
        IElement element,
        double contentWidth,
        HtmlRenderBoxStyle parentStyle) {
        var partitions = new List<MultiColumnSpanPartition>();
        var nodes = new List<INode>();
        foreach (INode node in element.ChildNodes) {
            if (node is not IElement child || ShouldSkipElement(child)) {
                nodes.Add(node);
                continue;
            }

            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, contentWidth, parentStyle);
            bool isSpanner = childStyle.Display != "none"
                && childStyle.ColumnSpan == "all"
                && childStyle.FloatSide == "none"
                && !ShouldExtractOutOfFlow(childStyle)
                && HtmlRenderStyleResolver.IsBlockElement(child, childStyle);
            if (!isSpanner) {
                nodes.Add(node);
                continue;
            }

            partitions.Add(new MultiColumnSpanPartition(nodes, child, childStyle));
            nodes = new List<INode>();
        }

        if (partitions.Count == 0) return Array.Empty<MultiColumnSpanPartition>();
        partitions.Add(new MultiColumnSpanPartition(nodes, null, null));
        return partitions;
    }

    private void AppendMultiColumnPartition(
        IReadOnlyList<HtmlRenderFlowBlock> children,
        int requestedCount,
        double columnWidth,
        double gap,
        HtmlRenderBoxStyle style,
        string source,
        double offsetY,
        ICollection<HtmlRenderVisual> visuals,
        ISet<double> breakOffsets,
        out double partitionHeight) {
        if (children.Count == 0) {
            partitionHeight = 0D;
            return;
        }

        double targetHeight = Math.Max(0.01D, ResolveBalancedColumnHeight(children, requestedCount));
        MultiColumnPlan plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount, throwOnLimit: true);
        EnsureMultiColumnLimit(plan.ColumnCount);
        partitionHeight = Math.Max(targetHeight, plan.UsedHeight);
        AddColumnRuleVisuals(visuals, style, 0D, offsetY, columnWidth, gap, Math.Max(requestedCount, plan.ColumnCount), partitionHeight, source);
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            double x = fragment.Column * (columnWidth + gap);
            double y = offsetY + fragment.Y;
            IReadOnlyList<HtmlRenderVisual> fragmentVisuals = SliceBlockVisuals(fragment.Block, fragment.Start, fragment.End);
            foreach (HtmlRenderVisual visual in fragmentVisuals) {
                visuals.Add(visual.Translate(x, y, visuals.Count));
            }
        }

        foreach (double offset in ResolveMultiColumnBreakOffsets(plan, 0D, partitionHeight)) {
            breakOffsets.Add(offsetY + offset);
        }
    }

    private sealed class MultiColumnSpanPartition {
        internal MultiColumnSpanPartition(IEnumerable<INode> nodes, IElement? spanner, HtmlRenderBoxStyle? spannerStyle) {
            Nodes = new List<INode>(nodes).AsReadOnly();
            Spanner = spanner;
            SpannerStyle = spannerStyle;
        }

        internal IReadOnlyList<INode> Nodes { get; }
        internal IElement? Spanner { get; }
        internal HtmlRenderBoxStyle? SpannerStyle { get; }
    }
}
