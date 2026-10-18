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
        var contentSafeBreakRanges = new List<HtmlRenderBreakRange>();
        var contentLineBreakGroups = new List<HtmlRenderLineBreakGroup>();
        var contentAvoidBreakRanges = new List<HtmlRenderAvoidBreakRange>();
        var contentForcedBreaks = new List<HtmlRenderForcedBreak>();
        var runningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        bool pagedColumnRegion = _columnContentDepth == 0 && CanPaginateColumnRegion(element, style);
        ColumnPaginationContext? columnContext = null;
        double contentHeight = 0D;
        for (int index = 0; index < partitions.Count; index++) {
            MultiColumnSpanPartition partition = partitions[index];
            IReadOnlyList<HtmlRenderFlowBlock> children = BuildMultiColumnChildBlocks(
                element,
                partition.Nodes,
                columnWidth,
                style,
                depth,
                includeGeneratedBefore: index == 0,
                includeGeneratedAfter: index == partitions.Count - 1);
            if (BuildColumnFlowBreaks(children).HasPageBreaks) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
                    "Authored page breaks inside multi-column content retain the source-flow fallback.",
                    HtmlDiagnosticSeverity.Warning, source, "page breaks in column content");
            }
            AppendMultiColumnPartition(
                element,
                children,
                requestedCount,
                columnWidth,
                gap,
                style,
                source,
                contentHeight,
                contentWidth,
                partition.Spanner == null,
                pagedColumnRegion,
                contentVisuals,
                contentBreakOffsets,
                contentSafeBreakRanges,
                contentLineBreakGroups,
                contentAvoidBreakRanges,
                contentForcedBreaks,
                runningStringAssignments,
                out double partitionHeight,
                out ColumnPaginationContext? partitionContext);
            columnContext ??= partitionContext;
            contentHeight += partitionHeight;

            if (partition.Spanner == null || partition.SpannerStyle == null) continue;
            RecordNormalFlowPlacement(partition.Spanner, element, 0D, contentHeight, partition.SpannerStyle);
            HtmlRenderFlowBlock spanner = LayoutColumnSpanner(partition.Spanner, contentWidth, partition.SpannerStyle, style, depth + 1);
            if (IsForcedPageBreak(spanner.BreakBefore) || IsForcedPageBreak(spanner.BreakAfter)
                || spanner.ForcedBreaks.Any(item => IsForcedPageBreak(item.Target))
                || partition.SpannerStyle.AvoidBreakBeforeScope != HtmlFragmentationScope.None
                || partition.SpannerStyle.AvoidBreakAfterScope != HtmlFragmentationScope.None) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
                    "Authored fragmentation controls on a column spanner retain the source-flow fallback.",
                    HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(partition.Spanner),
                    "fragmentation controls on column spanner");
            }
            foreach (HtmlRenderVisual visual in spanner.Visuals) {
                contentVisuals.Add(visual.Translate(0D, contentHeight, contentVisuals.Count));
            }
            foreach (double offset in spanner.BreakOffsets) {
                contentBreakOffsets.Add(contentHeight + offset);
            }
            contentSafeBreakRanges.AddRange(HtmlRenderBreakRange.Normalize(spanner.SafeBreakRanges, spanner.Height)
                .Select(range => range.Translate(contentHeight)));
            contentLineBreakGroups.AddRange(SliceColumnLineBreakGroups(
                spanner.LineBreakGroups, 0D, spanner.Height, contentHeight));
            contentAvoidBreakRanges.AddRange(SliceColumnAvoidBreakRanges(spanner, 0D, spanner.Height, contentHeight));
            foreach (HtmlCssRunningStringAssignment assignment in spanner.RunningStringAssignments) {
                runningStringAssignments.Add(assignment.Translate(contentHeight));
            }
            contentHeight += spanner.Height;
        }

        double boxHeight = ResolveBoxHeight(contentHeight, boxWidth, style);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        var positionedRunningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.Negative,
            visuals,
            positionedRunningStringAssignments);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        double safeContentHeight = columnContext != null
            ? Math.Max(contentHeight, boxHeight - style.VerticalInsets) : contentHeight;
        if (safeContentHeight > contentHeight)
            contentSafeBreakRanges.Add(new HtmlRenderBreakRange(contentHeight, safeContentHeight));
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
            visuals,
            positionedRunningStringAssignments);
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
            lineBreakGroups: contentLineBreakGroups.Select(group => group.Translate(contentY)),
            pageName: style.PageName,
            forcedBreaks: contentForcedBreaks.Select(item => item.Translate(contentY)),
            avoidBreakRanges: contentAvoidBreakRanges.Select(range => range.Translate(contentY)),
            runningStringAssignments: runningStringAssignments
                .Select(assignment => assignment.Translate(contentY))
                .Concat(positionedRunningStringAssignments)
                .OrderBy(assignment => assignment.OrderOffset),
            safeBreakRanges: ConstrainParallelPaintBreakRanges(
                ProjectColumnTerminalSafeBreakRanges(
                    contentSafeBreakRanges, safeContentHeight, contentY, outerHeight).ToArray(), visuals));
        block.UnfragmentedHeight = columnContext?.UnfragmentedHeight;
        if (columnContext != null && _columnPaginationAttempt is { } attempt
            && ReferenceEquals(attempt.Element, element)) attempt.Context = columnContext;
        if (pagedColumnRegion) _pagedColumnRegions[element] = block;
        return true;
    }

    /// <summary>Retains loss reporting for column controls inside full-width spanners.</summary>
    private HtmlRenderFlowBlock LayoutColumnSpanner(IElement element, double width,
        HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, int depth) {
        _columnContentDepth++;
        _columnBreakFallbackDepth++;
        try {
            return LayoutElement(element, width, style, parentStyle, depth);
        } finally {
            _columnBreakFallbackDepth--;
            _columnContentDepth--;
        }
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
        IElement owner,
        IReadOnlyList<HtmlRenderFlowBlock> children,
        int requestedCount,
        double columnWidth,
        double gap,
        HtmlRenderBoxStyle style,
        string source,
        double offsetY,
        double contentWidth,
        bool finalPartition,
        bool pagedColumnRegion,
        ICollection<HtmlRenderVisual> visuals,
        ISet<double> breakOffsets,
        ICollection<HtmlRenderBreakRange> safeBreakRanges,
        ICollection<HtmlRenderLineBreakGroup> lineBreakGroups,
        ICollection<HtmlRenderAvoidBreakRange> avoidBreakRanges,
        ICollection<HtmlRenderForcedBreak> forcedBreaks,
        ICollection<HtmlCssRunningStringAssignment> runningStringAssignments,
        out double partitionHeight,
        out ColumnPaginationContext? columnContext) {
        columnContext = null;
        if (children.Count == 0) {
            partitionHeight = 0D;
            return;
        }

        double targetHeight = ResolveSpanningColumnHeight(children, requestedCount, style, offsetY, finalPartition);
        ColumnPageBands? pageBands = finalPartition && pagedColumnRegion
            ? ResolveSpanningColumnPageBands(owner, contentWidth, style, requestedCount, targetHeight, offsetY, out columnContext)
            : null;
        MultiColumnPlan plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount,
            throwOnLimit: true, pageBands: pageBands);
        if (pageBands is { Supported: false }) {
            pageBands = null;
            columnContext = null;
            plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount, throwOnLimit: true);
        }
        EnsureMultiColumnLimit(plan.ColumnCount);
        if (_options.Mode == HtmlRenderMode.Paged && plan.ColumnCount > requestedCount
            && (!finalPartition || style.ColumnFill != "auto")) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
                "Balanced spanning-set overflow across physical pages retains the source-flow fallback.",
                HtmlDiagnosticSeverity.Warning, source, "balanced spanning-set overflow");
        }
        double minimumExtent = pageBands?.MinimumBoxExtent(plan.ColumnCount) ?? 0D;
        IReadOnlyList<double> pageBreaks = Array.Empty<double>();
        double? declaredHeight = ResolveDeclaredColumnContentHeight(style);
        bool shortDefiniteColumns = declaredHeight.HasValue && pageBands == null
            && declaredHeight.Value < _activePageGeometry.ContentHeight - 0.0001D;
        if (finalPartition && style.OverflowX == "visible" && style.OverflowY == "visible" && !shortDefiniteColumns)
            plan = ResolvePagedColumnOverflow(plan, requestedCount, targetHeight, out pageBreaks, pageBands);
        foreach (double offset in pageBreaks) forcedBreaks.Add(new HtmlRenderForcedBreak(offsetY + offset, HtmlPageBreakTarget.Page));
        partitionHeight = Math.Max(minimumExtent, Math.Max(targetHeight, plan.UsedHeight));
        if (pageBands != null) {
            // A single sequential column can shrink to its painted extent. An
            // advance to another column consumes the complete physical band.
            partitionHeight = Math.Max(minimumExtent, plan.UsedHeight);
        }
        AddColumnRuleVisuals(visuals, style, 0D, offsetY, columnWidth, gap, Math.Max(requestedCount, plan.ColumnCount), partitionHeight, source);
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            double x = fragment.Column * (columnWidth + gap);
            double y = offsetY + fragment.Y;
            IReadOnlyList<HtmlRenderVisual> fragmentVisuals = SliceColumnFragmentVisuals(fragment);
            foreach (HtmlRenderVisual visual in fragmentVisuals) {
                visuals.Add(visual.Translate(x, y, visuals.Count));
            }
        }

        foreach (double offset in ResolveMultiColumnBreakOffsets(plan, 0D, partitionHeight)) {
            breakOffsets.Add(offsetY + offset);
        }
        foreach (HtmlRenderBreakRange range in ResolveMultiColumnSafeBreakRanges(plan)) {
            safeBreakRanges.Add(range.Translate(offsetY));
        }
        if (pageBands != null && partitionHeight > plan.UsedHeight)
            safeBreakRanges.Add(new HtmlRenderBreakRange(offsetY + plan.UsedHeight, offsetY + partitionHeight));
        foreach (HtmlRenderLineBreakGroup group in EnumerateMultiColumnLineBreakGroups(plan, offsetY)) {
            lineBreakGroups.Add(group);
        }
        foreach (HtmlRenderAvoidBreakRange range in EnumerateMultiColumnAvoidBreakRanges(plan, offsetY)) {
            avoidBreakRanges.Add(range);
        }
        foreach (HtmlCssRunningStringAssignment assignment in EnumerateMultiColumnRunningStringAssignments(plan, offsetY)) {
            runningStringAssignments.Add(assignment);
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
