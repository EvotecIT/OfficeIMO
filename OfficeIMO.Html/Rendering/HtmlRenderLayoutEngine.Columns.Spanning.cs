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
        IReadOnlyList<MultiColumnSpanPartition> partitions = ResolveMultiColumnSpanPartitions(element, contentWidth, style, columnWidth, depth);
        if (partitions.Count <= 1) return false;
        using IDisposable sourceFlow = EnterColumnSourceFlow();
        PrepareColumnSourcePositionedRequests(partitions);

        var pieces = new List<HtmlRenderFlowBlock>();
        bool outerPageFlow = _columnContentDepth == 0 && CanUseOuterColumnPageFlow(element, style);
        bool pagedColumnRegion = outerPageFlow && style.ColumnFill == "auto";
        bool directSpanners = _columnContentDepth == 0
            && (_options.Mode != HtmlRenderMode.Paged || outerPageFlow);
        bool hasPrecedingSpanner = false;
        ColumnPaginationContext? columnContext = null;
        double contentHeight = 0D;
        for (int index = 0; index < partitions.Count; index++) {
            CheckCancellation();
            ChargeLayoutOperation("spanning column partition");
            MultiColumnSpanPartition partition = partitions[index];
            IReadOnlyList<HtmlRenderFlowBlock> children;
            using (EnterColumnSourcePartition(partition)) {
                children = BuildMultiColumnChildBlocks(
                    element, partition.Nodes, columnWidth, style, depth,
                    includeGeneratedBefore: index == 0,
                    includeGeneratedAfter: index == partitions.Count - 1);
            }
            bool childPageBreaks = BuildColumnFlowBreaks(children).HasPageBreaks;
            if (childPageBreaks) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
                    "Authored page breaks inside multi-column content retain the source-flow fallback.",
                    HtmlDiagnosticSeverity.Warning, source, "page breaks in column content");
            }
            HtmlRenderFlowBlock? prefix = pagedColumnRegion && hasPrecedingSpanner
                && partition.Spanner == null && children.Count > 0 && !childPageBreaks
                ? BuildSpanningFlowBlock(pieces.Concat(new[] { children[0] }).ToArray(), contentWidth, source, true)
                : null;
            HtmlRenderFlowBlock? columnSet = BuildMultiColumnPartition(
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
                directSpanners && _options.Mode == HtmlRenderMode.Paged
                    && partition.SpannerStyle != null,
                prefix,
                out ColumnPaginationContext? partitionContext);
            columnContext ??= partitionContext;
            if (columnSet != null) {
                pieces.Add(columnSet);
                contentHeight += columnSet.Height;
            }

            if (partition.Spanner == null || partition.SpannerStyle == null) continue;
            RegisterColumnSpannerContainingOwner(partition.Spanner, element);
            RecordNormalFlowPlacement(partition.Spanner, element, 0D, contentHeight, partition.SpannerStyle);
            HtmlRenderFlowBlock spanner = LayoutColumnSpanner(partition.Spanner, contentWidth,
                partition.SpannerStyle, style, partition.SpannerDepth);
            spanner = ApplyColumnSpannerAncestorSemantics(spanner, partition.SpannerAncestors);
            hasPrecedingSpanner = true;
            if (!directSpanners && (spanner.ForcedBreaks.Any(item => IsForcedPageBreak(item.Target))
                || IsForcedPageBreak(spanner.BreakBefore) || IsForcedPageBreak(spanner.BreakAfter)
                    || partition.SpannerStyle.AvoidBreakBeforeScope != HtmlFragmentationScope.None
                    || partition.SpannerStyle.AvoidBreakAfterScope != HtmlFragmentationScope.None)) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
                    "Authored fragmentation controls on a column spanner retain the source-flow fallback.",
                    HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(partition.Spanner),
                    "fragmentation controls on column spanner");
            }
            pieces.Add(spanner);
            contentHeight += spanner.Height;
        }

        HtmlRenderFlowBlock content = ApplyColumnSourceOpacity(
            BuildSpanningFlowBlock(pieces, contentWidth, source, directSpanners), partitions);
        IReadOnlyList<HtmlRenderVisual> contentVisuals = content.Visuals;
        IReadOnlyList<double> contentBreakOffsets = content.BreakOffsets;
        var contentSafeBreakRanges = new List<HtmlRenderBreakRange>(content.SafeBreakRanges);
        IReadOnlyList<HtmlRenderLineBreakGroup> contentLineBreakGroups = content.LineBreakGroups;
        IReadOnlyList<HtmlRenderAvoidBreakRange> contentAvoidBreakRanges = content.AvoidBreakRanges;
        IReadOnlyList<HtmlRenderForcedBreak> contentForcedBreaks = content.ForcedBreaks;
        IReadOnlyList<HtmlCssRunningStringAssignment> runningStringAssignments = content.RunningStringAssignments;

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
        double pagedPaintExtent = _options.Mode == HtmlRenderMode.Paged && style.OverflowY == "visible"
            ? Math.Max(outerHeight, contentY + content.PagedPaintExtent) : outerHeight;
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
            .Where(offset => offset > 0.0001D && offset < content.PagedPaintExtent - 0.0001D)
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
            pagedPaintExtent: pagedPaintExtent,
            forcedBreaks: contentForcedBreaks.Select(item => item.Translate(contentY)),
            avoidBreakRanges: contentAvoidBreakRanges.Select(range => range.Translate(contentY)),
            runningStringAssignments: runningStringAssignments
                .Select(assignment => assignment.Translate(contentY))
                .Concat(positionedRunningStringAssignments)
                .OrderBy(assignment => assignment.OrderOffset),
            safeBreakRanges: ConstrainParallelPaintBreakRanges(
                ProjectColumnTerminalSafeBreakRanges(
                    contentSafeBreakRanges, safeContentHeight, contentY, pagedPaintExtent).ToArray(), visuals));
        block.UnfragmentedHeight = columnContext?.UnfragmentedHeight;
        if (columnContext != null && _columnPaginationAttempt is { } attempt
            && ReferenceEquals(attempt.Element, element)) attempt.Context = columnContext;
        if (pagedColumnRegion) _pagedColumnRegions[element] = block;
        return true;
    }

    /// <summary>Retains loss reporting for column controls inside full-width spanners.</summary>
    private HtmlRenderFlowBlock LayoutColumnSpanner(IElement element, double width,
        HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, int depth) {
        // The spanner is outside its own column flow. A column-only control
        // would target an outer multicol context, which is absent here.
        if (_columnContentDepth == 0) {
            style = style.Clone();
            if (style.BreakBefore == HtmlPageBreakTarget.Column) style.BreakBefore = HtmlPageBreakTarget.None;
            if (style.BreakAfter == HtmlPageBreakTarget.Column) style.BreakAfter = HtmlPageBreakTarget.None;
            style.AvoidBreakBeforeScope &= ~HtmlFragmentationScope.Column;
            style.AvoidBreakAfterScope &= ~HtmlFragmentationScope.Column;
            style.AvoidBreakInsideScope &= ~HtmlFragmentationScope.Column;
            style.UnsupportedColumnBreaks = string.Empty;
        }
        _columnContentDepth++;
        _columnBreakFallbackDepth++;
        try {
            return LayoutElement(element, width, style, parentStyle, depth);
        } finally {
            _columnBreakFallbackDepth--;
            _columnContentDepth--;
        }
    }

    private HtmlRenderFlowBlock? BuildMultiColumnPartition(
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
        bool extendRuleToPageFragmentEnd,
        HtmlRenderFlowBlock? prefix,
        out ColumnPaginationContext? columnContext) {
        columnContext = null;
        if (children.Count == 0) return null;
        // Empty ancestor fragments still carry source navigation and boundary
        // facts. They do not allocate a minimum-height column set between two
        // adjacent spanners, which could otherwise displace a fitting line.
        if (children.All(child => child.Height == 0D))
            return BuildSpanningFlowBlock(children, contentWidth, source, false);

        var visuals = new List<HtmlRenderVisual>();
        var breakOffsets = new SortedSet<double>();
        var safeBreakRanges = new List<HtmlRenderBreakRange>();
        var lineBreakGroups = new List<HtmlRenderLineBreakGroup>();
        var avoidBreakRanges = new List<HtmlRenderAvoidBreakRange>();
        var forcedBreaks = new List<HtmlRenderForcedBreak>();
        var runningStringAssignments = new List<HtmlCssRunningStringAssignment>();

        double targetHeight = ResolveSpanningColumnHeight(children, requestedCount, style, offsetY, finalPartition);
        ColumnPageBands? pageBands = finalPartition && pagedColumnRegion
            ? ResolveSpanningColumnPageBands(owner, contentWidth, style, requestedCount, targetHeight, offsetY, prefix, out columnContext)
            : null;
        bool hadPhysicalColumnBands = pageBands != null;
        MultiColumnPlan plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount,
            throwOnLimit: true, pageBands: pageBands,
            balanceVisibleOverflow: !finalPartition || style.ColumnFill != "auto");
        if (pageBands != null)
            plan = ProjectMultiColumnPaintOverflow(plan, targetHeight, pageBands);
        if (pageBands is { Supported: false }) {
            pageBands = null;
            columnContext = null;
            plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount, throwOnLimit: true,
                balanceVisibleOverflow: !finalPartition || style.ColumnFill != "auto");
        }
        if (pageBands == null && (hadPhysicalColumnBands || ResolveDeclaredColumnContentHeight(style).HasValue
            || !finalPartition || style.ColumnFill != "auto"))
            plan = ProjectMultiColumnPaintOverflow(plan, targetHeight, pageBands);
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
        foreach (double offset in pageBreaks) forcedBreaks.Add(new HtmlRenderForcedBreak(offset, HtmlPageBreakTarget.Page));
        double partitionHeight = Math.Max(minimumExtent, Math.Max(targetHeight, plan.UsedHeight));
        if (pageBands != null) {
            // A single sequential column can shrink to its painted extent. An
            // advance to another column consumes the complete physical band.
            partitionHeight = Math.Max(minimumExtent, plan.UsedHeight);
        }
        AddColumnRuleVisuals(visuals, style, 0D, 0D, columnWidth, gap,
            plan.PaintColumnCount, partitionHeight, source, plan.RuleBands,
            extendToPageFragmentEnd: extendRuleToPageFragmentEnd);
        foreach (MultiColumnFragment fragment in plan.PaintFragments) {
            double x = fragment.Column * (columnWidth + gap);
            double y = fragment.Y;
            IReadOnlyList<HtmlRenderVisual> fragmentVisuals = SliceColumnFragmentVisuals(fragment);
            foreach (HtmlRenderVisual visual in fragmentVisuals) {
                visuals.Add(visual.Translate(x, y, visuals.Count));
            }
        }

        foreach (double offset in ResolveMultiColumnBreakOffsets(plan, 0D, plan.PaintExtent)) {
            breakOffsets.Add(offset);
        }
        foreach (HtmlRenderBreakRange range in ResolveMultiColumnSafeBreakRanges(plan)) {
            safeBreakRanges.Add(range);
        }
        if (pageBands != null && partitionHeight > plan.UsedHeight)
            safeBreakRanges.Add(new HtmlRenderBreakRange(plan.UsedHeight, partitionHeight));
        foreach (HtmlRenderLineBreakGroup group in EnumerateMultiColumnLineBreakGroups(plan, 0D)) {
            lineBreakGroups.Add(group);
        }
        foreach (HtmlRenderAvoidBreakRange range in EnumerateMultiColumnAvoidBreakRanges(plan, 0D)) {
            avoidBreakRanges.Add(range);
        }
        foreach (HtmlCssRunningStringAssignment assignment in EnumerateMultiColumnRunningStringAssignments(plan, 0D)) {
            runningStringAssignments.Add(assignment);
        }
        return new HtmlRenderFlowBlock(contentWidth, partitionHeight, visuals,
            HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, false, source, breakOffsets,
            lineBreakGroups: lineBreakGroups, forcedBreaks: forcedBreaks,
            avoidBreakRanges: avoidBreakRanges, safeBreakRanges: safeBreakRanges,
            runningStringAssignments: runningStringAssignments,
            pagedPaintExtent: Math.Max(partitionHeight, plan.PaintExtent));
    }

}
