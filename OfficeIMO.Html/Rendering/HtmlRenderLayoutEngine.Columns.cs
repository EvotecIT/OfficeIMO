using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool HasMultiColumnLayout(HtmlRenderBoxStyle style) => style.ColumnCount.HasValue || style.ColumnWidth.HasValue;

    private bool TryLayoutMultiColumnContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        out HtmlRenderFlowBlock block) {
        block = null!;
        if (!HasMultiColumnLayout(style)) return false;
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        double gap = style.ColumnGapWasSpecified ? style.ColumnGap : style.Font.Size;
        int requestedCount = ResolveRequestedColumnCount(style, contentWidth, gap);
        EnsureMultiColumnLimit(requestedCount);
        double columnWidth = Math.Max(0.01D, (contentWidth - gap * Math.Max(0, requestedCount - 1)) / requestedCount);
        if (TryLayoutSpanningMultiColumnContainer(
                element,
                containingWidth,
                style,
                depth,
                boxWidth,
                contentWidth,
                columnWidth,
                gap,
                requestedCount,
                source,
                out block)) {
            return true;
        }
        bool pagedColumnRegion = CanPaginateColumnRegion(element, style);
        IReadOnlyList<HtmlRenderFlowBlock> children = BuildColumnChildren(
            element, columnWidth, style, depth);
        if (BuildColumnFlowBreaks(children).Forced.Any(boundary => boundary.Page)) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
                "Authored page breaks inside multi-column content retain the source-flow fallback.",
                HtmlDiagnosticSeverity.Warning, source, "page breaks in column content");
        }
        double? declaredHeight = ResolveDeclaredColumnContentHeight(style);
        double? maximumHeight = style.MaxHeight.HasValue
            ? Math.Max(0.01D, Math.Max(style.MaxHeight.Value, style.MinHeight ?? 0D)
                - (style.BorderBox ? style.VerticalInsets : 0D)) : null;
        double naturalHeight = children.Sum(child => child.Height);
        double targetHeight;
        if (style.ColumnFill == "auto") {
            // Without a block-size constraint, sequential filling has no
            // reason to advance to another column. Measure at column width;
            // physical page bands can still supply a smaller capacity below.
            targetHeight = declaredHeight ?? Math.Max(0.01D,
                ResolveBoxHeight(ResolveNaturalColumnRunHeight(children), boxWidth, style) - style.VerticalInsets);
        } else {
            double balanced = ResolveBalancedColumnHeight(children, requestedCount, element, columnWidth);
            targetHeight = declaredHeight.HasValue ? Math.Min(declaredHeight.Value, balanced) : balanced;
            if (maximumHeight.HasValue) targetHeight = Math.Min(targetHeight, maximumHeight.Value);
        }
        targetHeight = Math.Max(0.01D, targetHeight);

        ColumnPageBands? pageBands = pagedColumnRegion
            ? ResolveColumnPageBands(element, contentWidth, declaredHeight, requestedCount, targetHeight) : null;
        MultiColumnPlan plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount,
            throwOnLimit: true, pageBands: pageBands);
        plan = ResolveColumnReservedLayout(element, children, plan, columnWidth, targetHeight, pageBands);
        if (pageBands != null && pageBands.FirstHeight < pageBands.FreshHeight - 0.0001D
            && plan.Fragments.FirstOrDefault(fragment => fragment.End > fragment.Start + 0.0001D) is { } first
            && first.Column >= requestedCount) {
            // No source content can start in either partial column. Move the
            // complete box, including its authored height, to the fresh page.
            pageBands.DeferFirstBand();
            plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount,
                throwOnLimit: true, pageBands: pageBands);
            plan = ResolveColumnReservedLayout(element, children, plan, columnWidth, targetHeight, pageBands);
        }
        pageBands?.ValidateExtent(declaredHeight ?? 0D);
        if (pageBands is { Supported: false }) {
            pageBands = null;
            plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount, throwOnLimit: true);
            plan = ResolveColumnReservedLayout(element, children, plan, columnWidth, targetHeight);
        }
        if (pageBands != null && _columnPaginationAttempt is { } attempt
            && ReferenceEquals(attempt.Element, element)) attempt.Context = pageBands.Context;
        EnsureMultiColumnLimit(plan.ColumnCount);
        double minimumColumnBoxExtent = pageBands?.MinimumBoxExtent(plan.ColumnCount) ?? 0D;
        IReadOnlyList<double> columnPageBreaks = Array.Empty<double>();
        bool shortDefiniteColumns = declaredHeight.HasValue && pageBands == null
            && declaredHeight.Value < _activePageGeometry.ContentHeight - 0.0001D;
        if (style.OverflowX == "visible" && style.OverflowY == "visible" && !shortDefiniteColumns) {
            plan = ResolvePagedColumnOverflow(plan, requestedCount, targetHeight, out columnPageBreaks, pageBands);
        }
        double contentHeight = pageBands != null
            ? Math.Max(minimumColumnBoxExtent, Math.Max(declaredHeight ?? 0D, plan.UsedHeight))
            : columnPageBreaks.Count > 0
                ? Math.Max(declaredHeight ?? targetHeight, plan.UsedHeight)
                : declaredHeight ?? (style.ColumnFill == "auto" && !(maximumHeight.HasValue && naturalHeight > maximumHeight.Value)
                    ? plan.UsedHeight : Math.Max(targetHeight, plan.UsedHeight));
        double boxHeight = ResolveBoxHeight(contentHeight, boxWidth, style);
        if (columnPageBreaks.Count > 0) boxHeight = Math.Max(boxHeight, contentHeight + style.VerticalInsets);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
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
        AddColumnRuleVisuals(visuals, style, contentX, contentY, columnWidth, gap, Math.Max(requestedCount, plan.ColumnCount), contentHeight, source);
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            double x = contentX + fragment.Column * (columnWidth + gap);
            double y = contentY + fragment.Y;
            IReadOnlyList<HtmlRenderVisual> fragmentVisuals = SliceColumnFragmentVisuals(fragment);
            foreach (HtmlRenderVisual visual in fragmentVisuals) {
                visuals.Add(visual.Translate(x, y, visuals.Count));
            }
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

        IReadOnlyList<double> breakOffsets = ResolveMultiColumnBreakOffsets(plan, contentY, outerHeight);

        block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            breakOffsets,
            lineBreakGroups: EnumerateMultiColumnLineBreakGroups(plan, contentY),
            pageName: style.PageName,
            forcedBreaks: columnPageBreaks.Select(offset => new HtmlRenderForcedBreak(contentY + offset, HtmlPageBreakTarget.Page)),
            avoidBreakRanges: ResolveColumnReservationKeepRanges(element, columnPageBreaks, contentY, contentHeight)
                .Concat(EnumerateMultiColumnAvoidBreakRanges(plan, contentY)),
            runningStringAssignments: EnumerateMultiColumnRunningStringAssignments(plan, contentY)
                .Concat(positionedRunningStringAssignments)
                .OrderBy(assignment => assignment.OrderOffset),
            safeBreakRanges: ConstrainParallelPaintBreakRanges(
                ResolveMultiColumnSafeBreakRanges(plan)
                    .Concat(pageBands != null && contentHeight > plan.UsedHeight
                        ? new[] { new HtmlRenderBreakRange(plan.UsedHeight, contentHeight) }
                        : Array.Empty<HtmlRenderBreakRange>())
                    .Select(range => range.Translate(contentY)).ToArray(), visuals));
        block.UnfragmentedHeight = pageBands?.Context.UnfragmentedHeight;
        if (pagedColumnRegion) _pagedColumnRegions[element] = block;
        return true;
    }

    private int ResolveRequestedColumnCount(HtmlRenderBoxStyle style, double contentWidth, double gap) {
        int count = style.ColumnCount ?? _options.MaxColumnCount;
        if (style.ColumnWidth.HasValue) {
            int fitting = Math.Max(1, (int)Math.Floor((contentWidth + gap) / (style.ColumnWidth.Value + gap)));
            count = style.ColumnCount.HasValue ? Math.Min(count, fitting) : fitting;
        }
        return Math.Max(1, count);
    }

    private static double? ResolveDeclaredColumnContentHeight(HtmlRenderBoxStyle style) {
        if (!style.ExplicitHeight.HasValue) return null;
        double boxHeight = ResolveBoxHeight(0D, 0D, style);
        return Math.Max(0.01D, boxHeight - style.VerticalInsets);
    }

    private double ResolveBalancedColumnHeight(IReadOnlyList<HtmlRenderFlowBlock> blocks, int requestedCount,
        IElement? owner = null, double width = 0D) {
        ColumnEdgeFloatEntry[] floats = _columnEdgeFloatEntries.Values.Where(e => ReferenceEquals(e.Owner, owner)).ToArray();
        double totalHeight = blocks.Sum(block => block.Height) + floats.Sum(e => e.Body.Height);
        double upper = Math.Max(0.01D, ResolveNaturalColumnRunHeight(blocks) + floats.Sum(entry => entry.Body.Height));
        double lower = Math.Max(0.01D, Math.Min(upper, totalHeight / requestedCount));
        int fittingColumnCount = Math.Max(requestedCount,
            BuildMultiColumnPlan(blocks, upper, _options.MaxColumnCount, throwOnLimit: true).ColumnCount);
        if (floats.Length > 0) {
            MultiColumnPlan initial = BuildMultiColumnPlan(blocks, Math.Max(0.01D, upper), _options.MaxColumnCount, throwOnLimit: true);
            foreach (ColumnEdgeFloatEntry entry in floats) {
                FindColumnFloatAnchor(initial, entry.Anchor, upper, out double anchorHeight);
                lower = Math.Max(lower, entry.Body.Height + anchorHeight + 0.01D);
            }
            upper = Math.Max(upper, lower);
        }
        if (totalHeight <= 0.01D || requestedCount <= 1) return Math.Max(0.01D, upper);
        for (int iteration = 0; iteration < 24; iteration++) {
            double candidate = (lower + upper) / 2D;
            MultiColumnPlan plan = BuildMultiColumnPlan(blocks, candidate, fittingColumnCount, throwOnLimit: false);
            if (floats.Length > 0) {
                try {
                    plan = BuildMultiColumnPlan(blocks, candidate, _options.MaxColumnCount, throwOnLimit: true);
                    plan = ResolveColumnReservedLayout(owner!, blocks, plan, width, candidate);
                } catch (HtmlDomLimitException error) when (error.Code == HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded) {
                    // An infeasible height is a search result, not a failure of
                    // the final layout. Other resource/cancellation limits still
                    // propagate, and the final plan enforces the column limit.
                    lower = candidate;
                    continue;
                }
            }
            bool fits = plan.ColumnCount <= fittingColumnCount && plan.UsedHeight <= candidate + 0.0001D;
            if (fits) upper = candidate;
            else lower = candidate;
        }
        return upper;
    }

    private MultiColumnPlan BuildMultiColumnPlan(
        IReadOnlyList<HtmlRenderFlowBlock> blocks,
        double targetHeight,
        int maximumGeneratedColumns,
        bool throwOnLimit,
        ColumnNotePlan? notes = null, ColumnEdgeFloatPlan? floats = null, ColumnPageBands? pageBands = null) {
        var fragments = new List<MultiColumnFragment>();
        ColumnFlowBreaks source = BuildColumnFlowBreaks(blocks);
        int forcedIndex = 0;
        double childSourceStart = 0D;
        int column = 0;
        double y = floats?.Top(0) ?? 0D;
        double usedHeight = 0D;
        for (int childIndex = 0; childIndex < blocks.Count; childIndex++) {
            HtmlRenderFlowBlock child = blocks[childIndex];
            ApplyForcedColumnBoundaries(childSourceStart, childIndex);
            if (column + 1 > maximumGeneratedColumns) {
                if (throwOnLimit) EnsureMultiColumnLimit(column + 1);
                return new MultiColumnPlan(fragments, column + 1, usedHeight);
            }
            if (child.Height <= 0.0001D) {
                // Zero-advance figure anchors still participate in column
                // ownership; dropping their blocks loses the reservation owner.
                EnsureMultiColumnLimit(column + 1);
                fragments.Add(new MultiColumnFragment(child, 0D, 0D, column, y));
                continue;
            }
            double start = 0D;
            while (start < child.Height - 0.0001D) {
                CheckCancellation();
                ChargeLayoutOperation("column body fragmentation");
                ApplyForcedColumnBoundaries(childSourceStart + start, childIndex);
                if (column + 1 > maximumGeneratedColumns) {
                    if (throwOnLimit) EnsureMultiColumnLimit(column + 1);
                    return new MultiColumnPlan(fragments, column + 1, usedHeight);
                }
                double columnHeight = pageBands?.Height(column) ?? targetHeight;
                double freshHeight = pageBands?.FreshHeight ?? targetHeight;
                double available = columnHeight - (notes?.Reserved(column) ?? 0D) - (floats?.Bottom(column) ?? 0D) - y;
                double remaining = child.Height - start;
                if (available <= 0.0001D) {
                    if (column + 2 > maximumGeneratedColumns) {
                        if (throwOnLimit) EnsureMultiColumnLimit(column + 2);
                        return new MultiColumnPlan(fragments, column + 2, usedHeight);
                    }
                    column++;
                    y = floats?.Top(column) ?? 0D;
                    continue;
                }

                double sourceStart = childSourceStart + start;
                double maximumEnd = childSourceStart + child.Height;
                double nextForced = NextForcedColumnBoundary();
                bool forcedEnd = nextForced <= maximumEnd + 0.0001D
                    && nextForced <= sourceStart + available + 0.0001D;
                bool freshColumn = y <= (floats?.Top(column) ?? 0D) + 0.0001D
                    && columnHeight >= freshHeight - 0.0001D;
                double end;
                if (forcedEnd) {
                    // A forced class-A boundary takes precedence over keeps and
                    // paragraph constraints. It must be consumed even when the
                    // remainder would otherwise fit this column whole.
                    end = nextForced - childSourceStart;
                } else if (remaining <= available + 0.0001D
                    && !ColumnKeepRequiresBacktracking(source, sourceStart, maximumEnd,
                        sourceStart + available, freshHeight, freshColumn)) {
                    end = child.Height;
                } else {
                    ChargeLayoutOperations((long)source.Metrics.LineBreakGroups.Count
                        + source.Metrics.AvoidBreakRanges.Count, "column source constraints");
                    end = FindFragmentEnd(source.Metrics, sourceStart, available,
                        Math.Min(maximumEnd, nextForced), fullPageHeight: freshHeight,
                        fragmentationScope: HtmlFragmentationScope.Column, isFreshFragmentainer: freshColumn) - childSourceStart;
                    if (end <= start + 0.0001D
                        && (y > 0.0001D || columnHeight < freshHeight - 0.0001D)) {
                        if (column + 2 > maximumGeneratedColumns) {
                            if (throwOnLimit) EnsureMultiColumnLimit(column + 2);
                            return new MultiColumnPlan(fragments, column + 2, usedHeight);
                        }
                        column++;
                        y = floats?.Top(column) ?? 0D;
                        continue;
                    }
                    if (end <= start + 0.0001D) end = FindNextColumnBreak(child, start);
                    if (childSourceStart + end >= nextForced - 0.0001D) {
                        end = nextForced - childSourceStart;
                        forcedEnd = true;
                    }
                }

                if (notes != null || floats != null) {
                    CheckCancellation();
                    ChargeLayoutOperation("column note body fragmentation");
                    // A forced atomic fragment may exceed a fresh column, but
                    // must never paint through an existing note reservation.
                    if ((notes?.Reserved(column) ?? 0D) + (floats?.Reserved(column) ?? 0D) > 0.0001D
                        && end - start > available + 0.0001D) {
                        EnsureMultiColumnLimit(column + 2);
                        column++;
                        y = floats?.Top(column) ?? 0D;
                        continue;
                    }
                    if (notes != null) end = RestrictFragmentBeforeDeferredColumnCall(child, start, end, column, freshHeight, notes);
                    if (end <= start + 0.0001D) {
                        EnsureMultiColumnLimit(column + 2);
                        column++;
                        y = floats?.Top(column) ?? 0D;
                        continue;
                    }
                }

                double height = Math.Max(0.01D, end - start);
                fragments.Add(new MultiColumnFragment(child, start, end, column, y));
                y += height;
                usedHeight = Math.Max(usedHeight, y);
                start = end;
                if (start < child.Height - 0.0001D && !forcedEnd) {
                    if (column + 2 > maximumGeneratedColumns) {
                        if (throwOnLimit) EnsureMultiColumnLimit(column + 2);
                        return new MultiColumnPlan(fragments, column + 2, usedHeight);
                    }
                    column++;
                    y = floats?.Top(column) ?? 0D;
                }
            }
            childSourceStart += child.Height;
        }
        return new MultiColumnPlan(fragments, fragments.Count == 0 ? 1 : column + 1, usedHeight);

        double NextForcedColumnBoundary() {
            while (forcedIndex < source.Forced.Count
                && (!source.Forced[forcedIndex].Column || source.Forced[forcedIndex].Page)) forcedIndex++;
            return forcedIndex < source.Forced.Count ? source.Forced[forcedIndex].Offset : double.PositiveInfinity;
        }

        void ApplyForcedColumnBoundaries(double offset, int childIndex) {
            while (NextForcedColumnBoundary() <= offset + 0.0001D
                && source.Forced[forcedIndex].ChildIndex <= childIndex) {
                CheckCancellation();
                ChargeLayoutOperation("authored column advance");
                column++;
                if (throwOnLimit) EnsureMultiColumnLimit(column + 1);
                y = floats?.Top(column) ?? 0D;
                forcedIndex++;
            }
        }
    }

    private IReadOnlyList<HtmlRenderVisual> SliceColumnFragmentVisuals(MultiColumnFragment fragment) =>
        fragment.Start <= 0.0001D && fragment.End >= fragment.Block.Height - 0.0001D
            ? fragment.Block.Visuals
            : SliceBlockVisuals(fragment.Block, fragment.Start, fragment.End);

    private static double FindNextColumnBreak(HtmlRenderFlowBlock block, double start) {
        double next = block.BreakOffsets.FirstOrDefault(offset => offset > start + 0.0001D);
        if (next <= start + 0.0001D) next = block.Height;
        foreach (HtmlRenderBreakRange range in block.SafeBreakRanges) {
            if (range.End <= start + 0.0001D) continue;
            double candidate = range.Start > start + 0.0001D ? range.Start : range.End;
            next = Math.Min(next, candidate);
            break;
        }
        return next;
    }

    private static IReadOnlyList<double> ResolveMultiColumnBreakOffsets(MultiColumnPlan plan, double contentY, double outerHeight) {
        if (plan.Fragments.Count == 0) return new[] { outerHeight };
        var offsetsByColumn = new List<HashSet<long>>(plan.ColumnCount);
        var heightsByColumn = new double[plan.ColumnCount];
        for (int column = 0; column < plan.ColumnCount; column++) offsetsByColumn.Add(new HashSet<long> { 0L });
        var candidates = new SortedSet<long>();
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            HashSet<long> offsets = offsetsByColumn[fragment.Column];
            AddColumnBreak(offsets, candidates, fragment.Y);
            AddColumnBreak(offsets, candidates, fragment.Y + fragment.Height);
            heightsByColumn[fragment.Column] = Math.Max(heightsByColumn[fragment.Column], fragment.Y + fragment.Height);
            foreach (double childOffset in fragment.Block.BreakOffsets) {
                if (childOffset < fragment.Start - 0.0001D || childOffset > fragment.End + 0.0001D) continue;
                AddColumnBreak(offsets, candidates, fragment.Y + childOffset - fragment.Start);
            }
        }

        var resolved = new List<double>();
        foreach (long candidateKey in candidates) {
            double candidate = candidateKey / 10000D;
            if (candidate <= 0.0001D) continue;
            bool safe = true;
            for (int column = 0; column < plan.ColumnCount; column++) {
                if (candidate >= heightsByColumn[column] - 0.0001D) continue;
                if (offsetsByColumn[column].Contains(candidateKey)) continue;
                safe = false;
                break;
            }
            if (safe) resolved.Add(Math.Min(outerHeight, contentY + candidate));
        }
        resolved.Add(outerHeight);
        return resolved;
    }

    private static void AddColumnBreak(ISet<long> offsets, ISet<long> candidates, double value) {
        long key = (long)Math.Round(value * 10000D, MidpointRounding.AwayFromZero);
        offsets.Add(key);
        candidates.Add(key);
    }

    private static IEnumerable<HtmlCssRunningStringAssignment> EnumerateMultiColumnRunningStringAssignments(
        MultiColumnPlan plan,
        double offsetY) {
        double totalFlowHeight = plan.Fragments.Sum(fragment => fragment.Height);
        double flowOffset = 0D;
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            bool finalFragment = fragment.End >= fragment.Block.Height - 0.0001D;
            foreach (HtmlCssRunningStringAssignment assignment in fragment.Block.RunningStringAssignments) {
                if (assignment.Offset < fragment.Start - 0.0001D
                    || assignment.Offset > fragment.End + 0.0001D
                    || (!finalFragment && assignment.Offset >= fragment.End - 0.0001D)) {
                    continue;
                }
                double visualOffset = assignment.Offset + offsetY + fragment.Y - fragment.Start;
                double orderOffset = offsetY + (flowOffset + assignment.OrderOffset - fragment.Start)
                    * plan.UsedHeight / Math.Max(0.01D, totalFlowHeight);
                yield return assignment.Place(visualOffset, orderOffset);
            }
            flowOffset += fragment.Height;
        }
    }

    private void EnsureMultiColumnLimit(int count) {
        if (count <= _options.MaxColumnCount) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded,
            "Multi-column layout exceeded the configured maximum column count.",
            nameof(HtmlRenderOptions.MaxColumnCount),
            count,
            _options.MaxColumnCount);
    }

    private void ReportUnsupportedMultiColumnValues(IElement element, HtmlRenderBoxStyle style) {
        var details = new List<string>(4);
        if (style.UnsupportedColumns.Length > 0) details.Add(style.UnsupportedColumns);
        if (style.UnsupportedColumnFill.Length > 0) details.Add("column-fill=" + style.UnsupportedColumnFill);
        if (style.UnsupportedColumnSpan.Length > 0) details.Add("column-span=" + style.UnsupportedColumnSpan);
        if (style.UnsupportedColumnRule.Length > 0) details.Add(style.UnsupportedColumnRule);
        if (_columnContentDepth > 0 && style.UnsupportedColumnBreaks.Length > 0
            && !HasOrdinaryColumnBreakContext(element, style))
            details.Add(style.UnsupportedColumnBreaks.TrimStart(';'));
        if (details.Count == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
            "A multi-column property value used its deterministic fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
    }

    private bool HasOrdinaryColumnBreakContext(IElement element, HtmlRenderBoxStyle style) {
        for (IElement? current = element; current != null; current = current.ParentElement) {
            HtmlRenderBoxStyle currentStyle;
            if (ReferenceEquals(current, element)) currentStyle = style;
            else if (!_layoutStyles.TryGetValue(current, out currentStyle!)) continue;
            if (!ReferenceEquals(current, element) && HasMultiColumnLayout(currentStyle)) return true;
            if (currentStyle.Display is "flex" or "inline-flex" or "grid" or "inline-grid" or "table" or "inline-table"
                || currentStyle.WritingMode != "horizontal-tb" || currentStyle.ColumnSpan == "all"
                || currentStyle.FloatSide != "none" || HasMultiColumnLayout(currentStyle)) return false;
        }
        return false;
    }

    private sealed class MultiColumnPlan {
        internal MultiColumnPlan(IReadOnlyList<MultiColumnFragment> fragments, int columnCount, double usedHeight) {
            Fragments = fragments;
            ColumnCount = columnCount;
            UsedHeight = usedHeight;
        }
        internal IReadOnlyList<MultiColumnFragment> Fragments { get; }
        internal int ColumnCount { get; }
        internal double UsedHeight { get; }
    }

    private sealed class MultiColumnFragment {
        internal MultiColumnFragment(HtmlRenderFlowBlock block, double start, double end, int column, double y) {
            Block = block;
            Start = start;
            End = end;
            Column = column;
            Y = y;
        }
        internal HtmlRenderFlowBlock Block { get; }
        internal double Start { get; }
        internal double End { get; }
        internal int Column { get; }
        internal double Y { get; }
        internal double Height => Math.Max(0.01D, End - Start);
    }
}
