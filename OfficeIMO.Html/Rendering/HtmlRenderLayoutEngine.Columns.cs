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
        IReadOnlyList<HtmlRenderFlowBlock> children = BuildColumnChildren(
            element, columnWidth, style, depth);
        double? declaredHeight = ResolveDeclaredColumnContentHeight(style);
        double targetHeight;
        if (declaredHeight.HasValue && style.ColumnFill == "auto") {
            targetHeight = declaredHeight.Value;
        } else {
            double balanced = ResolveBalancedColumnHeight(children, requestedCount, element, columnWidth);
            targetHeight = declaredHeight.HasValue ? Math.Min(declaredHeight.Value, balanced) : balanced;
        }
        targetHeight = Math.Max(0.01D, targetHeight);

        MultiColumnPlan plan = BuildMultiColumnPlan(children, targetHeight, _options.MaxColumnCount, throwOnLimit: true);
        plan = ResolveColumnReservedLayout(element, children, plan, columnWidth, targetHeight);
        EnsureMultiColumnLimit(plan.ColumnCount);
        IReadOnlyList<double> columnPageBreaks = Array.Empty<double>();
        if (style.OverflowX == "visible" && style.OverflowY == "visible") {
            plan = ResolvePagedColumnOverflow(plan, requestedCount, targetHeight, out columnPageBreaks);
        }
        double contentHeight = columnPageBreaks.Count > 0
            ? Math.Max(declaredHeight ?? targetHeight, plan.UsedHeight)
            : declaredHeight ?? Math.Max(targetHeight, plan.UsedHeight);
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
            IReadOnlyList<HtmlRenderVisual> fragmentVisuals = SliceBlockVisuals(fragment.Block, fragment.Start, fragment.End);
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
            pageName: style.PageName,
            forcedBreaks: columnPageBreaks.Select(offset => new HtmlRenderForcedBreak(contentY + offset, HtmlPageBreakTarget.Page)),
            avoidBreakRanges: ResolveColumnReservationKeepRanges(element, columnPageBreaks, contentY, contentHeight),
            runningStringAssignments: EnumerateMultiColumnRunningStringAssignments(plan, contentY)
                .Concat(positionedRunningStringAssignments)
                .OrderBy(assignment => assignment.OrderOffset));
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
        double lower = Math.Max(0.01D, totalHeight / requestedCount);
        double upper = Math.Max(lower, totalHeight);
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
            MultiColumnPlan plan = BuildMultiColumnPlan(blocks, candidate, requestedCount, throwOnLimit: false);
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
            bool fits = plan.ColumnCount <= requestedCount && plan.UsedHeight <= candidate + 0.0001D;
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
        ColumnNotePlan? notes = null, ColumnEdgeFloatPlan? floats = null) {
        var fragments = new List<MultiColumnFragment>();
        int column = 0;
        double y = floats?.Top(0) ?? 0D;
        double usedHeight = 0D;
        foreach (HtmlRenderFlowBlock child in blocks) {
            if (child.Height <= 0.0001D) {
                // Zero-advance figure anchors still participate in column
                // ownership; dropping their blocks loses the reservation owner.
                EnsureMultiColumnLimit(column + 1);
                fragments.Add(new MultiColumnFragment(child, 0D, 0D, column, y));
                continue;
            }
            double start = 0D;
            while (start < child.Height - 0.0001D) {
                double available = targetHeight - (notes?.Reserved(column) ?? 0D) - (floats?.Bottom(column) ?? 0D) - y;
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

                double end;
                if (remaining <= available + 0.0001D) {
                    end = child.Height;
                } else {
                    end = FindFragmentEnd(child, start, available, child.Height, fullPageHeight: targetHeight);
                    if (end <= start + 0.0001D && y > 0.0001D) {
                        if (column + 2 > maximumGeneratedColumns) {
                            if (throwOnLimit) EnsureMultiColumnLimit(column + 2);
                            return new MultiColumnPlan(fragments, column + 2, usedHeight);
                        }
                        column++;
                        y = floats?.Top(column) ?? 0D;
                        continue;
                    }
                    if (end <= start + 0.0001D) end = FindNextColumnBreak(child, start);
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
                    if (notes != null) end = RestrictFragmentBeforeDeferredColumnCall(child, start, end, column, targetHeight, notes);
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
                if (start < child.Height - 0.0001D) {
                    if (column + 2 > maximumGeneratedColumns) {
                        if (throwOnLimit) EnsureMultiColumnLimit(column + 2);
                        return new MultiColumnPlan(fragments, column + 2, usedHeight);
                    }
                    column++;
                    y = floats?.Top(column) ?? 0D;
                }
            }
        }
        return new MultiColumnPlan(fragments, fragments.Count == 0 ? 1 : column + 1, usedHeight);
    }

    private static double FindNextColumnBreak(HtmlRenderFlowBlock block, double start) {
        double next = block.BreakOffsets.FirstOrDefault(offset => offset > start + 0.0001D);
        return next > start + 0.0001D ? next : block.Height;
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
        if (details.Count == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported,
            "A multi-column property value used its deterministic fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
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
