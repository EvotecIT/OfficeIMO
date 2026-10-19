namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double SkipUnpaintedLeadingMarginAtPageStart(HtmlRenderFlowBlock block, double start,
        HtmlFragmentationScope fragmentationScope = HtmlFragmentationScope.Page) {
        // Flex items do not collapse adjoining margins. A completed item's
        // bottom margin can be immediately followed by the next item's top
        // margin; truncate each unpainted interval without treating it as a
        // content continuation or crossing retained paint/metadata.
        while (start < block.Height - 0.0001D) {
            double next = SkipOneUnpaintedLeadingMarginAtPageStart(block, start, fragmentationScope);
            if (next <= start + 0.0001D) break;
            start = next;
        }
        return start;
    }

    private double SkipOneUnpaintedLeadingMarginAtPageStart(HtmlRenderFlowBlock block, double start, HtmlFragmentationScope fragmentationScope) {
        double discardableMargin = ResolvePageStartDiscardableMargin(block, start);
        if (discardableMargin <= 0.0001D) return start;

        double afterMargin = Math.Min(block.Height, start + discardableMargin);
        if (block.ForcedBreaks.Any(item => IsForcedBreakForScope(item.Target, fragmentationScope) && item.Offset > start + 0.0001D && item.Offset <= afterMargin + 0.0001D)
            || block.RunningStringAssignments.Any(item => item.Offset >= start - 0.0001D && item.Offset < afterMargin - 0.0001D)
            || SliceBlockVisuals(block, start, afterMargin).Any(visual =>
                ContainsPageStartMarginContent(visual, block))) {
            return start;
        }

        return afterMargin;
    }

    private static double ResolvePageStartDiscardableMargin(HtmlRenderFlowBlock block, double start) {
        double discardableMargin = block.HasCollapsibleMargins
            && block.CollapsibleMarginBottom > 0.0001D
            && Math.Abs(start - (block.Height - block.CollapsibleMarginBottom)) <= 0.0001D
                ? block.CollapsibleMarginBottom
                : 0D;
        foreach (HtmlInlineBreakProgress progress in block.InlineBreakProgress) {
            if (Math.Abs(progress.Offset - start) > 0.0001D || !(progress.IsBlockEntry || progress.IsBlockExit || progress.IsFlexGap)
                || progress.PageStartDiscardableMargin <= 0.0001D) continue;
            discardableMargin = progress.PageStartDiscardableMargin;
            break;
        }
        return discardableMargin;
    }

    private bool ContainsPageStartMarginContent(HtmlRenderVisual visual, HtmlRenderFlowBlock spacingBlock) {
        // Print-fitting boxes and empty wrappers measure layout but do not paint.
        // Retain every other leaf, including navigation and bookmark metadata.
        if (visual is HtmlRenderLayoutBox) return false;
        // A body's continuous background paints through margins and row gaps.
        // Discarding spacing must not mistake the backdrop for content or discard borders,
        // positioned paint, navigation metadata, or other authored boxes.
        if (ReferenceEquals(spacingBlock.OwnerElement, _document.Body)
            && visual is HtmlRenderShape shape && !shape.IsAtomicReplacedPlaceholder
            && shape.Source == spacingBlock.Source && shape.LinkUri == null
            && shape.InnerShape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle
            && shape.InnerShape.StrokeWidth <= 0D
            && shape.Width >= spacingBlock.Width - 0.0001D
            && shape.Height >= spacingBlock.Height - 0.0001D) return false;
        IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
        return children == null || children.Any(child => ContainsPageStartMarginContent(child, spacingBlock));
    }

    private static double FindFragmentEnd(HtmlRenderFlowBlock block, double start, double available, double? maximumEnd = null,
        double fullPageHeight = 0D, HtmlRenderTrailingGroup? reservedFooter = null,
        HtmlFragmentationScope fragmentationScope = HtmlFragmentationScope.Page,
        bool isFreshFragmentainer = true) {
        double limit = Math.Min(maximumEnd ?? block.Height, Math.Min(block.Height, start + available));
        IReadOnlyList<double> offsets = block.BreakOffsets;
        if (block.SafeBreakRanges.Count > 0) {
            var candidates = new SortedSet<double>(offsets.Where(offset => offset > start && offset <= limit));
            foreach (HtmlRenderBreakRange range in block.SafeBreakRanges) {
                Add(range.Start);
                Add(Math.Min(limit, range.End));
            }
            // Keeps and paragraph constraints can reject the available limit.
            // Their endpoints are the only additional positions where that
            // decision changes within a content-free interval.
            foreach (HtmlRenderAvoidBreakRange range in block.AvoidBreakRanges) {
                if ((range.Scope & fragmentationScope) != 0) Add(range.Start);
            }
            foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) {
                Add(group.Start);
                Add(group.End);
                foreach (double offset in group.Offsets) Add(offset);
            }
            offsets = candidates.ToArray();

            void Add(double candidate) {
                if (candidate <= start + 0.0001D || candidate > limit + 0.0001D) return;
                int low = 0, high = block.SafeBreakRanges.Count;
                while (low < high) {
                    int middle = low + ((high - low) >> 1);
                    if (block.SafeBreakRanges[middle].Start <= candidate) low = middle + 1;
                    else high = middle;
                }
                if (low > 0 && candidate <= block.SafeBreakRanges[low - 1].End) candidates.Add(candidate);
            }
        }
        for (int index = UpperBound(offsets, limit + 0.0001D) - 1; index >= 0; index--) {
            double offset = offsets[index];
            if (offset <= start + 0.0001D) break;
            // Finishing the body with its reserved footer consumes the original
            // footer too. Authored keeps spanning that footer are then complete,
            // even though only body geometry belongs to the sliced fragment.
            double consumedEnd = reservedFooter != null && Math.Abs(offset - reservedFooter.ContentEndsAt) <= 0.0001D
                ? reservedFooter.SourceEndsAt : offset;
            if (IsAllowedLineBreak(block, start, offset)
                && !BreaksAvoidedRangeThatFitsPage(block, start, consumedEnd, fullPageHeight, limit,
                    fragmentationScope, isFreshFragmentainer)) return offset;
        }

        return start;
    }

    // Keep authored ranges intact. A synthetic flex-row keep is weaker when
    // moving the whole row would leave more than half the page unused.
    private static bool BreaksAvoidedRangeThatFitsPage(HtmlRenderFlowBlock block, double start, double candidate,
        double fullPageHeight, double pageLimit, HtmlFragmentationScope fragmentationScope = HtmlFragmentationScope.Page,
        bool isFreshFragmentainer = true) =>
        fullPageHeight > 0D
        && (((fragmentationScope & HtmlFragmentationScope.Column) != 0
                && (block.AvoidBreakInsideScope & HtmlFragmentationScope.Column) != 0
                && start <= 0.0001D && candidate > 0.0001D && candidate < block.Height - 0.0001D
                && (block.Height <= fullPageHeight + 0.0001D || !isFreshFragmentainer))
            || block.AvoidBreakRanges.Any(range =>
                (range.Scope & fragmentationScope) != 0
                && start <= range.Start + 0.0001D
                && candidate > range.Start + 0.0001D
                && candidate < range.End - 0.0001D
                && (range.End - range.Start <= fullPageHeight + 0.0001D
                    || (range.Scope & fragmentationScope & HtmlFragmentationScope.Column) != 0
                        && (!isFreshFragmentainer || range.Start > start + 0.0001D))
                && !(range.Soft && pageLimit - range.Start > fullPageHeight * 0.5D + 0.0001D)));

    private static HtmlRenderTrailingGroup? ResolveTrailingGroup(HtmlRenderFlowBlock block, double start, double available, double fullPageHeight, out double fragmentLimit) {
        HtmlRenderTrailingGroup? active = block.TrailingGroups.FirstOrDefault(group => group.AppliesAt(start));
        if (active != null) {
            fragmentLimit = active.ContentEndsAt;
            return active;
        }

        HtmlRenderTrailingGroup? upcoming = block.TrailingGroups
            .Where(group => group.StartsAt > start + 0.0001D && group.StartsAt < start + available - 0.0001D)
            .OrderBy(group => group.StartsAt)
            .FirstOrDefault();
        if (upcoming == null) {
            fragmentLimit = block.Height;
            return null;
        }

        double candidateAvailable = Math.Max(0D, available - upcoming.Height);
        double candidateEnd = FindFragmentEnd(block, start, candidateAvailable, upcoming.ContentEndsAt, fullPageHeight, upcoming);
        if (candidateEnd > upcoming.StartsAt + 0.0001D) {
            fragmentLimit = upcoming.ContentEndsAt;
            return upcoming;
        }

        fragmentLimit = upcoming.StartsAt;
        return null;
    }

    private static bool IsAllowedLineBreak(HtmlRenderFlowBlock block, double start, double candidate, bool checkInteriorBreaks = false) {
        foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) {
            // A final-line cut before a trailing paragraph margin leaves no widows.
            // Other final-line cuts must retain authored fixed-height flow.
            if (group.FinalLineMarginBreak && candidate >= group.End - 0.0001D
                && ResolvePageStartDiscardableMargin(block, candidate) > 0.0001D) continue;
            IReadOnlyList<double> offsets = group.Offsets;
            int candidateIndex = UpperBound(offsets, candidate + 0.0001D) - 1;
            bool exactLineBreak = candidateIndex >= 0 && Math.Abs(offsets[candidateIndex] - candidate) <= 0.0001D;
            // A flex sibling can supply a break in this item's unpainted line-box
            // space. Only flex rows need the wider paragraph-span check.
            if (!exactLineBreak && (!(group.CheckInteriorBreaks || checkInteriorBreaks)
                || candidate <= group.Start + 0.0001D || candidate >= group.End - 0.0001D)) continue;
            int firstFragmentLine = UpperBound(offsets, start + 0.0001D);
            int fragmentLines = candidateIndex >= firstFragmentLine ? candidateIndex - firstFragmentLine + 1 : 0;
            int remainingLines = offsets.Count - candidateIndex - 1 + (group.HasImplicitFinalLine ? 1 : 0);
            if (fragmentLines < group.Orphans || remainingLines < group.Widows) return false;
        }

        return true;
    }

    private static int UpperBound(IReadOnlyList<double> values, double target) {
        int low = 0;
        int high = values.Count;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (values[middle] <= target) low = middle + 1;
            else high = middle;
        }

        return low;
    }

    private static bool HasInternalForcedBreak(HtmlRenderFlowBlock block) =>
        TryGetNextForcedBreak(block.ForcedBreaks, 0D, out HtmlRenderForcedBreak? forcedBreak)
        && forcedBreak!.Offset < block.Height - 0.0001D;

    private static bool TryGetNextForcedBreak(
        IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks,
        double offset,
        out HtmlRenderForcedBreak? forcedBreak,
        ISet<object>? consumedBoundaries = null,
        bool includeProjectedBoundaries = false) {
        int low = 0;
        int high = forcedBreaks.Count;
        double target = offset + (consumedBoundaries == null ? 0.0001D : -0.0001D);
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (forcedBreaks[middle].Offset <= target) low = middle + 1;
            else high = middle;
        }

        while (low < forcedBreaks.Count) {
            HtmlRenderForcedBreak item = forcedBreaks[low];
            if (IsForcedPageBreak(item.Target)
                && (consumedBoundaries == null || (includeProjectedBoundaries || !item.Projected)
                    && !consumedBoundaries.Contains(item.Boundary)
                    && !(item.Edge == HtmlRenderBreakEdge.Leading && item.Offset <= 0.0001D))) break;
            low++;
        }
        forcedBreak = low < forcedBreaks.Count ? forcedBreaks[low] : null;
        if (forcedBreak != null && consumedBoundaries != null && includeProjectedBoundaries) {
            // Adjacent after/before declarations and a named-page change share
            // one source boundary. A flex gap can separate their source offsets;
            // retain the first cut while combining page-side and name controls.
            forcedBreak = FoldForcedPageBoundary(forcedBreaks, low, forcedBreaks.Count);
        }
        return forcedBreak != null;
    }

    /// <summary>Combines one logical boundary without changing its physical cut or source ownership.</summary>
    private static HtmlRenderForcedBreak FoldForcedPageBoundary(
        IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, int firstIndex, int endIndex) {
        HtmlRenderForcedBreak first = forcedBreaks[firstIndex];
        HtmlRenderForcedBreak last = first;
        HtmlRenderForcedBreak? side = null;
        string? pageName = first.PageName;
        bool changesPageName = first.ChangesPageName;
        for (int index = firstIndex; index < endIndex; index++) {
            HtmlRenderForcedBreak candidate = forcedBreaks[index];
            if (!ReferenceEquals(candidate.Boundary, first.Boundary) || !IsForcedPageBreak(candidate.Target)) continue;
            last = candidate;
            // A plain page force, including a named-page marker, cannot erase
            // a side requirement. The following child's side wins over an
            // after declaration even when signed flow reverses their offsets.
            if (candidate.Target is HtmlPageBreakTarget.Left or HtmlPageBreakTarget.Right
                && (side == null || side.IsAfter || !candidate.IsAfter)) side = candidate;
            if (candidate.ChangesPageName) {
                pageName = candidate.PageName;
                changesPageName = true;
            }
        }
        return ReferenceEquals(first, last) ? first
            : new HtmlRenderForcedBreak(first.Offset, side?.Target ?? last.Target, pageName, changesPageName,
                first.Boundary, first.Edge, first.IsAfter, first.Projected, first.AdvancesFollowingFlow);
    }

    private HtmlPageBreakTarget ResolveForcedBreakAt(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, double offset) =>
        ResolveForcedBoundaryAt(forcedBreaks, offset)?.Target ?? HtmlPageBreakTarget.None;

    /// <summary>Combines forced types; only a later page-side value replaces an existing side.</summary>
    private static HtmlPageBreakTarget CombineForcedPageTargets(HtmlPageBreakTarget earlier, HtmlPageBreakTarget later) =>
        !IsForcedPageBreak(later) ? earlier
            : later is HtmlPageBreakTarget.Left or HtmlPageBreakTarget.Right ? later
            : earlier is HtmlPageBreakTarget.Left or HtmlPageBreakTarget.Right ? earlier : later;

    private HtmlPageBreakTarget ResolveRootBreakBefore(HtmlRenderFlowBlock block) =>
        CombineForcedPageTargets(IsForcedPageBreak(block.BreakBefore) ? block.BreakBefore : HtmlPageBreakTarget.None,
            ResolveForcedBreakAt(block.ForcedBreaks, 0D));

    private HtmlPageBreakTarget ResolveRootBreakAfter(HtmlRenderFlowBlock block) =>
        CombineForcedPageTargets(ResolveForcedBreakAt(block.ForcedBreaks, block.Height), block.BreakAfter);

    /// <summary>Skips an empty initial run that has entered no source content on the physical page.</summary>
    private void ConsumeUnenteredPageBoundaries(HtmlRenderFlowBlock block, ISet<object> consumedBoundaries) {
        if (block.ForcedBreaks.Count == 0 || block.ForcedBreaks[0].Offset > 0.0001D) return;
        ChargeLayoutOperations(2L * block.ForcedBreaks.Count, "unentered physical boundary qualification");
        var namedBoundaries = new HashSet<object>(block.ForcedBreaks
            .Where(item => item.ChangesPageName).Select(item => item.Boundary));
        foreach (HtmlRenderForcedBreak item in block.ForcedBreaks) {
            if (item.Offset > 0.0001D) break;
            if (item.Target == HtmlPageBreakTarget.Page && item.IsAfter && !item.Projected
                && !item.AdvancesFollowingFlow && !namedBoundaries.Contains(item.Boundary))
                consumedBoundaries.Add(item.Boundary);
        }
    }

    /// <summary>Consumes one source boundary or the shared physical destination of an already projected empty run.</summary>
    private void ConsumePagedForcedBreak(HtmlRenderFlowBlock block, HtmlRenderForcedBreak forcedBreak,
        ISet<object> consumedBoundaries) {
        consumedBoundaries.Add(forcedBreak.Boundary);
        if (!forcedBreak.Projected) return;
        // Projection has already merged an initial empty run into one physical
        // move. Its coincident destination markers must not create new pages.
        ChargeLayoutOperations(block.ForcedBreaks.Count, "projected physical boundary consumption");
        foreach (HtmlRenderForcedBreak item in block.ForcedBreaks) {
            if (item.Offset < forcedBreak.Offset - 0.0001D) continue;
            if (item.Offset > forcedBreak.Offset + 0.0001D) break;
            if (item.Projected) consumedBoundaries.Add(item.Boundary);
        }
    }

    private HtmlRenderForcedBreak? ResolveForcedBoundaryAt(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, double offset) {
        if (forcedBreaks.Count == 0) return null;
        CheckCancellation();
        ChargeLayoutOperations(3L * forcedBreaks.Count, "coincident forced page boundary resolution");
        int selected = -1;
        int rangeStart = -1;
        int rangeEnd = forcedBreaks.Count;
        for (int index = 0; index < forcedBreaks.Count; index++) {
            HtmlRenderForcedBreak forcedBreak = forcedBreaks[index];
            if (forcedBreak.Offset < offset - 0.0001D) continue;
            if (forcedBreak.Offset > offset + 0.0001D) { rangeEnd = index; break; }
            if (rangeStart < 0) rangeStart = index;
            if (IsForcedPageBreak(forcedBreak.Target)) selected = index;
        }
        if (selected < 0) return null;
        // Distinct empty siblings at one offset remain separate boundaries.
        // Fold only the selected boundary, starting at its first occurrence.
        for (int index = rangeStart; index < selected; index++) {
            if (IsForcedPageBreak(forcedBreaks[index].Target)
                && ReferenceEquals(forcedBreaks[index].Boundary, forcedBreaks[selected].Boundary)) {
                selected = index;
                break;
            }
        }
        return FoldForcedPageBoundary(forcedBreaks, selected, rangeEnd);
    }

    private static string? ResolvePageNameAt(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, double offset,
        string? currentPageName, object? boundary = null) {
        string? pageName = currentPageName;
        foreach (HtmlRenderForcedBreak forcedBreak in forcedBreaks) {
            if (forcedBreak.Offset < offset - 0.0001D) continue;
            if (forcedBreak.Offset > offset + 0.0001D) break;
            if (forcedBreak.ChangesPageName && (boundary == null || ReferenceEquals(boundary, forcedBreak.Boundary)))
                pageName = forcedBreak.PageName;
        }
        return pageName;
    }
}
