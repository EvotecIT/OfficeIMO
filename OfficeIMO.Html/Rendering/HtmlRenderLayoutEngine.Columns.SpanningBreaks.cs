namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool HasSpannerPageControls(HtmlRenderFlowBlock block, HtmlRenderBoxStyle style) =>
        IsForcedPageBreak(block.BreakBefore) || IsForcedPageBreak(block.BreakAfter)
        || ((style.AvoidBreakBeforeScope | style.AvoidBreakAfterScope | style.AvoidBreakInsideScope)
            & HtmlFragmentationScope.Page) != 0;

    /// <summary>
    /// Joins projected column sets and full-width source blocks without changing
    /// their source coordinates. Only the direct spanner's own page controls are
    /// promoted; its internal forced-break fallback remains diagnosed separately.
    /// </summary>
    private HtmlRenderFlowBlock BuildSpanningFlowBlock(IReadOnlyList<HtmlRenderFlowBlock> pieces,
        double width, string source, bool directSpanners) {
        var visuals = new List<HtmlRenderVisual>();
        var offsets = new SortedSet<double>();
        var safe = new List<HtmlRenderBreakRange>();
        var lines = new List<HtmlRenderLineBreakGroup>();
        var keeps = new List<HtmlRenderAvoidBreakRange>();
        var forced = new List<HtmlRenderForcedBreak>();
        var assignments = new List<HtmlCssRunningStringAssignment>();
        object[] boundaries = CreateFlowBoundaries(pieces);
        double height = 0D;
        for (int index = 0; index < pieces.Count; index++) {
            CheckCancellation();
            HtmlRenderFlowBlock piece = pieces[index];
            ChargeLayoutOperations(1L + piece.Visuals.Count + piece.BreakOffsets.Count
                + piece.SafeBreakRanges.Count + piece.RunningStringAssignments.Count,
                "spanner-delimited source projection");
            bool spanner = IsSpanningSourceBlock(piece);
            if (directSpanners && index > 0 && (spanner || IsSpanningSourceBlock(pieces[index - 1]))
                && (ResolveAvoidBreakBetweenScope(pieces[index - 1], piece) & HtmlFragmentationScope.Page) != 0) {
                HtmlRenderFlowBlock previous = pieces[index - 1];
                AppendKeepWithNextRange(pieces, index, height, keeps,
                    height - previous.Height + FinalKeptContentStart(previous, HtmlFragmentationScope.Page));
            }
            if (!spanner) {
                AppendChildForcedBreaks(pieces, index, height, forced, boundaries);
            } else if (directSpanners) {
                if (IsForcedPageBreak(piece.BreakBefore))
                    forced.Add(new HtmlRenderForcedBreak(height, piece.BreakBefore, boundary: boundaries[index],
                        edge: index == 0 ? HtmlRenderBreakEdge.Leading : HtmlRenderBreakEdge.Interior));
                if (IsForcedPageBreak(piece.BreakAfter))
                    forced.Add(new HtmlRenderForcedBreak(height + piece.Height, piece.BreakAfter,
                        boundary: boundaries[index + 1],
                        edge: index == pieces.Count - 1 ? HtmlRenderBreakEdge.Trailing : HtmlRenderBreakEdge.Interior));
            }
            foreach (HtmlRenderVisual visual in piece.Visuals)
                visuals.Add(visual.Translate(0D, height, visuals.Count));
            foreach (double offset in piece.BreakOffsets) offsets.Add(height + offset);
            safe.AddRange(piece.SafeBreakRanges.Select(range => range.Translate(height)));
            if (spanner) {
                lines.AddRange(SliceColumnLineBreakGroups(piece.LineBreakGroups, 0D, piece.Height, height));
                keeps.AddRange(SliceColumnAvoidBreakRanges(piece, 0D, piece.Height, height));
            } else {
                lines.AddRange(piece.LineBreakGroups.Select(group => group.Translate(height)));
                keeps.AddRange(piece.AvoidBreakRanges.Select(range => range.Translate(height)));
            }
            assignments.AddRange(piece.RunningStringAssignments.Select(item => item.Translate(height)));
            height += piece.Height;
        }
        return new HtmlRenderFlowBlock(width, height, visuals, HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None, false, source, offsets, lineBreakGroups: lines,
            forcedBreaks: forced, avoidBreakRanges: keeps, safeBreakRanges: safe,
            runningStringAssignments: assignments);
    }

    private bool IsSpanningSourceBlock(HtmlRenderFlowBlock block) => block.OwnerElement != null
        && _layoutStyles.TryGetValue(block.OwnerElement, out HtmlRenderBoxStyle? style)
        && style.ColumnSpan == "all";

    /// <summary>
    /// Uses the canonical fragment planner to locate the terminal set after
    /// forced page jumps and fitting page keeps. The preview includes its first
    /// source block so a spanner's keep-after can move the spanner too.
    /// </summary>
    private bool TryResolveSpanningPageOrigin(HtmlRenderFlowBlock prefix, double terminalStart,
        ColumnPaginationContext owner, out ColumnPaginationContext terminal) {
        terminal = owner;
        double start = 0D;
        double remaining = owner.Boundary.RemainingHeight;
        int page = owner.PageNumber;
        while (start < terminalStart - 0.0001D) {
            CheckCancellation();
            ChargeLayoutOperation("spanner physical page origin");
            PagedBlockFragment? fragment = PlanPagedFragment(prefix, start, remaining, owner.PageHeight,
                remaining >= owner.PageHeight - 0.0001D);
            if (fragment != null && (fragment.Forced || fragment.SuppressedHeader || fragment.SuppressedFooter)) return false;
            if (fragment != null && fragment.End >= terminalStart - 0.0001D) {
                remaining -= terminalStart - start;
                start = terminalStart;
                if (remaining > 0.0001D
                    && ResolveForcedBreakAt(prefix.ForcedBreaks, terminalStart) == HtmlPageBreakTarget.None) break;
            } else if (fragment != null) {
                start = fragment.End;
            }
            HtmlPageBreakTarget target = ResolveForcedBreakAt(prefix.ForcedBreaks, start);
            do {
                CheckCancellation();
                ChargeLayoutOperation("spanner physical page advance");
                if (page >= _options.MaxPageCount)
                    throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
                page++;
                HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(page, owner.PageName, _options);
                if (!SamePageGeometry(owner.Geometry, geometry)
                    || Math.Abs(owner.PageHeight - ResolvePageBodyContentHeight(page, geometry)) > 0.0001D) return false;
            } while (target == HtmlPageBreakTarget.Left && page % 2 != 0
                || target == HtmlPageBreakTarget.Right && page % 2 == 0);
            remaining = owner.PageHeight;
        }
        terminal = owner with {
            Boundary = new PagedFloatBoundary(remaining, owner.PageHeight),
            PageNumber = page
        };
        return true;
    }
}
