namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Joins projected column sets and full-width source blocks without changing
    /// their source coordinates. Qualified direct spanners retain page breaks
    /// through the ordinary sibling-boundary owner, including internal breaks.
    /// Column-only controls have no enclosing column target outside their set.
    /// </summary>
    private HtmlRenderFlowBlock BuildSpanningFlowBlock(IReadOnlyList<HtmlRenderFlowBlock> pieces,
        double width, string source, bool directSpanners) {
        var visuals = new List<HtmlRenderVisual>();
        var offsets = new SortedSet<double>();
        var safe = new List<HtmlRenderBreakRange>();
        var lines = new List<HtmlRenderLineBreakGroup>();
        var keeps = new List<HtmlRenderAvoidBreakRange>();
        var forced = new List<HtmlRenderForcedBreak>();
        var spannerForced = new List<HtmlRenderForcedBreak>();
        var assignments = new List<HtmlCssRunningStringAssignment>();
        object[] boundaries = CreateFlowBoundaries(pieces);
        double height = 0D;
        double paintExtent = 0D;
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
                spannerForced.Clear();
                AppendChildForcedBreaks(pieces, index, height, spannerForced, boundaries);
                forced.AddRange(spannerForced.Where(item => IsForcedPageBreak(item.Target)));
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
            paintExtent = Math.Max(paintExtent, height + piece.PagedPaintExtent);
            height += piece.Height;
        }
        return new HtmlRenderFlowBlock(width, height, visuals, HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None, false, source, offsets, lineBreakGroups: lines,
            forcedBreaks: forced, avoidBreakRanges: keeps, safeBreakRanges: safe,
            runningStringAssignments: assignments, pagedPaintExtent: Math.Max(height, paintExtent));
    }

    private bool IsSpanningSourceBlock(HtmlRenderFlowBlock block) => block.OwnerElement != null
        && _layoutStyles.TryGetValue(block.OwnerElement, out HtmlRenderBoxStyle? style)
        && style.ColumnSpan == "all";

    /// <summary>
    /// Uses the canonical fragment planner to locate the terminal set after
    /// natural fragmentation, forced page jumps and fitting page keeps. The
    /// preview includes its first source block so a spanner's keep-after can move
    /// the spanner too, and a completed fragment cannot donate page slack to it.
    /// </summary>
    private bool TryResolveSpanningPageOrigin(HtmlRenderFlowBlock prefix, double terminalStart,
        ColumnPaginationContext owner, out ColumnPaginationContext terminal) {
        terminal = owner;
        // The retained column mapping requires one stable page name and body
        // geometry. Internal page breaks must not silently qualify named pages.
        if (prefix.ForcedBreaks.Any(item => item.ChangesPageName)) return false;
        double start = 0D;
        double remaining = owner.Boundary.RemainingHeight;
        int page = owner.PageNumber;
        var consumedBoundaries = new HashSet<object>();
        while (start < terminalStart - 0.0001D
            || TryGetNextForcedBreak(prefix.ForcedBreaks, start, out HtmlRenderForcedBreak? pending,
                consumedBoundaries, includeProjectedBoundaries: true)
                && pending!.Offset <= terminalStart + 0.0001D) {
            CheckCancellation();
            ChargeLayoutOperation("spanner physical page origin");
            PagedBlockFragment? fragment = PlanPagedFragment(prefix, start, remaining, owner.PageHeight,
                remaining >= owner.PageHeight - 0.0001D, consumedBoundaries, includeProjectedBoundaries: true);
            if (fragment != null && (fragment.Forced || fragment.SuppressedHeader || fragment.SuppressedFooter)) return false;
            if (fragment != null && fragment.End > terminalStart + 0.0001D) {
                remaining -= terminalStart - start;
                break;
            } else if (fragment != null) {
                // The outer paginator commits every unfinished source fragment,
                // including one ending at terminalStart with unused body space.
                start = fragment.End;
                if (fragment.AuthoredBreak != null) ConsumePagedForcedBreak(prefix, fragment.AuthoredBreak, consumedBoundaries);
            }
            HtmlPageBreakTarget target = fragment?.AuthoredBreak?.Target
                ?? ResolveForcedBreakAt(prefix.ForcedBreaks, start);
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
