using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, HtmlRenderFlowBlock> _pagedColumnRegions = new();
    private readonly Dictionary<IElement, ColumnPaginationContext> _pagedColumnContexts = new();
    private readonly HashSet<IElement> _retainedSourceColumnRegions = new();
    private readonly HashSet<IElement> _deferredColumnWrappers = new();
    private ColumnPaginationAttempt? _columnPaginationAttempt;

    private void ResetColumnPagination() {
        _pagedColumnRegions.Clear();
        _pagedColumnContexts.Clear();
        _retainedSourceColumnRegions.Clear();
        _deferredColumnWrappers.Clear();
        _columnPaginationAttempt = null;
    }

    /// <summary>
    /// Rebuilds the owning flow when an unconsumed column region first reaches a
    /// physical page boundary. Completed placements supply its actual origin,
    /// including collapsed margins and formatting-container offsets.
    /// </summary>
    private bool TryRelayoutBlockForColumns(HtmlRenderFlowBlock block, double offset,
        double available, double pageHeight, int pageNumber, HtmlCssPageGeometry geometry,
        out HtmlRenderFlowBlock reflowed) {
        reflowed = block;
        if (block.OwnerElement == null || available <= 0.0001D || pageHeight <= 0D) return false;
        IElement owner = block.OwnerElement;
        IElement? target = null;
        double targetContentStart = 0D;
        double targetStart = double.PositiveInfinity;
        double unfragmentedHeight = 0D;
        foreach (KeyValuePair<IElement, HtmlRenderFlowBlock> entry in _pagedColumnRegions) {
            CheckCancellation();
            ChargeLayoutOperation("column page-boundary discovery");
            if (_pagedColumnContexts.ContainsKey(entry.Key) || _retainedSourceColumnRegions.Contains(entry.Key)
                || !ContainsElementOrSelf(owner, entry.Key)
                || !_layoutStyles.TryGetValue(entry.Key, out HtmlRenderBoxStyle? entryStyle)) continue;
            double start;
            double contentStart;
            if (ReferenceEquals(owner, entry.Key)) {
                start = 0D;
                contentStart = entryStyle.MarginTop + entryStyle.BorderTopWidth + entryStyle.PaddingTop;
            } else if (TryResolveContentOrigin(entry.Key, owner, out PositionedPoint origin)) {
                contentStart = origin.Y;
                start = contentStart - entryStyle.MarginTop - entryStyle.BorderTopWidth - entryStyle.PaddingTop;
            } else continue;
            if (start < offset - 0.0001D || start >= offset + available - 0.0001D
                || ResolveDeclaredColumnContentHeight(entryStyle).HasValue
                    && start + entry.Value.PagedPaintExtent <= offset + available + 0.0001D) continue;
            // Let an earlier authored page or page-name break establish the
            // destination before capturing this region's pagination context.
            // Synthetic overflow breaks inside the region remain eligible.
            if (TryGetNextForcedBreak(block.ForcedBreaks, offset, out HtmlRenderForcedBreak? precedingBreak)
                && precedingBreak!.Offset <= start + 0.0001D) continue;
            // CSS order and grid placement need not follow registration order.
            // Project earlier flows before capturing a later flow's capacity.
            if (start >= targetStart) continue;
            target = entry.Key;
            targetContentStart = contentStart;
            targetStart = start;
            unfragmentedHeight = entry.Value.UnfragmentedHeight ?? entry.Value.Height;
        }
        if (target == null) return false;

        // Earlier column plans retain the page geometry that established their
        // source-to-destination mapping when their ancestor is rebuilt.
        if (_pagedColumnContexts.Any(entry => ContainsElementOrSelf(owner, entry.Key)
            && !SamePageGeometry(entry.Value.Geometry, geometry))) return false;
        IElement root = _document.Body ?? _document.DocumentElement ?? owner;
        bool isRoot = ReferenceEquals(owner, root);
        if (!isRoot && !ReferenceEquals(owner.ParentElement, root)) return false;
        HtmlRenderBoxStyle rootStyle = ResolveRootStyle(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = isRoot ? rootStyle : _styleResolver.Resolve(owner, geometry.ContentWidth, rootStyle);
        double leading = isRoot ? 0D : block.LeadingFlowAdjustment;
        var boundary = new PagedFloatBoundary(offset + available + leading, pageHeight, offset + leading);
        PagedFloatBoundary columnBoundary = boundary.Shift(targetContentStart);
        int columnPageNumber = pageNumber;
        if (TryDeferLeadingColumnWrapper(target, owner, offset, columnBoundary, pageNumber, geometry,
                block.PageName, out double freshContentHeight)) {
            columnBoundary = new PagedFloatBoundary(freshContentHeight, pageHeight);
            columnPageNumber++;
        }
        var attempt = new ColumnPaginationAttempt(target, columnPageNumber, geometry, pageHeight,
            block.PageName, columnBoundary, unfragmentedHeight);
        _columnPaginationAttempt = attempt;
        try {
            HtmlRenderFlowBlock candidate = isRoot
                ? LayoutRootElement(root, geometry.ContentWidth, rootStyle, pageBoundary: boundary)
                : LayoutElement(owner, geometry.ContentWidth, style, rootStyle, 1, pageBoundary: boundary)
                    .AdjustLeadingFlowSpace(leading);
            // A short fixed box or an unsupported geometry retains source
            // layout. Keep the rebuilt owner and continue discovery so that
            // this fallback cannot hide an eligible parallel sibling. Each
            // retained region is attempted only once in this layout pass.
            if (attempt.Context == null) _retainedSourceColumnRegions.Add(target);
            else _pagedColumnContexts.Add(target, attempt.Context);
            reflowed = candidate;
            return true;
        } finally {
            _columnPaginationAttempt = null;
        }
    }

    /// <summary>
    /// Carries an unconsumed ordinary wrapper's leading decoration with its first
    /// column content, instead of leaving that decoration before a synthetic break.
    /// </summary>
    private bool TryDeferLeadingColumnWrapper(IElement target, IElement owner, double consumed,
        PagedFloatBoundary boundary, int pageNumber, HtmlCssPageGeometry geometry, string? pageName,
        out double freshContentHeight) {
        freshContentHeight = 0D;
        if (!_normalFlowPlacements.TryGetValue(target, out NormalFlowPlacement? placement)
            || Math.Abs(placement.Y) > 0.0001D
            || !_layoutStyles.TryGetValue(target, out HtmlRenderBoxStyle? targetStyle)
            || targetStyle.ExplicitHeight.HasValue
            || Math.Abs(targetStyle.MarginTop) > 0.0001D
            || targetStyle.BorderTopWidth + targetStyle.PaddingTop > 0.0001D
            || !_layoutStyles.TryGetValue(placement.Parent, out HtmlRenderBoxStyle? wrapper)
            || wrapper.Display is not ("block" or "flow-root")
            || wrapper.ExplicitHeight.HasValue && !wrapper.AutoHeightLayoutStretch
            || IsRootLayoutContainer(placement.Parent)
            || !TryResolveContentOrigin(placement.Parent, owner, out PositionedPoint origin)) return false;
        double inset = wrapper.BorderTopWidth + wrapper.PaddingTop;
        double start = origin.Y - wrapper.MarginTop - inset;
        freshContentHeight = boundary.PageHeight - inset;
        if (start < consumed - 0.0001D || inset <= 0.0001D && Math.Abs(wrapper.MarginTop) <= 0.0001D
            || freshContentHeight <= 0.0001D
            || !_pagedColumnRegions.TryGetValue(target, out HtmlRenderFlowBlock? source)
            || FindFragmentEnd(source, 0D, boundary.RemainingHeight, fullPageHeight: boundary.PageHeight) > 0.0001D
            || FindFragmentEnd(source, 0D, freshContentHeight, fullPageHeight: boundary.PageHeight) <= 0.0001D) return false;
        HtmlCssPageGeometry next = _pageRules.ResolveGeometry(pageNumber + 1, pageName, _options);
        if (!SamePageGeometry(geometry, next)
            || Math.Abs(ResolvePageBodyContentHeight(pageNumber + 1, next) - boundary.PageHeight) > 0.0001D) return false;
        _deferredColumnWrappers.Add(placement.Parent);
        return true;
    }

    private sealed record ColumnPaginationContext(PagedFloatBoundary Boundary, int PageNumber,
        HtmlCssPageGeometry Geometry, double PageHeight, double Width, string? PageName, double UnfragmentedHeight);

    private sealed class ColumnPaginationAttempt {
        internal ColumnPaginationAttempt(IElement element, int pageNumber, HtmlCssPageGeometry geometry,
            double pageHeight, string? pageName, PagedFloatBoundary boundary, double unfragmentedHeight) {
            Element = element;
            PageNumber = pageNumber;
            Geometry = geometry;
            PageHeight = pageHeight;
            PageName = pageName;
            Boundary = boundary;
            UnfragmentedHeight = unfragmentedHeight;
        }
        internal IElement Element { get; }
        internal int PageNumber { get; }
        internal HtmlCssPageGeometry Geometry { get; }
        internal double PageHeight { get; }
        internal string? PageName { get; }
        internal PagedFloatBoundary Boundary { get; }
        internal double UnfragmentedHeight { get; }
        internal ColumnPaginationContext? Context { get; set; }
    }
}
