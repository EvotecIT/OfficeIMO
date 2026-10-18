using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, HtmlRenderFlowBlock> _pagedColumnRegions = new();
    private readonly Dictionary<IElement, ColumnPaginationContext> _pagedColumnContexts = new();
    private readonly HashSet<IElement> _retainedSourceColumnRegions = new();
    private ColumnPaginationAttempt? _columnPaginationAttempt;

    private void ResetColumnPagination() {
        _pagedColumnRegions.Clear();
        _pagedColumnContexts.Clear();
        _retainedSourceColumnRegions.Clear();
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
                    && start + entry.Value.Height <= offset + available + 0.0001D) continue;
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
        var attempt = new ColumnPaginationAttempt(target, pageNumber, geometry, pageHeight,
            block.PageName, boundary.Shift(targetContentStart));
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

    private sealed record ColumnPaginationContext(PagedFloatBoundary Boundary, int PageNumber,
        HtmlCssPageGeometry Geometry, double PageHeight, double Width, string? PageName);

    private sealed class ColumnPaginationAttempt {
        internal ColumnPaginationAttempt(IElement element, int pageNumber, HtmlCssPageGeometry geometry,
            double pageHeight, string? pageName, PagedFloatBoundary boundary) {
            Element = element;
            PageNumber = pageNumber;
            Geometry = geometry;
            PageHeight = pageHeight;
            PageName = pageName;
            Boundary = boundary;
        }
        internal IElement Element { get; }
        internal int PageNumber { get; }
        internal HtmlCssPageGeometry Geometry { get; }
        internal double PageHeight { get; }
        internal string? PageName { get; }
        internal PagedFloatBoundary Boundary { get; }
        internal ColumnPaginationContext? Context { get; set; }
    }
}
