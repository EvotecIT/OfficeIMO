using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, HtmlRenderFlowBlock> _parallelTableRegions = new();
    private readonly Dictionary<IElement, ParallelTableProjectionContext> _parallelTableProjections = new();
    private readonly HashSet<IElement> _rejectedParallelTableProjections = new();
    private ParallelTableProjectionAttempt? _parallelTableProjectionAttempt;

    private void ResetParallelTablePagination() {
        _parallelTableRegions.Clear();
        _parallelTableProjections.Clear();
        _rejectedParallelTableProjections.Clear();
        _parallelTableProjectionAttempt = null;
    }

    private bool CanProjectParallelTableItems(IElement element, HtmlRenderBoxStyle style,
        IReadOnlyList<HtmlRenderFlowBlock> blocks, IEnumerable<HtmlRenderBoxStyle> itemStyles) {
        if (_options.Mode != HtmlRenderMode.Paged || blocks.Count < 2 || HasBoundedCrossSize(style)
            || style.WritingMode != "horizontal-tb" || style.OverflowY != "visible"
            || style.UsedTransform != "none" || style.UsedIndividualScale != "none"
            || !blocks.Any(item => item.ContinuationGroups.Count > 0 || item.TrailingGroups.Count > 0)) return false;
        foreach (HtmlRenderBoxStyle item in itemStyles) {
            CheckCancellation();
            ChargeLayoutOperation("parallel table item qualification");
            if (HasBoundedCrossSize(item) || item.WritingMode != "horizontal-tb"
                || item.OverflowY != "visible" || item.UsedTransform != "none"
                || item.UsedIndividualScale != "none" || item.FloatSide != "none") return false;
        }
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (_layoutStyles.TryGetValue(parent, out HtmlRenderBoxStyle? ancestor)
                && (HasBoundedCrossSize(ancestor) || ancestor.OverflowY != "visible"
                    || ancestor.UsedTransform != "none" || ancestor.UsedIndividualScale != "none")) return false;
        }
        // Nested parallel regions retain their existing diagnosed contract. One
        // projection must not reinterpret another region's virtual page bands.
        if (_parallelTableRegions.Keys.Any(child => !ReferenceEquals(child, element)
            && ContainsElementOrSelf(element, child))) return false;
        return blocks.All(item => item.ForcedBreaks.Count == 0
            && Math.Abs(item.PagedPaintExtent - item.Height) <= 0.0001D);
    }

    private bool TryProjectParallelTableItems(IElement element, double width,
        IReadOnlyList<HtmlRenderFlowBlock> blocks, PagedFloatBoundary? boundary,
        out IReadOnlyList<HtmlRenderFlowBlock> projected) {
        projected = blocks;
        bool retained = _parallelTableProjections.TryGetValue(element, out ParallelTableProjectionContext? context);
        if (!retained) {
            ParallelTableProjectionAttempt? attempt = _parallelTableProjectionAttempt;
            if (attempt == null || !ReferenceEquals(attempt.Element, element) || !boundary.HasValue
                || boundary.Value.RemainingHeight <= 0.0001D || boundary.Value.FragmentStart > 0.0001D) return false;
            context = new ParallelTableProjectionContext(boundary.Value, attempt.PageNumber,
                attempt.Geometry, attempt.PageHeight, width, attempt.PageName);
        }
        if (Math.Abs(context!.Width - width) > 0.0001D
            || !TryProjectParallelTableBands(blocks, context, out projected)) return false;
        if (!retained) _parallelTableProjectionAttempt!.Projection = context;
        return true;
    }

    private bool TryRelayoutBlockForParallelTables(HtmlRenderFlowBlock block, double offset,
        double available, double pageHeight, int pageNumber, HtmlCssPageGeometry geometry,
        out HtmlRenderFlowBlock reflowed) {
        reflowed = block;
        if (block.OwnerElement == null || HasInternalForcedBreak(block) || available <= 0.0001D) return false;
        IElement owner = block.OwnerElement;
        IElement? target = null;
        foreach (KeyValuePair<IElement, HtmlRenderFlowBlock> entry in _parallelTableRegions) {
            CheckCancellation();
            ChargeLayoutOperation("parallel table boundary discovery");
            if (_parallelTableProjections.ContainsKey(entry.Key) || _rejectedParallelTableProjections.Contains(entry.Key)
                || !ContainsElementOrSelf(owner, entry.Key)
                || !_layoutStyles.TryGetValue(entry.Key, out HtmlRenderBoxStyle? entryStyle)) continue;
            double start;
            if (ReferenceEquals(owner, entry.Key)) start = 0D;
            else if (TryResolveContentOrigin(entry.Key, owner, out PositionedPoint origin))
                start = origin.Y - entryStyle.MarginTop - entryStyle.BorderTopWidth - entryStyle.PaddingTop;
            else continue;
            if (start < offset - 0.0001D || start >= offset + available - 0.0001D
                || start + entry.Value.Height <= offset + available + 0.0001D) continue;
            target = entry.Key;
            break;
        }
        if (target == null) return false;
        // Rebuilding a root must retain earlier projections under the geometry
        // that established their source-to-destination mapping.
        if (_parallelTableProjections.Any(entry => ContainsElementOrSelf(owner, entry.Key)
            && !SamePageGeometry(entry.Value.Geometry, geometry))) return false;
        IElement root = _document.Body ?? _document.DocumentElement ?? owner;
        bool isRoot = ReferenceEquals(owner, root);
        if (!isRoot && !ReferenceEquals(owner.ParentElement, root)) return false;
        HtmlRenderBoxStyle rootStyle = ResolveRootStyle(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = isRoot ? rootStyle : _styleResolver.Resolve(owner, geometry.ContentWidth, rootStyle);
        double leading = isRoot ? 0D : block.LeadingFlowAdjustment;
        var boundary = new PagedFloatBoundary(offset + available + leading, pageHeight, offset + leading);
        var attempt = new ParallelTableProjectionAttempt(target, pageNumber, geometry, pageHeight, block.PageName);
        _parallelTableProjectionAttempt = attempt;
        try {
            HtmlRenderFlowBlock candidate = isRoot
                ? LayoutRootElement(root, geometry.ContentWidth, rootStyle, pageBoundary: boundary)
                : LayoutElement(owner, geometry.ContentWidth, style, rootStyle, 1, pageBoundary: boundary)
                    .AdjustLeadingFlowSpace(leading);
            if (attempt.Projection == null) {
                _rejectedParallelTableProjections.Add(target);
                return false;
            }
            _parallelTableProjections.Add(target, attempt.Projection);
            reflowed = candidate;
            return true;
        } finally {
            _parallelTableProjectionAttempt = null;
        }
    }

    private sealed record ParallelTableProjectionContext(PagedFloatBoundary Boundary, int PageNumber,
        HtmlCssPageGeometry Geometry, double PageHeight, double Width, string? PageName);

    private sealed class ParallelTableProjectionAttempt {
        internal ParallelTableProjectionAttempt(IElement element, int pageNumber, HtmlCssPageGeometry geometry, double pageHeight, string? pageName) {
            Element = element;
            PageNumber = pageNumber;
            Geometry = geometry;
            PageHeight = pageHeight;
            PageName = pageName;
        }
        internal IElement Element { get; }
        internal int PageNumber { get; }
        internal HtmlCssPageGeometry Geometry { get; }
        internal double PageHeight { get; }
        internal string? PageName { get; }
        internal ParallelTableProjectionContext? Projection { get; set; }
    }
}
