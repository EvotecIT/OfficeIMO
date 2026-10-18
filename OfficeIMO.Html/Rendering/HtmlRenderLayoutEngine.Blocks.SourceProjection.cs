using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HashSet<IElement> _constrainedOverflowElements = new();
    private readonly HashSet<IElement> _projectedConstrainedOwners = new();
    private ConstrainedSourceProjection? _constrainedSourceProjection;

    private void ResetConstrainedSourceProjection() {
        _constrainedOverflowElements.Clear();
        _projectedConstrainedOwners.Clear();
        _constrainedSourceProjection = null;
    }

    private void RememberConstrainedOverflow(IElement element, HtmlRenderBoxStyle style, HtmlRenderFlowBlock block) {
        if (_constrainedSourceProjection == null && _options.Mode == HtmlRenderMode.Paged
            && HasConstrainedBlockHeight(style) && style.OverflowY == "visible"
            && (block.PagedPaintExtent > block.Height + 0.0001D
                || block.Visuals.Any(ContainsInFlowPaint)))
            _constrainedOverflowElements.Add(element);
    }

    /// <summary>Rebuilds independent ordinary sources against the physical page bands before flattening paint.</summary>
    private bool TryRelayoutConstrainedOverflow(HtmlRenderFlowBlock block, double available,
        double pageHeight, int pageNumber, HtmlCssPageGeometry geometry, out HtmlRenderFlowBlock reflowed) {
        reflowed = block;
        IElement? owner = block.OwnerElement;
        if (owner == null || available <= 0.0001D || pageHeight <= 0D
            || block.Height <= available + 0.0001D || _projectedConstrainedOwners.Contains(owner)) return false;
        IElement[] targets = _constrainedOverflowElements.Where(element => ContainsElementOrSelf(owner, element)).ToArray();
        if (targets.Length == 0) return false;
        IElement root = _document.Body ?? _document.DocumentElement ?? owner;
        bool isRoot = ReferenceEquals(owner, root);
        if (!isRoot && !ReferenceEquals(owner.ParentElement, root)) return false;
        _projectedConstrainedOwners.Add(owner);
        string? boundary = ResolveConstrainedProjectionBoundary(owner, targets);
        var context = new ConstrainedSourceProjection(targets, geometry, pageNumber, pageHeight,
            block.PageName, boundary);
        double leading = isRoot ? 0D : block.LeadingFlowAdjustment;
        HtmlRenderBoxStyle rootStyle = ResolveRootStyle(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = isRoot ? rootStyle : _styleResolver.Resolve(owner, geometry.ContentWidth, rootStyle);
        ChargeLayoutOperations(_normalFlowPlacements.Count + _layoutStyles.Count + _inlineFloatOverhangs.Count,
            "constrained projection fallback state");
        KeyValuePair<IElement, NormalFlowPlacement>[] placements = _normalFlowPlacements.ToArray();
        KeyValuePair<IElement, HtmlRenderBoxStyle>[] styles = _layoutStyles.ToArray();
        KeyValuePair<IElement, double>[] floatOverhangs = _inlineFloatOverhangs.ToArray();
        int semanticOrder = _nextSemanticNodeId;
        int logicalOrder = _nextLogicalTextOrder;
        bool accepted = false;
        _constrainedSourceProjection = context;
        try {
            var physical = new PagedFloatBoundary(available + leading, pageHeight);
            HtmlRenderFlowBlock candidate = isRoot
                ? LayoutRootElement(root, geometry.ContentWidth, rootStyle, pageBoundary: physical)
                : LayoutElement(owner, geometry.ContentWidth, style, rootStyle, 1, pageBoundary: physical)
                    .AdjustLeadingFlowSpace(leading);
            if (context.FailedReason != null) {
                if (context.NeedsProjection) _diagnostics.Add(ComponentName,
                    HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending,
                    "Overlapping constrained content retained its shared source pagination outside the supported independent ordinary-flow projection.",
                    HtmlDiagnosticSeverity.Warning, block.Source, context.FailedReason);
                return false;
            }
            if (!context.Changed) return false;
            reflowed = candidate;
            accepted = true;
            return true;
        } finally {
            _constrainedSourceProjection = null;
            if (!accepted) {
                _normalFlowPlacements.Clear();
                foreach (KeyValuePair<IElement, NormalFlowPlacement> item in placements) {
                    CheckCancellation();
                    _normalFlowPlacements.Add(item.Key, item.Value);
                }
                _layoutStyles.Clear();
                foreach (KeyValuePair<IElement, HtmlRenderBoxStyle> item in styles) {
                    CheckCancellation();
                    _layoutStyles.Add(item.Key, item.Value);
                }
                _inlineFloatOverhangs.Clear();
                foreach (KeyValuePair<IElement, double> item in floatOverhangs) {
                    CheckCancellation();
                    _inlineFloatOverhangs.Add(item.Key, item.Value);
                }
                _nextSemanticNodeId = semanticOrder;
                _nextLogicalTextOrder = logicalOrder;
            }
        }
    }

    private string? ResolveConstrainedProjectionBoundary(IElement owner, IReadOnlyList<IElement> targets) {
        foreach (IElement target in targets) {
            CheckCancellation();
            ChargeLayoutOperations(targets.Count, "constrained projection ownership");
            if (targets.Any(other => !ReferenceEquals(other, target) && ContainsElementOrSelf(target, other)))
                return "nested-independent-constrained-overflow";
        }
        foreach (KeyValuePair<IElement, HtmlRenderBoxStyle> item in _layoutStyles) {
            CheckCancellation();
            ChargeLayoutOperation("constrained projection formatting boundary");
            if (!ContainsElementOrSelf(owner, item.Key)) continue;
            HtmlRenderBoxStyle style = item.Value;
            if (style.WritingMode != "horizontal-tb" || HasMultiColumnLayout(style)
                || style.Display is "flex" or "inline-flex" or "grid" or "inline-grid" or "table" or "inline-table"
                || style.FloatSide != "none" || ShouldExtractOutOfFlow(style)
                || style.UsedTransform != "none" || style.UsedIndividualScale != "none")
                return "independent-formatting-or-positioned-source-owner";
        }
        return null;
    }

    private ConstrainedProjectionScope EnterConstrainedProjection(IElement element) =>
        new(_constrainedSourceProjection, element);

    /// <summary>Uses the canonical fragment planner; inserted gaps occupy destination space, not extra CSS height.</summary>
    private HtmlRenderFlowBlock ProjectOrdinaryContent(HtmlRenderFlowBlock source, PagedFloatBoundary boundary,
        out double entryGap) {
        entryGap = 0D;
        ConstrainedSourceProjection? context = _constrainedSourceProjection;
        if (context == null || context.FailedReason != null || source.PagedPaintExtent <= 0.0001D) return source;
        double available = boundary.RemainingHeight;
        double pageOffset = 0D;
        if (available <= 0.0001D) {
            pageOffset = Math.Floor(-available / context.PageHeight) + 1D;
            available += pageOffset * context.PageHeight;
        }
        if (pageOffset < 0D || pageOffset >= _options.MaxPageCount) {
            throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
        }
        int pageNumber = context.PageNumber + (int)pageOffset;
        HtmlRenderFlowBlock projected = source;
        double cursor = 0D;
        while (cursor < projected.PagedPaintExtent - 0.0001D) {
            CheckCancellation();
            ChargeLayoutOperation("constrained ordinary source fragment");
            if (pageNumber > _options.MaxPageCount)
                throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
            HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(pageNumber, context.PageName, _options);
            if (!SamePageGeometry(context.Geometry, geometry)
                || Math.Abs(context.PageHeight - ResolvePageBodyContentHeight(pageNumber, geometry)) > 0.0001D) {
                context.FailedReason = "different-page-geometry-or-reserved-body-space";
                return source;
            }
            double remaining = projected.PagedPaintExtent - cursor;
            if (remaining <= available + 0.0001D) break;
            HtmlRenderFlowBlock paged = projected.ForPagination();
            bool movesOwnKeep = cursor <= 0.0001D && source.AvoidBreakInside
                && source.PagedPaintExtent <= context.PageHeight + 0.0001D
                && remaining > available + 0.0001D;
            PagedBlockFragment? plan = movesOwnKeep ? null
                : PlanPagedFragment(paged, cursor, available, context.PageHeight,
                    canForceFragment: available >= context.PageHeight - 0.0001D);
            if (plan?.Forced == true || plan?.Header != null || plan?.Footer != null
                || HasInternalForcedBreak(projected)) {
                context.FailedReason = "forced-or-repeated-source-fragment";
                return source;
            }
            double cut = plan?.End ?? cursor;
            double gap = available - (cut - cursor);
            if (gap > 0.0001D) {
                context.NeedsProjection = true;
                if (context.UnsupportedBoundary != null) {
                    context.FailedReason = context.UnsupportedBoundary;
                    return source;
                }
                if (cursor <= 0.0001D && cut <= 0.0001D) {
                    // No content entered this box. The gap belongs before the
                    // complete CSS box and therefore advances its following flow.
                    entryGap += gap;
                } else {
                    double advance = cut < projected.Height - 0.0001D ? gap : 0D;
                    projected = InsertPagedSourceBreakGap(projected, cut, gap, false, advance);
                    cut += gap;
                }
                context.Changed = true;
            }
            cursor = cut;
            available = context.PageHeight;
            pageNumber++;
        }
        return projected;
    }

    private sealed class ConstrainedSourceProjection {
        internal ConstrainedSourceProjection(IEnumerable<IElement> targets, HtmlCssPageGeometry geometry,
            int pageNumber, double pageHeight, string? pageName, string? unsupportedBoundary) {
            Targets = new HashSet<IElement>(targets);
            Geometry = geometry;
            PageNumber = pageNumber;
            PageHeight = pageHeight;
            PageName = pageName;
            UnsupportedBoundary = unsupportedBoundary;
        }
        internal HashSet<IElement> Targets { get; }
        internal HtmlCssPageGeometry Geometry { get; }
        internal int PageNumber { get; }
        internal double PageHeight { get; }
        internal string? PageName { get; }
        internal string? UnsupportedBoundary { get; }
        internal int ConstrainedDepth { get; set; }
        internal bool Changed { get; set; }
        internal bool NeedsProjection { get; set; }
        internal string? FailedReason { get; set; }
    }

    private sealed class ConstrainedProjectionScope : IDisposable {
        private readonly ConstrainedSourceProjection? _context;
        private readonly bool _containsSource;
        internal ConstrainedProjectionScope(ConstrainedSourceProjection? context, IElement element) {
            _context = context;
            CanProject = context != null && context.ConstrainedDepth == 0;
            _containsSource = context?.Targets.Contains(element) == true;
            if (_containsSource) context!.ConstrainedDepth++;
        }
        internal bool CanProject { get; }
        public void Dispose() {
            if (_containsSource) _context!.ConstrainedDepth--;
        }
    }
}
