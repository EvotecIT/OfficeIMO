using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HashSet<IElement> _constrainedOverflowElements = new();
    private readonly HashSet<IElement> _constrainedPaintOverflowElements = new();
    private readonly HashSet<IElement> _projectedConstrainedOwners = new();
    private ConstrainedSourceProjection? _constrainedSourceProjection;
    private int _specializedPageSourceDepth;
    private int _tablePageSourceDepth = -1;

    private void ResetConstrainedSourceProjection() {
        _constrainedOverflowElements.Clear();
        _constrainedPaintOverflowElements.Clear();
        _projectedConstrainedOwners.Clear();
        _constrainedSourceProjection = null;
        _specializedPageSourceDepth = 0;
        _tablePageSourceDepth = -1;
    }

    private void RememberConstrainedOverflow(IElement element, bool ownsConstrainedSource, HtmlRenderFlowBlock block) {
        if (_constrainedSourceProjection != null || _options.Mode != HtmlRenderMode.Paged) return;
        ChargeLayoutOperation("constrained source registration");
        if (!ownsConstrainedSource) {
            _constrainedOverflowElements.Remove(element);
            _constrainedPaintOverflowElements.Remove(element);
            return;
        }
        _constrainedOverflowElements.Add(element);
        if (block.PagedPaintExtent > block.Height + 0.0001D) _constrainedPaintOverflowElements.Add(element);
        else _constrainedPaintOverflowElements.Remove(element);
    }

    private static bool HasOrdinaryFragmentSource(HtmlRenderVisual visual) => !visual.IsOutOfFlowPaint
        && (visual is HtmlRenderText or HtmlRenderImage or HtmlRenderDrawing or HtmlRenderFormField
            or HtmlRenderShape { IsAtomicReplacedPlaceholder: true }
            || GetGroupChildren(visual) is { } children && children.Any(HasOrdinaryFragmentSource));

    private static bool IsAtomicPaginationOwner(HtmlRenderBoxStyle style) =>
        style.OverflowY != "visible" || style.ClipPath != "none" || style.Clip != "auto";

    private OrdinaryPageSourceScope EnterOrdinaryPageSource(HtmlRenderBoxStyle style) => new(this,
        style.Display is "flex" or "inline-flex" or "grid" or "inline-grid" or "table" or "inline-table"
        || HasMultiColumnLayout(style) || style.FloatSide != "none" || ShouldExtractOutOfFlow(style));

    /// <summary>Leaves table cuts with their canonical owner while retaining visible cell-content extents.</summary>
    private OrdinaryPageSourceScope EnterTablePageSource() => new(this, specialized: true, tableOwner: true);

    /// <summary>Rebuilds independent ordinary sources against the physical page bands before flattening paint.</summary>
    private bool TryRelayoutConstrainedOverflow(HtmlRenderFlowBlock block, double available,
        double pageHeight, int pageNumber, HtmlCssPageGeometry geometry, out HtmlRenderFlowBlock reflowed) {
        reflowed = block;
        IElement? owner = block.OwnerElement;
        CheckCancellation();
        ChargeLayoutOperations(block.ForcedBreaks.Count, "constrained source forced-entry eligibility");
        if (owner == null || available <= 0.0001D || pageHeight <= 0D
            || (block.Height <= available + 0.0001D && !block.ForcedBreaks.Any(item => IsForcedPageBreak(item.Target)))
            || _projectedConstrainedOwners.Contains(owner)) return false;
        CheckCancellation();
        ChargeLayoutOperations(_constrainedOverflowElements.Count, "constrained source owner selection");
        IElement[] targets = _constrainedOverflowElements.Where(element => ContainsElementOrSelf(owner, element)).ToArray();
        if (targets.Length == 0) return false;
        IElement root = _document.Body ?? _document.DocumentElement ?? owner;
        bool isRoot = ReferenceEquals(owner, root);
        if (!isRoot && !ReferenceEquals(owner.ParentElement, root)) return false;
        _projectedConstrainedOwners.Add(owner);
        string? boundary = ContainsParallelRangeTable(block.Visuals)
            ? "table-formatting-source-owner" : ResolveConstrainedProjectionBoundary(owner);
        // Specialized owners retain their established planner and state. Do not
        // enter a speculative rebuild and then restore only part of its caches.
        ChargeLayoutOperations(block.ForcedBreaks.Count, "constrained source forced-boundary qualification");
        // The outer paginator has already applied side controls at the source
        // entry. Interior side events still require their existing fallback.
        boundary ??= block.ForcedBreaks.Any(item => item.ChangesPageName
            || item.Offset > 0.0001D && (item.Target is HtmlPageBreakTarget.Left or HtmlPageBreakTarget.Right))
                ? "internal-forced-source-break" : null;
        if (boundary != null) {
            ChargeLayoutOperations(targets.Length, "constrained overflow boundary qualification");
            if (targets.Any(_constrainedPaintOverflowElements.Contains))
                ReportConstrainedProjectionBoundary(block, boundary);
            return false;
        }
        var targetAncestors = new HashSet<IElement>();
        foreach (IElement target in targets) {
            for (IElement? parent = target.ParentElement; parent != null; parent = parent.ParentElement) {
                CheckCancellation();
                ChargeLayoutOperation("constrained source ancestor ownership");
                targetAncestors.Add(parent);
                if (ReferenceEquals(parent, owner)) break;
            }
        }
        var context = new ConstrainedSourceProjection(targetAncestors, geometry, pageNumber, pageHeight,
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
                if (context.NeedsProjection) ReportConstrainedProjectionBoundary(block, context.FailedReason);
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

    private void ReportConstrainedProjectionBoundary(HtmlRenderFlowBlock block, string reason) =>
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending,
            "Constrained overflow retained shared pagination outside the qualified independent ordinary-flow projection.",
            HtmlDiagnosticSeverity.Warning, block.Source, reason);

    private string? ResolveConstrainedProjectionBoundary(IElement owner) {
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
        foreach (IElement element in _registeredAbsoluteElements.Concat(_registeredFixedElements)) {
            CheckCancellation();
            ChargeLayoutOperation("constrained projection positioned ownership");
            if (ContainsElementOrSelf(owner, element)) return "positioned-source-owner";
        }
        return null;
    }

    private ConstrainedProjectionScope EnterConstrainedProjection(HtmlRenderBoxStyle style) =>
        new(_constrainedSourceProjection, IsAtomicPaginationOwner(style));

    /// <summary>Uses the canonical fragment planner; inserted gaps occupy destination space, not extra CSS height.</summary>
    private HtmlRenderFlowBlock ProjectOrdinaryContent(HtmlRenderFlowBlock source, PagedFloatBoundary boundary,
        IReadOnlyList<FlowPaintLayer>? childSources, out double entryGap, out ConstrainedInteriorGap? interiorGap,
        out double forcedFollowingMinimum) {
        entryGap = 0D;
        interiorGap = null;
        forcedFollowingMinimum = 0D;
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
        var consumedBoundaries = new HashSet<object>();
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
            ChargeLayoutOperations(projected.ForcedBreaks.Count, "constrained authored-boundary cursor");
            bool hasForcedBoundary = TryGetNextForcedBreak(projected.ForcedBreaks, cursor, out HtmlRenderForcedBreak? forcedBreak, consumedBoundaries)
                && (forcedBreak!.Offset < projected.PagedPaintExtent - 0.0001D
                    || forcedBreak!.IsAfter && !forcedBreak.Projected);
            if (remaining <= available + 0.0001D && !hasForcedBoundary) break;
            HtmlRenderFlowBlock paged = projected.ForPagination();
            bool movesOwnKeep = cursor <= 0.0001D && source.AvoidBreakInside
                && !hasForcedBoundary
                && source.PagedPaintExtent <= context.PageHeight + 0.0001D
                && remaining > available + 0.0001D;
            PagedBlockFragment? plan = movesOwnKeep ? null
                : PlanPagedFragment(paged, cursor, available, context.PageHeight,
                    canForceFragment: available >= context.PageHeight - 0.0001D, consumedBoundaries: consumedBoundaries);
            if (plan?.Forced == true || plan?.Header != null || plan?.Footer != null) {
                context.FailedReason = "forced-or-repeated-source-fragment";
                return source;
            }
            double cut = plan?.End ?? cursor;
            object? projectedForcedBoundary = plan?.AuthoredBreak?.Boundary;
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
                    if (childSources != null) {
                        foreach (FlowPaintLayer child in childSources) {
                            CheckCancellation();
                            ChargeLayoutOperation("constrained interior child-entry ownership");
                            if (child.Block.OwnerElement is { } entry
                                && child.Block.PagedPaintExtent > 0.0001D
                                && Math.Abs(child.Y - cut) <= 0.0001D) {
                                interiorGap = new ConstrainedInteriorGap(entry, gap);
                                return source;
                            }
                        }
                    }
                    double advance = cut < projected.Height - 0.0001D ? gap : 0D;
                    projected = InsertPagedSourceBreakGap(projected, cut, gap, false, advance,
                        plan == null ? null : ResolvePagedBackgroundWindow(cursor, plan), projectedForcedBoundary);
                    cut += gap;
                }
                context.Changed = true;
            } else if (projectedForcedBoundary != null) {
                projected = projected.WithVisuals(projected.Visuals, forcedBreaks: projected.ForcedBreaks.Select(item =>
                    ReferenceEquals(item.Boundary, projectedForcedBoundary)
                        ? item.ProjectAt(item.Offset) : item));
                context.Changed = true;
            }
            if (projectedForcedBoundary != null) consumedBoundaries.Add(projectedForcedBoundary);
            cursor = cut;
            available = context.PageHeight;
            pageNumber++;
        }
        foreach (HtmlRenderForcedBreak item in projected.ForcedBreaks) {
            CheckCancellation();
            ChargeLayoutOperation("constrained authored following-flow boundary");
            if (item.IsAfter && item.Projected && item.AdvancesFollowingFlow)
                forcedFollowingMinimum = Math.Max(forcedFollowingMinimum, item.Offset);
        }
        return projected;
    }

    private sealed class ConstrainedSourceProjection {
        internal ConstrainedSourceProjection(HashSet<IElement> targetAncestors, HtmlCssPageGeometry geometry,
            int pageNumber, double pageHeight, string? pageName, string? unsupportedBoundary) {
            TargetAncestors = targetAncestors;
            Geometry = geometry;
            PageNumber = pageNumber;
            PageHeight = pageHeight;
            PageName = pageName;
            UnsupportedBoundary = unsupportedBoundary;
        }
        internal HashSet<IElement> TargetAncestors { get; }
        internal HashSet<IElement> EntryRelayoutOwners { get; } = new();
        internal Stack<ConstrainedInteriorRelayoutFrame> InteriorRelayoutFrames { get; } = new();
        internal HtmlCssPageGeometry Geometry { get; }
        internal int PageNumber { get; }
        internal double PageHeight { get; }
        internal string? PageName { get; }
        internal string? UnsupportedBoundary { get; }
        internal int AtomicDepth { get; set; }
        internal bool Changed { get; set; }
        internal bool NeedsProjection { get; set; }
        internal string? FailedReason { get; set; }
    }

    private sealed class ConstrainedProjectionScope : IDisposable {
        private readonly ConstrainedSourceProjection? _context;
        private readonly bool _atomic;
        internal ConstrainedProjectionScope(ConstrainedSourceProjection? context, bool atomic) {
            _context = context;
            _atomic = context != null && atomic;
            // Each ordinary child establishes its own physical entry before
            // its ancestor imports the child's source. This keeps an overflow
            // stream independent of a later sibling's page-entry spacing.
            CanProject = context != null && context.AtomicDepth == 0 && !atomic;
            CanProjectAtomic = context != null && context.AtomicDepth == 0 && atomic;
            if (_atomic) context!.AtomicDepth++;
        }
        internal bool CanProject { get; }
        internal bool CanProjectAtomic { get; }
        public void Dispose() {
            if (_atomic) _context!.AtomicDepth--;
        }
    }

    private sealed class OrdinaryPageSourceScope : IDisposable {
        private readonly HtmlRenderLayoutEngine _engine;
        private readonly bool _specialized;
        private readonly bool _tableOwner;
        private readonly int _previousTableDepth;
        internal OrdinaryPageSourceScope(HtmlRenderLayoutEngine engine, bool specialized, bool tableOwner = false) {
            _engine = engine;
            _specialized = specialized;
            _tableOwner = tableOwner;
            _previousTableDepth = engine._tablePageSourceDepth;
            if (specialized) engine._specializedPageSourceDepth++;
            if (tableOwner) engine._tablePageSourceDepth = engine._specializedPageSourceDepth;
            IsOrdinary = engine._specializedPageSourceDepth == 0;
            // A nested flex/grid/float/column owner must not inherit permission
            // from its enclosing table. A nested table establishes its own cuts.
            RetainsConstrainedPaint = IsOrdinary
                || engine._tablePageSourceDepth == engine._specializedPageSourceDepth;
        }
        internal bool IsOrdinary { get; }
        internal bool RetainsConstrainedPaint { get; }
        public void Dispose() {
            if (_tableOwner) _engine._tablePageSourceDepth = _previousTableDepth;
            if (_specialized) _engine._specializedPageSourceDepth--;
        }
    }
}
