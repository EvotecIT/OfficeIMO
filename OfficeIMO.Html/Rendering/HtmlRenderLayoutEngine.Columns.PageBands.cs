using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Qualifies ordinary outer column flows for physical-page planning.</summary>
    private bool CanPaginateColumnRegion(IElement element, HtmlRenderBoxStyle style) {
        if (_options.Mode != HtmlRenderMode.Paged || _activeColumnOwner != null
            || style.ColumnFill != "auto" || style.WritingMode != "horizontal-tb"
            || style.OverflowX != "visible" || style.OverflowY != "visible"
            || EstablishesTransformedContainingBlock(style)) return false;
        HtmlRenderBoxStyle childStyle = style;
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (!_layoutStyles.TryGetValue(parent, out HtmlRenderBoxStyle? ancestor)) continue;
            if (ancestor.WritingMode != "horizontal-tb" || ancestor.OverflowY != "visible"
                || EstablishesTransformedContainingBlock(ancestor)) return false;
            bool columnFlex = (ancestor.Display == "flex" || ancestor.Display == "inline-flex")
                && (ancestor.FlexDirection == "column" || ancestor.FlexDirection == "column-reverse");
            bool flexOrGrid = ancestor.Display == "flex" || ancestor.Display == "inline-flex"
                || ancestor.Display == "grid" || ancestor.Display == "inline-grid";
            string alignment = columnFlex ? ancestor.JustifyContent
                : ResolveFlexAlignment(childStyle.AlignSelf, ancestor.AlignItems);
            bool startsAtStableEdge = alignment == "normal" || alignment == "start" || alignment == "flex-start"
                || !columnFlex && alignment == "stretch";
            if (flexOrGrid && (childStyle.MarginTopAuto || childStyle.MarginBottomAuto
                || columnFlex && (ancestor.FlexDirection != "column" || ancestor.FlexWrap != "nowrap")
                || !startsAtStableEdge)) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
                    "Multi-column pagination retains its source layout inside a container whose vertical alignment can move when column height changes.",
                    HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element));
                return false;
            }
            childStyle = ancestor;
        }
        return true;
    }

    private ColumnPageBands? ResolveColumnPageBands(IElement element, double width, double? declaredHeight, int requestedCount, double sourceHeight) {
        ColumnPaginationContext? context;
        if (!_pagedColumnContexts.TryGetValue(element, out context)) {
            ColumnPaginationAttempt? attempt = _columnPaginationAttempt;
            if (attempt == null || !ReferenceEquals(attempt.Element, element)
                || attempt.Boundary.RemainingHeight <= 0.0001D || attempt.Boundary.FragmentStart > 0.0001D) return null;
            context = new ColumnPaginationContext(attempt.Boundary, attempt.PageNumber, attempt.Geometry,
                attempt.PageHeight, width, attempt.PageName, attempt.UnfragmentedHeight);
        }
        // Short fixed-height sets retain their established overflow contract.
        // Their inline overflow is independent of the physical page fragment.
        bool intrinsicAllocation = _layoutStyles.TryGetValue(element, out HtmlRenderBoxStyle? style)
            && style.IntrinsicColumnHeightAllocation;
        if (Math.Abs(context.Width - width) > 0.0001D
            || !intrinsicAllocation && declaredHeight.HasValue && declaredHeight.Value < context.PageHeight - 0.0001D) return null;
        if (_columnEdgeFloatEntries.Values.Any(entry => ReferenceEquals(entry.Owner, element)
            && entry.Body.Height >= context.PageHeight - 0.0001D)) return null;
        bool reservations = _footnoteEntries.Values.Any(entry => ReferenceEquals(entry.ColumnOwner, element))
            || _columnEdgeFloatEntries.Values.Any(entry => ReferenceEquals(entry.Owner, element));
        return new ColumnPageBands(this, context, requestedCount, sourceHeight,
            HtmlRenderStyleResolver.DescribeSource(element),
            !intrinsicAllocation && !reservations ? declaredHeight : null);
    }

    /// <summary>
    /// Separates logical column capacity from the mapped block advance. Two
    /// columns consume two source slices but share one physical page band.
    /// </summary>
    private sealed class ColumnPageBands {
        private readonly HtmlRenderLayoutEngine _engine;
        private readonly int _count;
        private readonly double _sourceHeight;
        private readonly string _source;
        private readonly double? _declaredExtent;
        private readonly HashSet<int> _validated = new();

        internal ColumnPageBands(HtmlRenderLayoutEngine engine, ColumnPaginationContext context, int count,
            double sourceHeight, string source, double? declaredExtent = null) {
            _engine = engine;
            Context = context;
            _count = count;
            _sourceHeight = sourceHeight;
            _source = source;
            _declaredExtent = declaredExtent;
            FirstHeight = Math.Min(context.PageHeight, context.Boundary.RemainingHeight);
        }

        internal ColumnPaginationContext Context { get; }
        internal double FirstHeight { get; private set; }
        internal double FreshHeight => Context.PageHeight;
        internal bool DeferredFirstBand { get; private set; }
        internal bool Supported { get; private set; } = true;

        internal void DeferFirstBand() {
            DeferredFirstBand = true;
            FirstHeight = FreshHeight;
            _validated.Clear();
        }

        internal double Height(int column) {
            int band = BandForColumn(column);
            ValidateBand(band);
            double physical = band == 0 ? FirstHeight : FreshHeight;
            return !Supported ? _sourceHeight : _declaredExtent.HasValue
                ? Math.Max(0.01D, Math.Min(physical, _declaredExtent.Value - Start(band))) : physical;
        }

        private double LastBand => _declaredExtent.HasValue
            ? Math.Max(0D, Math.Ceiling((_declaredExtent.Value - FirstHeight - 0.0001D) / FreshHeight))
            : double.PositiveInfinity;

        internal int BandForColumn(int column) => (int)Math.Min(column / _count, LastBand);
        internal int ColumnInBand(int column) => column - BandForColumn(column) * _count;

        internal double Start(int band) {
            band = (int)Math.Min(band, LastBand);
            return band == 0 ? 0D : FirstHeight + (band - 1D) * FreshHeight;
        }

        // A finite final fragment has no later full-height column to defer to.
        // Oversized atomic content must still make progress in that fragment.
        internal double FreshHeightForColumn(int column) {
            if (!_declaredExtent.HasValue) return FreshHeight;
            int band = BandForColumn(column);
            double later = band < LastBand
                ? Math.Min(FreshHeight, _declaredExtent.Value - Start(band + 1)) : 0D;
            return Math.Max(Height(column), later);
        }

        // Entering another column consumes the preceding column's available
        // height, including line-fitting slack. Following siblings cannot use
        // that slack as though the column set ended at its final painted line.
        internal double MinimumBoxExtent(int columns) => ColumnInBand(columns - 1) > 0
            ? Start(BandForColumn(columns - 1)) + Height(columns - 1) : 0D;

        internal void ValidateExtent(double height) {
            int bands = (int)Math.Min(_engine._options.MaxPageCount,
                Math.Max(1D, 1D + Math.Ceiling(Math.Max(0D, height - FirstHeight) / FreshHeight)));
            for (int band = 0; band < bands; band++) ValidateBand(band);
        }

        private void ValidateBand(int band) {
            if (!_validated.Add(band)) return;
            _engine.CheckCancellation();
            _engine.ChargeLayoutOperation("column physical page geometry");
            int pageNumber = Context.PageNumber + band + (DeferredFirstBand ? 1 : 0);
            if (pageNumber > _engine._options.MaxPageCount)
                throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
            HtmlCssPageGeometry geometry = _engine._pageRules.ResolveGeometry(pageNumber, Context.PageName, _engine._options);
            if (Supported && (!SamePageGeometry(Context.Geometry, geometry)
                || Math.Abs(FreshHeight - _engine.ResolvePageBodyContentHeight(pageNumber, geometry)) > 0.0001D)) {
                Supported = false;
                _engine._diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
                    "Multi-column pagination retains its source layout because subsequent page geometry or reserved body space differs.",
                    HtmlDiagnosticSeverity.Warning, _source);
            }
        }
    }
}
