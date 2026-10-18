using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Qualifies ordinary outer column flows for physical-page planning.</summary>
    private bool CanPaginateColumnRegion(IElement element, HtmlRenderBoxStyle style) {
        if (_options.Mode != HtmlRenderMode.Paged || _activeColumnOwner != null
            || style.ColumnFill != "auto" || style.WritingMode != "horizontal-tb"
            || style.OverflowX != "visible" || style.OverflowY != "visible"
            || EstablishesTransformedContainingBlock(style)) return false;
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (_layoutStyles.TryGetValue(parent, out HtmlRenderBoxStyle? ancestor)
                && (ancestor.WritingMode != "horizontal-tb" || ancestor.OverflowY != "visible"
                    || EstablishesTransformedContainingBlock(ancestor))) return false;
        }
        return true;
    }

    private ColumnPageBands? ResolveColumnPageBands(IElement element, double width, double? declaredHeight, int requestedCount) {
        ColumnPaginationContext? context;
        if (!_pagedColumnContexts.TryGetValue(element, out context)) {
            ColumnPaginationAttempt? attempt = _columnPaginationAttempt;
            if (attempt == null || !ReferenceEquals(attempt.Element, element)
                || attempt.Boundary.RemainingHeight <= 0.0001D || attempt.Boundary.FragmentStart > 0.0001D) return null;
            context = new ColumnPaginationContext(attempt.Boundary, attempt.PageNumber, attempt.Geometry,
                attempt.PageHeight, width, attempt.PageName);
        }
        // Short fixed-height sets retain their established overflow contract.
        // Their inline overflow is independent of the physical page fragment.
        if (Math.Abs(context.Width - width) > 0.0001D
            || declaredHeight.HasValue && declaredHeight.Value < context.PageHeight - 0.0001D) return null;
        return new ColumnPageBands(this, context, requestedCount);
    }

    /// <summary>
    /// Separates logical column capacity from the mapped block advance. Two
    /// columns consume two source slices but share one physical page band.
    /// </summary>
    private sealed class ColumnPageBands {
        private readonly HtmlRenderLayoutEngine _engine;
        private readonly int _count;
        private readonly HashSet<int> _validated = new();

        internal ColumnPageBands(HtmlRenderLayoutEngine engine, ColumnPaginationContext context, int count) {
            _engine = engine;
            Context = context;
            _count = count;
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
            int band = column / _count;
            ValidateBand(band);
            return band == 0 ? FirstHeight : FreshHeight;
        }

        internal double Start(int band) => band == 0 ? 0D : FirstHeight + (band - 1D) * FreshHeight;

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
            HtmlCssPageGeometry geometry = _engine._pageRules.ResolveGeometry(pageNumber, Context.PageName, _engine._options);
            if (!SamePageGeometry(Context.Geometry, geometry)
                || Math.Abs(FreshHeight - _engine.ResolvePageBodyContentHeight(pageNumber, geometry)) > 0.0001D)
                Supported = false;
        }
    }
}
