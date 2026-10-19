namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private int _suppressedPageBoundaryDepth;

    /// <summary>Ends sibling margin collapse only where the enclosing flow retains page controls.</summary>
    private bool StartsPageMarginGroup(HtmlRenderFlowBlock? previous, HtmlRenderFlowBlock following) {
        if (_options.Mode != HtmlRenderMode.Paged || _suppressedPageBoundaryDepth > 0 || _columnContentDepth > 0)
            return false;
        if (IsForcedPageBreak(ResolveRootBreakBefore(following))) return true;
        if (previous == null) return false;
        return IsForcedPageBreak(ResolveRootBreakAfter(previous))
            || !string.Equals(previous.PageName, following.PageName, StringComparison.Ordinal);
    }

    /// <summary>Preserves source-flow margins beneath owners that omit descendant page boundaries.</summary>
    private readonly struct DescendantPageBoundaryScope : IDisposable {
        private readonly HtmlRenderLayoutEngine _engine;
        private readonly bool _suppressed;

        internal DescendantPageBoundaryScope(HtmlRenderLayoutEngine engine, bool retainsPageBoundaries) {
            _engine = engine;
            _suppressed = !retainsPageBoundaries;
            if (_suppressed) engine._suppressedPageBoundaryDepth++;
        }

        public void Dispose() {
            if (_suppressed) _engine._suppressedPageBoundaryDepth--;
        }
    }
}
