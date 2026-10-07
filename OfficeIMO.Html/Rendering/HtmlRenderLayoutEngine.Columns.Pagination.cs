namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Maps overflow columns to later page fragments without widening the column
    /// set or changing continuous layout. Returns mandatory row-boundary breaks.
    /// </summary>
    private MultiColumnPlan ResolvePagedColumnOverflow(
        MultiColumnPlan plan,
        int requestedCount,
        double columnHeight,
        out IReadOnlyList<double> pageBreakOffsets) {
        pageBreakOffsets = Array.Empty<double>();
        if (_options.Mode != HtmlRenderMode.Paged || plan.ColumnCount <= requestedCount) return plan;

        // Continuous multicol overflow advances inline. Paged overflow instead
        // resumes the requested column set in the next fragmentainer. Keep the
        // logical block extent so subsequent siblings follow the final set.
        // A legal atomic fragment can exceed the declared column height. Its
        // extent must fit the row before a mandatory page break is introduced.
        columnHeight = Math.Max(columnHeight, plan.UsedHeight);
        var fragments = new List<MultiColumnFragment>(plan.Fragments.Count);
        double usedHeight = 0D;
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("paged column continuation");
            int row = fragment.Column / requestedCount;
            double y = row * columnHeight + fragment.Y;
            fragments.Add(new MultiColumnFragment(fragment.Block, fragment.Start, fragment.End,
                fragment.Column % requestedCount, y));
            usedHeight = Math.Max(usedHeight, y + fragment.Height);
        }
        int rows = (plan.ColumnCount + requestedCount - 1) / requestedCount;
        pageBreakOffsets = Enumerable.Range(1, rows - 1).Select(row => row * columnHeight).ToArray();
        return new MultiColumnPlan(fragments, requestedCount, usedHeight);
    }
}
