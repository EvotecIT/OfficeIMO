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
        out IReadOnlyList<double> pageBreakOffsets,
        ColumnPageBands? pageBands = null) {
        pageBreakOffsets = Array.Empty<double>();
        if (_options.Mode != HtmlRenderMode.Paged || pageBands == null && plan.ColumnCount <= requestedCount) return plan;

        // Continuous multicol overflow advances inline. Paged overflow instead
        // resumes the requested column set in the next fragmentainer. Keep the
        // logical block extent so subsequent siblings follow the final set.
        // A legal atomic fragment can exceed the declared column height. Its
        // extent must fit the row before a mandatory page break is introduced.
        if (pageBands == null) columnHeight = Math.Max(columnHeight, plan.UsedHeight);
        var fragments = new List<MultiColumnFragment>(plan.Fragments.Count);
        double usedHeight = 0D;
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("paged column continuation");
            int row = pageBands?.BandForColumn(fragment.Column) ?? fragment.Column / requestedCount;
            double y = (pageBands?.Start(row) ?? row * columnHeight) + fragment.Y;
            fragments.Add(new MultiColumnFragment(fragment.Block, fragment.Start, fragment.End,
                pageBands?.ColumnInBand(fragment.Column) ?? fragment.Column % requestedCount, y));
            usedHeight = Math.Max(usedHeight, y + fragment.Height);
        }
        int rows = (pageBands?.BandForColumn(plan.ColumnCount - 1)
            ?? (plan.ColumnCount - 1) / requestedCount) + 1;
        pageBreakOffsets = Enumerable.Range(1, rows - 1)
            .Select(row => pageBands?.Start(row) ?? row * columnHeight)
            .Concat(pageBands is { DeferredFirstBand: true } ? new[] { 0D } : Array.Empty<double>())
            .OrderBy(offset => offset).ToArray();
        return new MultiColumnPlan(fragments,
            Math.Max(requestedCount, fragments.Select(fragment => fragment.Column + 1).DefaultIfEmpty(1).Max()), usedHeight);
    }
}
