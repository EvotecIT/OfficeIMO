namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Paints this table's group and row layers below all of its cells, including cells spanning
    /// later rows. Cell-bearing semantic rows and nested formatting/effect contexts stay intact.
    /// Repeated sections supply only their own layers and retain their existing source offset.
    /// </summary>
    private void InsertTableRowBackgroundLayers(
        List<HtmlRenderVisual> visuals,
        IEnumerable<TableRowBackgroundLayer> layers,
        int insertionIndex,
        double backgroundOffsetY) {
        TableRowBackgroundLayer[] retained = layers.ToArray();
        if (retained.Length == 0) return;
        var ordered = new List<HtmlRenderVisual>(visuals.Count);
        for (int index = 0; index < insertionIndex; index++) {
            AppendExisting(visuals[index]);
        }
        foreach (TableRowBackgroundLayer layer in retained.Where(layer => layer.IsGroup)
            .Concat(retained.Where(layer => !layer.IsGroup))) {
            CheckCancellation();
            ChargeLayoutOperation("table background layering");
            foreach (HtmlRenderVisual background in layer.Visuals) {
                ordered.Add(background.Translate(0D, backgroundOffsetY, ordered.Count));
            }
        }
        for (int index = insertionIndex; index < visuals.Count; index++) {
            AppendExisting(visuals[index]);
        }
        visuals.Clear();
        visuals.AddRange(ordered);

        void AppendExisting(HtmlRenderVisual visual) {
            CheckCancellation();
            // Only the table's direct row envelope needs a new paint index. Keep
            // cell descendants, nested tables and their contexts as the same tree.
            ordered.Add(visual.PaintOrder == ordered.Count ? visual
                : visual is HtmlRenderSemanticGroup group
                    ? visual.CopyStackingContextTo(group.ProjectPaint(group.Visuals, 0D, 0D, ordered.Count))
                    : visual.Translate(0D, 0D, ordered.Count));
        }
    }

    private sealed record TableRowBackgroundLayer(
        IReadOnlyList<HtmlRenderVisual> Visuals,
        bool IsGroup,
        bool IsHeader,
        bool IsFooter);
}
