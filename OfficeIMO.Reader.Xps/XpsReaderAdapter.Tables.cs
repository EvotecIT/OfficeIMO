namespace OfficeIMO.Reader.Xps;

internal static partial class XpsReaderAdapter {
    // Page Markdown needs one representation of each cell. Keep all literal blocks
    // at document scope, and retain uncovered rows when a rectangular grid is truncated.
    private static (HashSet<string> Blocks, HashSet<int> Tables) PageTableCoverage(OfficeDocumentModel model,
        IReadOnlyDictionary<OfficeDocumentModelTable, ReaderTable> tables, CancellationToken token) {
        var available = model.Pages.SelectMany(page => page.Tables).ToDictionary(table => table.Location!.TableIndex!.Value, table => tables[table]);
        var blocks = new HashSet<string>(StringComparer.Ordinal); var nestedTables = new HashSet<int>();
        void Visit(OfficeDocumentModelNode node, bool covered) {
            token.ThrowIfCancellationRequested();
            if (covered && !string.IsNullOrEmpty(node.Text) && node.Location.BlockAnchor != null) blocks.Add(node.Location.BlockAnchor);
            if (node.Kind == "table" && node.Location.TableIndex is int index && available.TryGetValue(index, out var table)) {
                if (covered) nestedTables.Add(index);
                int row = 0;
                foreach (var group in node.Children)
                    foreach (var child in group.Children) Visit(child, covered || row++ < table.Rows.Count);
            } else foreach (var child in node.Children) Visit(child, covered);
        }
        foreach (var story in model.Structure) Visit(story, false);
        return (blocks, nestedTables);
    }
}
