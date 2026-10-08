namespace OfficeIMO.Xps;

internal sealed partial class XpsDocumentModelProjection {
    private bool AddTable(XpsStructureNode table, OfficeDocumentModelLocation location) {
        var nativeRows = table.Children.SelectMany(g => g.Children).ToArray();
        if (nativeRows.Where((row, index) => row.Children.Any(cell => cell.ColumnSpan > 1000 || cell.RowSpan > nativeRows.Length - index)).Any()) {
            Diagnostic("XpsTableGridUnavailable", "The rectangular table projection cannot represent these spans; the native recursive structure retains them.");
            return false;
        }
        var rows = new List<IReadOnlyList<string>>();
        var contentPages = new HashSet<int>();
        var occupied = new Dictionary<int, int>(); int width = 0;
        foreach (var row in nativeRows) {
            _budget.Charge(); var values = new Dictionary<int, string>(); int column = 0;
            var next = occupied.Where(pair => pair.Value > 1).ToDictionary(pair => pair.Key, pair => pair.Value - 1);
            foreach (var cell in row.Children) {
                _budget.Charge();
                while (Enumerable.Range(column, cell.ColumnSpan).Any(occupied.ContainsKey)) {
                    _budget.Charge(); column++;
                    if (column + cell.ColumnSpan > 1000) { Diagnostic("XpsTableGridUnavailable", "The rectangular table exceeds 1000 columns; native structure is retained."); return false; }
                }
                if (column + cell.ColumnSpan > 1000) { Diagnostic("XpsTableGridUnavailable", "The rectangular table exceeds 1000 columns; native structure is retained."); return false; }
                values.Add(column, CellText(cell, contentPages));
                for (int offset = 0; offset < cell.ColumnSpan; offset++) {
                    _budget.Charge();
                    if (cell.RowSpan > 1) next.Add(column + offset, cell.RowSpan - 1);
                }
                column += cell.ColumnSpan;
            }
            int rowWidth = Math.Max(column, occupied.Count == 0 ? 0 : occupied.Keys.Max() + 1);
            _budget.Charge(rowWidth); width = Math.Max(width, rowWidth);
            rows.Add(Enumerable.Range(0, rowWidth).Select(index => values.TryGetValue(index, out string? value) ? value : string.Empty).ToArray());
            occupied = next;
        }
        var normalized = new List<IReadOnlyList<string>>(rows.Count);
        foreach (var row in rows) {
            _budget.Charge(width - row.Count);
            normalized.Add(row.Concat(Enumerable.Repeat(string.Empty, width - row.Count)).ToArray());
        }
        var projected = new OfficeDocumentModelTable { Kind = "xps-native-table", Location = location,
            Columns = Enumerable.Range(1, width).Select(i => "Column " + i.ToString(CultureInfo.InvariantCulture)).ToArray(),
            Rows = normalized, TotalRowCount = normalized.Count };
        _tables.Add(projected); _tablePages.Add(projected, contentPages);
        if (contentPages.Count > 1) Diagnostic("XpsCrossPageTable", "A logical table spans physical pages; its grid remains at document scope and page projections retain their own text.");
        return true;
    }
    private string CellText(XpsStructureNode node, HashSet<int> pages) {
        _budget.Charge();
        if (node.Content != null) pages.Add(node.Content.PageIndex + 1);
        if (node.Marker != null) pages.Add(node.Marker.PageIndex + 1);
        string text = node.Content?.Text ?? string.Concat(node.Children.Select(child => CellText(child, pages)));
        return (node.Marker?.Text ?? string.Empty) + text;
    }
}
