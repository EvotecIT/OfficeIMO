namespace OfficeIMO.Pdf;

internal sealed class TableBlock : IPdfBlock {
    public System.Collections.Generic.List<string[]> Rows { get; } = new();
    public PdfAlign Align { get; }
    public PdfTableStyle? Style { get; }
    // Optional per-cell link URIs: key = (rowIndex, colIndex)
    public System.Collections.Generic.Dictionary<(int Row, int Col), string> Links { get; } = new();
    public TableBlock(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align, PdfTableStyle? style) {
        Guard.NotNull(rows, nameof(rows));
        Align = align; Style = style;
        // Validate and normalize rows
        foreach (var r in rows) {
            if (r is null) throw new System.ArgumentException("Table rows cannot contain null entries.", nameof(rows));
            Rows.Add(r);
        }
        if (Rows.Count > 0) {
            // Ensure header row doesn't contain null cells to avoid rendering artifacts
            var header = Rows[0];
            for (int i = 0; i < header.Length; i++) if (header[i] is null) header[i] = string.Empty;
        }
    }
}
