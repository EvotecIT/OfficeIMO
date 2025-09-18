namespace OfficeIMO.Pdf;

internal sealed class TableBlock : IPdfBlock {
    public System.Collections.Generic.List<string[]> Rows { get; } = new();
    public PdfAlign Align { get; }
    public TableBlock(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align) { Align = align; Rows.AddRange(rows); }
}

