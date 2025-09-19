namespace OfficeIMO.Pdf;

/// <summary>Row builder with percentage-based columns.</summary>
public class PdfRowCompose {
    private readonly PdfDoc _doc;
    private readonly RowBlock _row = new RowBlock();
    internal PdfRowCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Adds a column with the given width percentage.</summary>
    public PdfRowCompose Column(double widthPercent, System.Action<PdfRowColumnCompose> build) {
        var col = new RowColumn(widthPercent);
        var cc = new PdfRowColumnCompose(col);
        build(cc);
        _row.Columns.Add(col);
        return this;
    }
    internal void Commit() { _doc.AddRow(_row); }
}

