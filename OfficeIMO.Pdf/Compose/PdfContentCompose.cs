namespace OfficeIMO.Pdf;

/// <summary>Top-level container for page content (columns, rows, items).</summary>
public class PdfContentCompose {
    private readonly PdfDoc _doc;
    internal PdfContentCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Sets extra bottom padding (reserved for future).</summary>
    public PdfContentCompose PaddingBottom(double points) { /* reserved for future */ return this; }
    /// <summary>Adds a single content column (stack of items).</summary>
    public PdfContentCompose Column(System.Action<PdfColumnCompose> build) { var col = new PdfColumnCompose(_doc); build(col); return this; }
    /// <summary>Adds a row with percentage-based columns.</summary>
    public PdfContentCompose Row(System.Action<PdfRowCompose> build) { var row = new PdfRowCompose(_doc); build(row); row.Commit(); return this; }
}

