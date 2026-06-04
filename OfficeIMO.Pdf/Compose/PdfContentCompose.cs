namespace OfficeIMO.Pdf;

/// <summary>Top-level container for page content (columns, rows, items).</summary>
public class PdfContentCompose {
    private readonly PdfDocument _doc;
    internal PdfContentCompose(PdfDocument doc) { _doc = doc; }
    /// <summary>Sets extra bottom padding (reserved for future).</summary>
    public PdfContentCompose PaddingBottom(double points) { /* reserved for future */ return this; }
    /// <summary>Adds one or more flow items directly to the page content.</summary>
    public PdfContentCompose Item(System.Action<PdfItemCompose> build) {
        Guard.NotNull(build, nameof(build));
        var item = new PdfItemCompose(_doc);
        build(item);
        return this;
    }
    /// <summary>Adds invisible vertical space directly to the page content flow.</summary>
    public PdfContentCompose Spacer(double height) {
        _doc.Spacer(height);
        return this;
    }
    /// <summary>Starts a new page directly from the page content flow.</summary>
    public PdfContentCompose PageBreak() {
        _doc.PageBreak();
        return this;
    }
    /// <summary>Adds a single content column (stack of items).</summary>
    public PdfContentCompose Column(System.Action<PdfColumnCompose> build) { Guard.NotNull(build, nameof(build)); var col = new PdfColumnCompose(_doc); build(col); return this; }
    /// <summary>Adds a row with percentage-based columns.</summary>
    public PdfContentCompose Row(System.Action<PdfRowCompose> build) { Guard.NotNull(build, nameof(build)); var row = new PdfRowCompose(_doc); build(row); row.Commit(); return this; }
}
