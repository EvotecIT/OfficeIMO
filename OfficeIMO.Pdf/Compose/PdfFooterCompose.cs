namespace OfficeIMO.Pdf;

/// <summary>Footer builder (alignment, text, page number tokens).</summary>
public class PdfFooterCompose {
    private readonly PdfOptions _opts;
    internal PdfFooterCompose(PdfOptions opts) { _opts = opts; }
    public PdfFooterCompose AlignLeft() { _opts.FooterAlign = PdfAlign.Left; return this; }
    public PdfFooterCompose AlignCenter() { _opts.FooterAlign = PdfAlign.Center; return this; }
    public PdfFooterCompose AlignRight() { _opts.FooterAlign = PdfAlign.Right; return this; }
    public PdfFooterCompose PageNumber() { _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}"; return this; }
    public PdfFooterCompose PageNumberWithTotal() { _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}/{pages}"; return this; }
    public PdfFooterCompose Text(System.Action<FooterTextBuilder> build) {
        _opts.FooterSegments = new System.Collections.Generic.List<FooterSegment>();
        var b = new FooterTextBuilder(_opts.FooterSegments);
        build(b);
        _opts.ShowPageNumbers = true; // might be needed when builder inserts tokens
        return this;
    }
}

