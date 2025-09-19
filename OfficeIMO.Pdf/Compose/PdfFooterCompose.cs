namespace OfficeIMO.Pdf;

/// <summary>Footer builder (alignment, text, page number tokens).</summary>
public class PdfFooterCompose {
    private readonly PdfOptions _opts;
    internal PdfFooterCompose(PdfOptions opts) { _opts = opts; }
    /// <summary>Sets footer alignment to the left.</summary>
    public PdfFooterCompose AlignLeft() { _opts.FooterAlign = PdfAlign.Left; return this; }
    /// <summary>Sets footer alignment to the center.</summary>
    public PdfFooterCompose AlignCenter() { _opts.FooterAlign = PdfAlign.Center; return this; }
    /// <summary>Sets footer alignment to the right.</summary>
    public PdfFooterCompose AlignRight() { _opts.FooterAlign = PdfAlign.Right; return this; }
    /// <summary>Renders the current page number in the footer.</summary>
    public PdfFooterCompose PageNumber() { _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}"; return this; }
    /// <summary>Renders the current page number and total pages in the footer.</summary>
    public PdfFooterCompose PageNumberWithTotal() { _opts.ShowPageNumbers = true; _opts.FooterFormat = "{page}/{pages}"; return this; }
    /// <summary>Builds a custom footer from text and tokens.</summary>
    /// <param name="build">Delegate to compose footer segments.</param>
    public PdfFooterCompose Text(System.Action<FooterTextBuilder> build) {
        _opts.FooterSegments = new System.Collections.Generic.List<FooterSegment>();
        var b = new FooterTextBuilder(_opts.FooterSegments);
        build(b);
        _opts.ShowPageNumbers = true; // might be needed when builder inserts tokens
        return this;
    }
}
