namespace OfficeIMO.Pdf;

public sealed class PdfFooterCompose {
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

public sealed class FooterTextBuilder {
    private readonly System.Collections.Generic.List<FooterSegment> _segments;
    internal FooterTextBuilder(System.Collections.Generic.List<FooterSegment> segs) { _segments = segs; }
    public FooterTextBuilder Text(string s) { _segments.Add(new FooterSegment(FooterSegmentKind.Text, s)); return this; }
    public FooterTextBuilder CurrentPage() { _segments.Add(new FooterSegment(FooterSegmentKind.PageNumber)); return this; }
    public FooterTextBuilder TotalPages() { _segments.Add(new FooterSegment(FooterSegmentKind.TotalPages)); return this; }
}

