namespace OfficeIMO.Rtf;

/// <content>Provides section-owned headers and footers with RTF inheritance.</content>
public sealed partial class RtfSection {
    private List<RtfHeaderFooter> _headerFooters = new List<RtfHeaderFooter>();

    /// <summary>Headers and footers explicitly declared by this section. Missing kinds inherit from preceding sections.</summary>
    public IReadOnlyList<RtfHeaderFooter> HeaderFooters => _headerFooters.AsReadOnly();

    /// <summary>Adds an explicit section header or footer. An empty destination clears inherited content for that kind.</summary>
    public RtfHeaderFooter AddHeaderFooter(RtfHeaderFooterKind kind) {
        var headerFooter = new RtfHeaderFooter(kind);
        _headerFooters.Add(headerFooter);
        _document?.AddParsedHeaderFooter(headerFooter);
        return headerFooter;
    }

    /// <summary>Adds a header to this section.</summary>
    public RtfHeaderFooter AddHeader(RtfHeaderFooterKind kind = RtfHeaderFooterKind.Header) {
        if (kind != RtfHeaderFooterKind.Header && kind != RtfHeaderFooterKind.LeftHeader &&
            kind != RtfHeaderFooterKind.RightHeader && kind != RtfHeaderFooterKind.FirstHeader) {
            throw new ArgumentException("Header kind must be a header destination.", nameof(kind));
        }
        return AddHeaderFooter(kind);
    }

    /// <summary>Adds a footer to this section.</summary>
    public RtfHeaderFooter AddFooter(RtfHeaderFooterKind kind = RtfHeaderFooterKind.Footer) {
        if (kind != RtfHeaderFooterKind.Footer && kind != RtfHeaderFooterKind.LeftFooter &&
            kind != RtfHeaderFooterKind.RightFooter && kind != RtfHeaderFooterKind.FirstFooter) {
            throw new ArgumentException("Footer kind must be a footer destination.", nameof(kind));
        }
        return AddHeaderFooter(kind);
    }

    internal void AddParsedHeaderFooter(RtfHeaderFooter headerFooter) {
        if (!_headerFooters.Contains(headerFooter)) _headerFooters.Add(headerFooter);
    }
}
