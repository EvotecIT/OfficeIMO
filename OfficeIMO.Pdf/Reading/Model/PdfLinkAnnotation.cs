namespace OfficeIMO.Pdf;

/// <summary>
/// Simple link annotation read from a PDF page.
/// </summary>
public sealed class PdfLinkAnnotation {
    internal PdfLinkAnnotation(string uri, string? contents, double x1, double y1, double x2, double y2, int? pageNumber = null) {
        Guard.AbsoluteUri(uri, nameof(uri));
        Uri = uri;
        DestinationName = null;
        Contents = contents;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        PageNumber = pageNumber;
    }

    internal PdfLinkAnnotation(string? uri, string? destinationName, string? contents, double x1, double y1, double x2, double y2, int? pageNumber = null) {
        if (uri != null && destinationName != null) {
            throw new ArgumentException("A PDF link annotation can target either a URI or a named destination, not both.", nameof(destinationName));
        }

        if (uri == null && destinationName == null) {
            throw new ArgumentException("A PDF link annotation requires a URI or named destination target.", nameof(uri));
        }

        if (uri != null) {
            Guard.AbsoluteUri(uri, nameof(uri));
        }

        if (destinationName != null) {
            Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));
        }

        Uri = uri;
        DestinationName = destinationName;
        Contents = contents;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        PageNumber = pageNumber;
    }

    /// <summary>One-based page number when known; null when read directly from a page without document context.</summary>
    public int? PageNumber { get; }

    /// <summary>Absolute URI opened by the link annotation, or null for an internal named-destination link.</summary>
    public string? Uri { get; }

    /// <summary>Named destination opened by the link annotation, or null for a URI link.</summary>
    public string? DestinationName { get; }

    /// <summary>True when the link annotation opens an absolute URI.</summary>
    public bool IsUriLink => Uri is not null;

    /// <summary>True when the link annotation opens an internal named destination.</summary>
    public bool IsNamedDestinationLink => DestinationName is not null;

    /// <summary>Optional annotation contents metadata.</summary>
    public string? Contents { get; }

    /// <summary>Left edge of the annotation rectangle in PDF points.</summary>
    public double X1 { get; }

    /// <summary>Bottom edge of the annotation rectangle in PDF points.</summary>
    public double Y1 { get; }

    /// <summary>Right edge of the annotation rectangle in PDF points.</summary>
    public double X2 { get; }

    /// <summary>Top edge of the annotation rectangle in PDF points.</summary>
    public double Y2 { get; }

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => X2 - X1;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => Y2 - Y1;

    internal PdfLinkAnnotation WithPageNumber(int pageNumber) =>
        new PdfLinkAnnotation(Uri, DestinationName, Contents, X1, Y1, X2, Y2, pageNumber);
}
