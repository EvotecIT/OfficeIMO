namespace OfficeIMO.Pdf;

/// <summary>
/// Simple link annotation read from a PDF page.
/// </summary>
public sealed class PdfLinkAnnotation {
    internal PdfLinkAnnotation(string uri, string? contents, double x1, double y1, double x2, double y2, int? pageNumber = null) {
        Guard.UriAction(uri, nameof(uri));
        Uri = uri;
        DestinationName = null;
        NamedAction = null;
        RemoteFile = null;
        RemoteDestinationName = null;
        Contents = contents;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        PageNumber = pageNumber;
    }

    internal PdfLinkAnnotation(
        string? uri,
        string? destinationName,
        string? contents,
        double x1,
        double y1,
        double x2,
        double y2,
        int? pageNumber = null,
        int? destinationPageNumber = null,
        int? destinationPageObjectNumber = null,
        double? destinationTop = null,
        PdfOpenActionDestinationMode? destinationMode = null,
        double? destinationLeft = null,
        double? destinationBottom = null,
        double? destinationRight = null,
        string? namedAction = null,
        string? remoteFile = null,
        string? remoteDestinationName = null,
        int? remoteDestinationPageNumber = null,
        double? remoteDestinationTop = null,
        PdfOpenActionDestinationMode? remoteDestinationMode = null,
        double? remoteDestinationLeft = null,
        double? remoteDestinationBottom = null,
        double? remoteDestinationRight = null) {
        bool hasDirectDestination = destinationPageNumber.HasValue ||
            destinationPageObjectNumber.HasValue ||
            destinationTop.HasValue ||
            destinationMode.HasValue ||
            destinationLeft.HasValue ||
            destinationBottom.HasValue ||
            destinationRight.HasValue;
        bool hasRemoteDestination = remoteFile != null ||
            remoteDestinationName != null ||
            remoteDestinationPageNumber.HasValue ||
            remoteDestinationTop.HasValue ||
            remoteDestinationMode.HasValue ||
            remoteDestinationLeft.HasValue ||
            remoteDestinationBottom.HasValue ||
            remoteDestinationRight.HasValue;
        int targetCount = 0;
        if (uri != null) {
            targetCount++;
        }

        if (destinationName != null) {
            targetCount++;
        }

        if (hasDirectDestination) {
            targetCount++;
        }

        if (namedAction != null) {
            targetCount++;
        }

        if (hasRemoteDestination) {
            targetCount++;
        }

        if (targetCount != 1) {
            throw new ArgumentException("A PDF link annotation requires exactly one URI, named destination, direct destination, named action, or remote destination target.", nameof(uri));
        }

        if (uri != null) {
            Guard.UriAction(uri, nameof(uri));
        }

        if (destinationName != null) {
            Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));
        }

        if (namedAction != null) {
            Guard.NotNullOrWhiteSpace(namedAction, nameof(namedAction));
        }

        if (hasRemoteDestination) {
            Guard.NotNullOrWhiteSpace(remoteFile, nameof(remoteFile));
        }

        if (remoteDestinationName != null) {
            Guard.NotNullOrWhiteSpace(remoteDestinationName, nameof(remoteDestinationName));
        }

        Uri = uri;
        DestinationName = destinationName;
        NamedAction = namedAction;
        RemoteFile = remoteFile;
        RemoteDestinationName = remoteDestinationName;
        RemoteDestinationPageNumber = remoteDestinationPageNumber;
        RemoteDestinationTop = remoteDestinationTop;
        RemoteDestinationMode = remoteDestinationMode;
        RemoteDestinationLeft = remoteDestinationLeft;
        RemoteDestinationBottom = remoteDestinationBottom;
        RemoteDestinationRight = remoteDestinationRight;
        DestinationPageNumber = destinationPageNumber;
        DestinationPageObjectNumber = destinationPageObjectNumber;
        DestinationTop = destinationTop;
        DestinationMode = destinationMode;
        DestinationLeft = destinationLeft;
        DestinationBottom = destinationBottom;
        DestinationRight = destinationRight;
        Contents = contents;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        PageNumber = pageNumber;
    }

    /// <summary>One-based page number when known; null when read directly from a page without document context.</summary>
    public int? PageNumber { get; }

    /// <summary>URI action target opened by the link annotation, or null for an internal named-destination link.</summary>
    public string? Uri { get; }

    /// <summary>Named destination opened by the link annotation, or null for a URI link.</summary>
    public string? DestinationName { get; }

    /// <summary>Named viewer action opened by the link annotation, for example NextPage, or null for other link target kinds.</summary>
    public string? NamedAction { get; }

    /// <summary>External PDF file targeted by a remote GoTo action, or null for other link target kinds.</summary>
    public string? RemoteFile { get; }

    /// <summary>Named destination in the external PDF targeted by a remote GoTo action, when present.</summary>
    public string? RemoteDestinationName { get; }

    /// <summary>One-based destination page number in the external PDF targeted by a simple remote GoTo destination array, when present.</summary>
    public int? RemoteDestinationPageNumber { get; }

    /// <summary>Top coordinate of the remote destination when present.</summary>
    public double? RemoteDestinationTop { get; }

    /// <summary>Left coordinate of the remote destination when present.</summary>
    public double? RemoteDestinationLeft { get; }

    /// <summary>Bottom coordinate of the remote destination rectangle when present.</summary>
    public double? RemoteDestinationBottom { get; }

    /// <summary>Right coordinate of the remote destination rectangle when present.</summary>
    public double? RemoteDestinationRight { get; }

    /// <summary>Viewer destination mode for the remote destination when present.</summary>
    public PdfOpenActionDestinationMode? RemoteDestinationMode { get; }

    /// <summary>One-based destination page number when the link targets a simple direct destination array.</summary>
    public int? DestinationPageNumber { get; }

    internal int? DestinationPageObjectNumber { get; }

    /// <summary>Top coordinate of the destination when present.</summary>
    public double? DestinationTop { get; }

    /// <summary>Left coordinate of the destination when present.</summary>
    public double? DestinationLeft { get; }

    /// <summary>Bottom coordinate of the destination rectangle when present.</summary>
    public double? DestinationBottom { get; }

    /// <summary>Right coordinate of the destination rectangle when present.</summary>
    public double? DestinationRight { get; }

    /// <summary>Viewer destination mode when the link uses a supported destination array.</summary>
    public PdfOpenActionDestinationMode? DestinationMode { get; }

    /// <summary>True when the link annotation opens a URI action target.</summary>
    public bool IsUriLink => Uri is not null;

    /// <summary>True when the link annotation opens an internal named destination.</summary>
    public bool IsNamedDestinationLink => DestinationName is not null;

    /// <summary>True when the link annotation opens a named viewer action such as NextPage.</summary>
    public bool IsNamedActionLink => NamedAction is not null;

    /// <summary>True when the link annotation opens a destination in another PDF file.</summary>
    public bool IsRemoteGoToLink => RemoteFile is not null;

    /// <summary>True when the link annotation opens an internal destination, either named or direct.</summary>
    public bool IsInternalDestinationLink => DestinationName is not null || DestinationPageNumber.HasValue || DestinationPageObjectNumber.HasValue || DestinationMode.HasValue || DestinationLeft.HasValue || DestinationTop.HasValue || DestinationBottom.HasValue || DestinationRight.HasValue;

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
        new PdfLinkAnnotation(Uri, DestinationName, Contents, X1, Y1, X2, Y2, pageNumber, DestinationPageNumber, DestinationPageObjectNumber, DestinationTop, DestinationMode, DestinationLeft, DestinationBottom, DestinationRight, NamedAction, RemoteFile, RemoteDestinationName, RemoteDestinationPageNumber, RemoteDestinationTop, RemoteDestinationMode, RemoteDestinationLeft, RemoteDestinationBottom, RemoteDestinationRight);

    internal PdfLinkAnnotation WithDestinationPageNumber(int? destinationPageNumber) =>
        new PdfLinkAnnotation(Uri, DestinationName, Contents, X1, Y1, X2, Y2, PageNumber, destinationPageNumber, DestinationPageObjectNumber, DestinationTop, DestinationMode, DestinationLeft, DestinationBottom, DestinationRight, NamedAction, RemoteFile, RemoteDestinationName, RemoteDestinationPageNumber, RemoteDestinationTop, RemoteDestinationMode, RemoteDestinationLeft, RemoteDestinationBottom, RemoteDestinationRight);
}
