namespace OfficeIMO.Pdf;

/// <summary>
/// A simple document open action discovered from the document catalog.
/// </summary>
public sealed class PdfDocumentOpenAction {
    internal PdfDocumentOpenAction(
        string actionType,
        int? pageNumber,
        double? destinationTop,
        PdfOpenActionDestinationMode? destinationMode = null,
        double? destinationLeft = null,
        double? destinationBottom = null,
        double? destinationRight = null,
        double? destinationZoom = null) {
        ActionType = actionType;
        PageNumber = pageNumber;
        DestinationTop = destinationTop;
        DestinationMode = destinationMode;
        DestinationLeft = destinationLeft;
        DestinationBottom = destinationBottom;
        DestinationRight = destinationRight;
        DestinationZoom = destinationZoom;
    }

    /// <summary>Open action type, for example Destination or GoTo.</summary>
    public string ActionType { get; }

    /// <summary>One-based target page number when the destination could be resolved.</summary>
    public int? PageNumber { get; }

    /// <summary>Top coordinate of the destination when present.</summary>
    public double? DestinationTop { get; }

    /// <summary>Left coordinate of the destination when present.</summary>
    public double? DestinationLeft { get; }

    /// <summary>Bottom coordinate of the destination rectangle when present.</summary>
    public double? DestinationBottom { get; }

    /// <summary>Right coordinate of the destination rectangle when present.</summary>
    public double? DestinationRight { get; }

    /// <summary>Viewer destination mode when the open action uses a supported destination array.</summary>
    public PdfOpenActionDestinationMode? DestinationMode { get; }

    /// <summary>Zoom factor of an XYZ destination when present.</summary>
    public double? DestinationZoom { get; }
}
