namespace OfficeIMO.Pdf;

/// <summary>
/// A named destination discovered from the document catalog.
/// </summary>
public sealed class PdfNamedDestination {
    internal PdfNamedDestination(
        string name,
        int? pageNumber,
        double? destinationTop,
        PdfOpenActionDestinationMode? destinationMode = null,
        double? destinationLeft = null,
        double? destinationBottom = null,
        double? destinationRight = null,
        double? destinationZoom = null) {
        Name = name;
        PageNumber = pageNumber;
        DestinationTop = destinationTop;
        DestinationMode = destinationMode;
        DestinationLeft = destinationLeft;
        DestinationBottom = destinationBottom;
        DestinationRight = destinationRight;
        DestinationZoom = destinationZoom;
    }

    /// <summary>Destination name used by links, outlines, or viewer navigation.</summary>
    public string Name { get; }

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

    /// <summary>Viewer destination mode when the destination uses a supported destination array.</summary>
    public PdfOpenActionDestinationMode? DestinationMode { get; }

    /// <summary>Zoom factor of an XYZ destination when present.</summary>
    public double? DestinationZoom { get; }
}
