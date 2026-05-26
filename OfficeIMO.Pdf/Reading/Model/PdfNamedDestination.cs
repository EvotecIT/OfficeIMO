namespace OfficeIMO.Pdf;

/// <summary>
/// A named destination discovered from the document catalog.
/// </summary>
public sealed class PdfNamedDestination {
    internal PdfNamedDestination(string name, int? pageNumber, double? destinationTop) {
        Name = name;
        PageNumber = pageNumber;
        DestinationTop = destinationTop;
    }

    /// <summary>Destination name used by links, outlines, or viewer navigation.</summary>
    public string Name { get; }

    /// <summary>One-based target page number when the destination could be resolved.</summary>
    public int? PageNumber { get; }

    /// <summary>Top coordinate of the destination when present.</summary>
    public double? DestinationTop { get; }
}
