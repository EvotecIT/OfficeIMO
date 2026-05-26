namespace OfficeIMO.Pdf;

/// <summary>
/// A simple document open action discovered from the document catalog.
/// </summary>
public sealed class PdfDocumentOpenAction {
    internal PdfDocumentOpenAction(string actionType, int? pageNumber, double? destinationTop) {
        ActionType = actionType;
        PageNumber = pageNumber;
        DestinationTop = destinationTop;
    }

    /// <summary>Open action type, for example Destination or GoTo.</summary>
    public string ActionType { get; }

    /// <summary>One-based target page number when the destination could be resolved.</summary>
    public int? PageNumber { get; }

    /// <summary>Top coordinate of the destination when present.</summary>
    public double? DestinationTop { get; }
}
