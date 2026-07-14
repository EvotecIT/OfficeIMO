namespace OfficeIMO.Pdf;

/// <summary>Captures the final destination and page position of a generated section.</summary>
public sealed class PdfSectionReference {
    private readonly object _sync = new object();
    private string? _destinationName;
    private string? _title;
    private int? _pageNumber;
    private double? _y;

    /// <summary>Generated named destination.</summary>
    public string? DestinationName { get { lock (_sync) return _destinationName; } }
    /// <summary>Section title.</summary>
    public string? Title { get { lock (_sync) return _title; } }
    /// <summary>One-based physical output page, or null before successful layout.</summary>
    public int? PageNumber { get { lock (_sync) return _pageNumber; } }
    /// <summary>Top position in PDF points, or null before successful layout.</summary>
    public double? Y { get { lock (_sync) return _y; } }

    internal void Set(string destinationName, string title, int pageNumber, double y) {
        lock (_sync) {
            _destinationName = destinationName;
            _title = title;
            _pageNumber = pageNumber;
            _y = y;
        }
    }
}
