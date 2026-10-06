namespace OfficeIMO.Studio.Features.Reader;

/// <summary>One search occurrence and its visual page geometry.</summary>
public sealed record PdfSearchHit(int PageNumber, string Snippet) {
    /// <summary>The union of the occurrence's line bounds.</summary>
    public StudioRectangle Bounds { get; init; }
    /// <summary>Highlight rectangles for a match that can span multiple lines.</summary>
    public IReadOnlyList<StudioRectangle> LineBounds { get; init; } = [];
    /// <summary>One-based occurrence number in the document.</summary>
    public int OccurrenceNumber { get; init; }
}
