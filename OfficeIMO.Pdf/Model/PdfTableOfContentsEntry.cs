namespace OfficeIMO.Pdf;

/// <summary>A table-of-contents entry targeting an authored bookmark.</summary>
public sealed class PdfTableOfContentsEntry {
    /// <summary>Creates an entry whose page number is resolved from final PDF layout.</summary>
    /// <param name="title">Visible entry text.</param>
    /// <param name="destinationName">Name supplied to the bookmark API.</param>
    /// <param name="level">Hierarchy level from one through nine.</param>
    /// <param name="linkContents">Accessible link description, or the visible title when omitted.</param>
    public PdfTableOfContentsEntry(string title, string destinationName, int level = 1, string? linkContents = null) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));
        if (level < 1 || level > 9) throw new ArgumentOutOfRangeException(nameof(level));
        Title = title;
        DestinationName = destinationName;
        Level = level;
        LinkContents = linkContents ?? title;
    }

    /// <summary>Visible entry text.</summary>
    public string Title { get; }
    /// <summary>Authored bookmark destination.</summary>
    public string DestinationName { get; }
    /// <summary>Hierarchy level.</summary>
    public int Level { get; }
    /// <summary>Accessible link description.</summary>
    public string LinkContents { get; }
}
