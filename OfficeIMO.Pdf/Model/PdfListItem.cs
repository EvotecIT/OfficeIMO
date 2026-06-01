namespace OfficeIMO.Pdf;

/// <summary>
/// List item text with optional rich inline runs.
/// </summary>
public sealed class PdfListItem {
    /// <summary>Plain text value for readback, wrapping fallback, and simple item APIs.</summary>
    public string Text { get; }
    /// <summary>Rich inline runs rendered as the list item body.</summary>
    public System.Collections.Generic.IReadOnlyList<TextRun> Runs { get; }
    /// <summary>Optional named destination anchored to this list item.</summary>
    public string? BookmarkName { get; }
    /// <summary>Optional explicit marker rendered instead of the list block default marker.</summary>
    public string? Marker { get; }

    /// <summary>Create a plain list item.</summary>
    public PdfListItem(string text, string? bookmarkName = null, string? marker = null) {
        Guard.NotNull(text, nameof(text));
        if (bookmarkName != null) {
            Guard.NotNullOrWhiteSpace(bookmarkName, nameof(bookmarkName));
        }
        if (marker != null) {
            Guard.NotNullOrWhiteSpace(marker, nameof(marker));
        }

        Text = text;
        Runs = new[] { TextRun.Normal(text) };
        BookmarkName = bookmarkName;
        Marker = marker;
    }

    /// <summary>Create a rich list item from inline text runs.</summary>
    public PdfListItem(System.Collections.Generic.IEnumerable<TextRun> runs, string? bookmarkName = null, string? marker = null) {
        Guard.NotNull(runs, nameof(runs));
        if (bookmarkName != null) {
            Guard.NotNullOrWhiteSpace(bookmarkName, nameof(bookmarkName));
        }
        if (marker != null) {
            Guard.NotNullOrWhiteSpace(marker, nameof(marker));
        }

        var snapshot = new System.Collections.Generic.List<TextRun>();
        foreach (TextRun? run in runs) {
            if (run == null) {
                throw new System.ArgumentException("List item runs cannot contain null entries.", nameof(runs));
            }

            snapshot.Add(run);
        }
        if (snapshot.Count == 0) {
            throw new System.ArgumentException("List item must contain at least one text run.", nameof(runs));
        }

        Runs = snapshot.AsReadOnly();
        Text = string.Concat(snapshot.Select(run => run.Text));
        BookmarkName = bookmarkName;
        Marker = marker;
    }

    /// <summary>Create a normal unstyled item.</summary>
    public static PdfListItem Plain(string text, string? bookmarkName = null, string? marker = null) => new PdfListItem(text, bookmarkName, marker);

    /// <summary>Create a rich item from inline text runs.</summary>
    public static PdfListItem Rich(System.Collections.Generic.IEnumerable<TextRun> runs, string? bookmarkName = null, string? marker = null) => new PdfListItem(runs, bookmarkName, marker);
}
