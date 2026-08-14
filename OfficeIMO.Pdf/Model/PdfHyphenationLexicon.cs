namespace OfficeIMO.Pdf;

/// <summary>
/// Immutable, dependency-free dictionary of language-specific word hyphenation points.
/// </summary>
/// <remarks>
/// Entries use an explicit marker, for example <c>typog-ra-phy</c>. The marker is removed from
/// lookup text and its UTF-16 position becomes a preferred generated-PDF line break. Dictionary
/// data remains caller-owned, so applications can ship only the languages they use.
/// </remarks>
public sealed class PdfHyphenationLexicon {
    private readonly OfficeIMO.Drawing.OfficeTextHyphenationLexicon _inner;

    /// <summary>Creates a case-insensitive hyphenation dictionary from marked words.</summary>
    /// <param name="hyphenatedWords">Words containing one or more explicit break markers.</param>
    /// <param name="breakMarker">Marker removed from entries and converted into a break position.</param>
    /// <param name="minimumPrefixLength">Minimum UTF-16 word length retained before a break.</param>
    /// <param name="minimumSuffixLength">Minimum UTF-16 word length retained after a break.</param>
    public PdfHyphenationLexicon(
        IEnumerable<string> hyphenatedWords,
        char breakMarker = '-',
        int minimumPrefixLength = 2,
        int minimumSuffixLength = 2) {
        _inner = new OfficeIMO.Drawing.OfficeTextHyphenationLexicon(hyphenatedWords, breakMarker, minimumPrefixLength, minimumSuffixLength);
    }

    /// <summary>Marker used by the source dictionary.</summary>
    public char BreakMarker => _inner.BreakMarker;

    /// <summary>Minimum UTF-16 length retained before each returned break.</summary>
    public int MinimumPrefixLength => _inner.MinimumPrefixLength;

    /// <summary>Minimum UTF-16 length retained after each returned break.</summary>
    public int MinimumSuffixLength => _inner.MinimumSuffixLength;

    /// <summary>Number of normalized words in the dictionary.</summary>
    public int Count => _inner.Count;

    /// <summary>Returns preferred UTF-16 break positions for a word, or an empty list when it is not present.</summary>
    public IReadOnlyList<int> GetBreakpoints(string token) {
        return _inner.GetBreakpoints(token);
    }

    /// <summary>Returns true when the normalized word is present.</summary>
    public bool Contains(string token) {
        return _inner.Contains(token);
    }

    /// <summary>Creates the callback shape consumed by <see cref="PdfOptions.SetTextHyphenation"/>.</summary>
    public PdfTextHyphenationCallback AsCallback() => GetBreakpoints;

}
