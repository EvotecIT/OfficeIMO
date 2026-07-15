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
    private readonly Dictionary<string, int[]> _entries;

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
        Guard.NotNull(hyphenatedWords, nameof(hyphenatedWords));
        if (char.IsLetterOrDigit(breakMarker) || char.IsSurrogate(breakMarker) || char.IsWhiteSpace(breakMarker)) {
            throw new ArgumentException("Hyphenation break markers cannot be letters, digits, surrogates, or whitespace.", nameof(breakMarker));
        }

        if (minimumPrefixLength < 1) {
            throw new ArgumentOutOfRangeException(nameof(minimumPrefixLength), "Hyphenation minimum prefix length must be positive.");
        }

        if (minimumSuffixLength < 1) {
            throw new ArgumentOutOfRangeException(nameof(minimumSuffixLength), "Hyphenation minimum suffix length must be positive.");
        }

        BreakMarker = breakMarker;
        MinimumPrefixLength = minimumPrefixLength;
        MinimumSuffixLength = minimumSuffixLength;
        _entries = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
        foreach (string entry in hyphenatedWords) {
            AddEntry(entry);
        }
    }

    /// <summary>Marker used by the source dictionary.</summary>
    public char BreakMarker { get; }

    /// <summary>Minimum UTF-16 length retained before each returned break.</summary>
    public int MinimumPrefixLength { get; }

    /// <summary>Minimum UTF-16 length retained after each returned break.</summary>
    public int MinimumSuffixLength { get; }

    /// <summary>Number of normalized words in the dictionary.</summary>
    public int Count => _entries.Count;

    /// <summary>Returns preferred UTF-16 break positions for a word, or an empty list when it is not present.</summary>
    public IReadOnlyList<int> GetBreakpoints(string token) {
        Guard.NotNull(token, nameof(token));
        if (!_entries.TryGetValue(token, out int[]? points)) {
            return Array.Empty<int>();
        }

        return points.ToArray();
    }

    /// <summary>Returns true when the normalized word is present.</summary>
    public bool Contains(string token) {
        Guard.NotNull(token, nameof(token));
        return _entries.ContainsKey(token);
    }

    /// <summary>Creates the callback shape consumed by <see cref="PdfOptions.SetTextHyphenation"/>.</summary>
    public PdfTextHyphenationCallback AsCallback() => GetBreakpoints;

    private void AddEntry(string entry) {
        if (string.IsNullOrWhiteSpace(entry)) {
            throw new ArgumentException("Hyphenation dictionary entries cannot be null, empty, or whitespace.", nameof(entry));
        }

        string value = entry.Trim();
        if (value[0] == BreakMarker || value[value.Length - 1] == BreakMarker) {
            throw new ArgumentException("Hyphenation dictionary entries cannot start or end with the break marker.", nameof(entry));
        }

        var word = new StringBuilder(value.Length);
        var points = new List<int>();
        bool previousWasMarker = false;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == BreakMarker) {
                if (previousWasMarker) {
                    throw new ArgumentException("Hyphenation dictionary entries cannot contain adjacent break markers.", nameof(entry));
                }

                points.Add(word.Length);
                previousWasMarker = true;
                continue;
            }

            word.Append(current);
            previousWasMarker = false;
        }

        if (points.Count == 0) {
            throw new ArgumentException("Hyphenation dictionary entries must contain at least one break marker.", nameof(entry));
        }

        string normalizedWord = word.ToString();
        int[] validPoints = points
            .Where(point => point >= MinimumPrefixLength && normalizedWord.Length - point >= MinimumSuffixLength)
            .Where(point => !(char.IsHighSurrogate(normalizedWord[point - 1]) && char.IsLowSurrogate(normalizedWord[point])))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
        if (validPoints.Length == 0) {
            throw new ArgumentException("Hyphenation dictionary entry break markers do not satisfy the configured prefix and suffix limits.", nameof(entry));
        }

        if (_entries.TryGetValue(normalizedWord, out int[]? existing)) {
            validPoints = existing.Concat(validPoints).Distinct().OrderBy(point => point).ToArray();
        }

        _entries[normalizedWord] = validPoints;
    }
}
