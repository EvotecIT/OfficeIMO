using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Supplies preferred UTF-16 hyphenation break positions for one unspaced token.</summary>
/// <param name="token">The token whose optional break positions should be resolved.</param>
/// <returns>UTF-16 indexes before which a generated hyphen may be inserted.</returns>
public delegate IReadOnlyList<int> OfficeTextHyphenationCallback(string token);

/// <summary>Immutable, dependency-free dictionary of explicit word hyphenation points shared by OfficeIMO renderers.</summary>
public sealed class OfficeTextHyphenationLexicon {
    private readonly Dictionary<string, int[]> _entries;

    /// <summary>Creates a case-insensitive dictionary from words such as <c>typog-ra-phy</c>.</summary>
    public OfficeTextHyphenationLexicon(
        IEnumerable<string> hyphenatedWords,
        char breakMarker = '-',
        int minimumPrefixLength = 2,
        int minimumSuffixLength = 2) {
        if (hyphenatedWords == null) throw new ArgumentNullException(nameof(hyphenatedWords));
        if (char.IsLetterOrDigit(breakMarker) || char.IsSurrogate(breakMarker) || char.IsWhiteSpace(breakMarker)) {
            throw new ArgumentException("Hyphenation break markers cannot be letters, digits, surrogates, or whitespace.", nameof(breakMarker));
        }
        if (minimumPrefixLength < 1) throw new ArgumentOutOfRangeException(nameof(minimumPrefixLength), "Hyphenation minimum prefix length must be positive.");
        if (minimumSuffixLength < 1) throw new ArgumentOutOfRangeException(nameof(minimumSuffixLength), "Hyphenation minimum suffix length must be positive.");

        BreakMarker = breakMarker;
        MinimumPrefixLength = minimumPrefixLength;
        MinimumSuffixLength = minimumSuffixLength;
        _entries = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
        foreach (string entry in hyphenatedWords) AddEntry(entry);
    }

    /// <summary>Marker removed from source entries.</summary>
    public char BreakMarker { get; }

    /// <summary>Minimum UTF-16 length retained before a break.</summary>
    public int MinimumPrefixLength { get; }

    /// <summary>Minimum UTF-16 length retained after a break.</summary>
    public int MinimumSuffixLength { get; }

    /// <summary>Number of normalized words in the dictionary.</summary>
    public int Count => _entries.Count;

    /// <summary>Returns preferred UTF-16 break positions for a word.</summary>
    public IReadOnlyList<int> GetBreakpoints(string token) {
        if (token == null) throw new ArgumentNullException(nameof(token));
        return _entries.TryGetValue(token, out int[]? points) ? points.ToArray() : Array.Empty<int>();
    }

    /// <summary>Returns true when the normalized word is present.</summary>
    public bool Contains(string token) {
        if (token == null) throw new ArgumentNullException(nameof(token));
        return _entries.ContainsKey(token);
    }

    /// <summary>Creates the callback shape consumed by renderer options.</summary>
    public OfficeTextHyphenationCallback AsCallback() => GetBreakpoints;

    private void AddEntry(string entry) {
        if (string.IsNullOrWhiteSpace(entry)) throw new ArgumentException("Hyphenation dictionary entries cannot be null, empty, or whitespace.", nameof(entry));
        string value = entry.Trim();
        if (value[0] == BreakMarker || value[value.Length - 1] == BreakMarker) {
            throw new ArgumentException("Hyphenation dictionary entries cannot start or end with the break marker.", nameof(entry));
        }

        var word = new StringBuilder(value.Length);
        var points = new List<int>();
        bool previousWasMarker = false;
        foreach (char current in value) {
            if (current == BreakMarker) {
                if (previousWasMarker) throw new ArgumentException("Hyphenation dictionary entries cannot contain adjacent break markers.", nameof(entry));
                points.Add(word.Length);
                previousWasMarker = true;
                continue;
            }
            word.Append(current);
            previousWasMarker = false;
        }

        if (points.Count == 0) throw new ArgumentException("Hyphenation dictionary entries must contain at least one break marker.", nameof(entry));
        string normalizedWord = word.ToString();
        int[] validPoints = points
            .Where(point => point >= MinimumPrefixLength && normalizedWord.Length - point >= MinimumSuffixLength)
            .Where(point => OfficeTextLineBreaks.IsValidBreakPosition(normalizedWord, point))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
        if (validPoints.Length == 0) throw new ArgumentException("Hyphenation dictionary entry break markers do not satisfy the configured limits.", nameof(entry));
        if (_entries.TryGetValue(normalizedWord, out int[]? existing)) validPoints = existing.Concat(validPoints).Distinct().OrderBy(point => point).ToArray();
        _entries[normalizedWord] = validPoints;
    }
}
