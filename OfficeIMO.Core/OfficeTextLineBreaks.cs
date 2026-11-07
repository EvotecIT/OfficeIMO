using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Dependency-free Unicode line-break opportunities shared by OfficeIMO renderers.</summary>
/// <remarks>
/// Opportunities follow the Unicode Line Breaking Algorithm (UAX #14, Unicode 16.0.0 data):
/// spaces, hyphens, solidus, zero-width space and CJK boundaries, with no-break, joiner, quote,
/// numeric-expression and East Asian punctuation constraints. Breaks never split a surrogate pair
/// or grapheme cluster.
/// </remarks>
internal static class OfficeTextLineBreaks {
    /// <summary>
    /// Returns safe UTF-16 indexes where text can wrap without inserting characters.
    /// </summary>
    public static IReadOnlyList<int> GetBreakPositions(string? text) =>
        GetBreakPositions(text, allowCjkBreaks: true);

    /// <summary>
    /// Returns safe UTF-16 indexes where text can wrap without inserting characters,
    /// optionally excluding boundaries that are not after a space, hyphen, solidus or other break character
    /// (or before a break-before character).
    /// </summary>
    public static IReadOnlyList<int> GetBreakPositions(string? text, bool allowCjkBreaks) =>
        GetBreakPositions(text, allowCjkBreaks, default);

    /// <summary>Returns safe UTF-16 wrap indexes, observing cancellation during Unicode boundary analysis.</summary>
    /// <param name="text">Text whose grapheme and line-break boundaries are analyzed.</param>
    /// <param name="allowCjkBreaks">Whether implicit East Asian line-break opportunities may be returned.</param>
    /// <param name="cancellationToken">Cancellation for the calling layout or rendering operation.</param>
    public static IReadOnlyList<int> GetBreakPositions(string? text, bool allowCjkBreaks, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (text == null || text.Length == 0) {
            return Array.Empty<int>();
        }

        if (TryGetSimpleAsciiBreakPositions(text, cancellationToken, out IReadOnlyList<int> simpleAsciiPositions))
            return simpleAsciiPositions;

        int[]? elementStarts = null;
        int[]? scalarStarts = null;
        int[] scalars;
        if (OfficeTextElements.HasOnlySingleCodeUnitElements(text, cancellationToken)) {
            if (text.Length < 2) return Array.Empty<int>();
            scalars = new int[text.Length];
            for (int index = 0; index < text.Length; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                scalars[index] = text[index];
            }
        } else {
            elementStarts = StringInfo.ParseCombiningCharacters(text);
            if (elementStarts.Length < 2) return Array.Empty<int>();
            scalars = OfficeUnicodeBidi.ToScalars(text, out scalarStarts);
        }
        OfficeUnicodeLineBreaker.BreakAction[] actions = OfficeUnicodeLineBreaker.Analyze(scalars, cancellationToken);
        var positions = new List<int>();
        for (int index = 1; index < scalars.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (actions[index] == OfficeUnicodeLineBreaker.BreakAction.Prohibited) continue;
            // A soft hyphen breaks only by inserting a visible hyphen; discretionary hyphenation owns it.
            if (scalars[index - 1] == 0x00AD) continue;
            int boundary = scalarStarts == null ? index : scalarStarts[index];
            if (elementStarts != null && Array.BinarySearch(elementStarts, boundary) < 0) continue;
            if (!allowCjkBreaks && actions[index] != OfficeUnicodeLineBreaker.BreakAction.Mandatory &&
                !IsExplicitBreakAfter(scalars[index - 1]) && !IsExplicitBreakBefore(scalars[index])) continue;
            positions.Add(boundary);
        }

        return positions.ToArray();
    }

    // AL/NU pairs never break; ordinary spaces permit a break only before the next word.
    // Validate the whole subset before allocating output so all other text retains the shared UAX path.
    private static bool TryGetSimpleAsciiBreakPositions(string text, System.Threading.CancellationToken cancellationToken,
        out IReadOnlyList<int> positions) {
        int count = 0;
        for (int index = 0; index < text.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            char value = text[index];
            if (value != ' ' && !OfficeIMO.Internal.OfficeRuntimeCompatibility.IsAsciiLetterOrDigit(value)) {
                positions = Array.Empty<int>();
                return false;
            }
            if (index > 0 && text[index - 1] == ' ' && value != ' ') count++;
        }
        if (count == 0) {
            positions = Array.Empty<int>();
            return true;
        }
        var found = new int[count];
        int next = 0;
        for (int index = 1; index < text.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (text[index - 1] == ' ' && text[index] != ' ') found[next++] = index;
        }
        positions = found;
        return true;
    }

    // Breaks that remain without CJK boundaries: after spaces and explicit break characters.
    private static bool IsExplicitBreakAfter(int scalar) => OfficeUnicodeTextData.GetLineBreakClass(scalar) is
        OfficeLineBreakClass.SP or OfficeLineBreakClass.BA or OfficeLineBreakClass.HY or OfficeLineBreakClass.ZW or
        OfficeLineBreakClass.B2 or OfficeLineBreakClass.SY or OfficeLineBreakClass.BK or OfficeLineBreakClass.CB;

    private static bool IsExplicitBreakBefore(int scalar) =>
        OfficeUnicodeTextData.GetLineBreakClass(scalar) == OfficeLineBreakClass.BB;

    /// <summary>Returns true when an index is a non-edge grapheme-cluster boundary.</summary>
    public static bool IsValidBreakPosition(string? text, int position) {
        if (string.IsNullOrEmpty(text) || position <= 0 || position >= text!.Length) {
            return false;
        }

        int[] elementStarts = StringInfo.ParseCombiningCharacters(text);
        return Array.BinarySearch(elementStarts, position) >= 0;
    }
}
