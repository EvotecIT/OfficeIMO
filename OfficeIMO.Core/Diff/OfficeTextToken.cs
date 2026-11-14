using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Diff;

/// <summary>Kind of an <see cref="OfficeTextToken"/>.</summary>
public enum OfficeTextTokenKind {
    /// <summary>A run of letters, digits, combining marks or underscores; or one ideograph or kana.</summary>
    Word,
    /// <summary>A run of whitespace, including line breaks.</summary>
    Whitespace,
    /// <summary>One punctuation or symbol grapheme.</summary>
    Punctuation,
    /// <summary>One line including its line terminator.</summary>
    Line,
    /// <summary>One user-perceived character (grapheme cluster).</summary>
    Character
}

/// <summary>A slice of source text produced by <see cref="OfficeTextTokenizer"/>.</summary>
public readonly struct OfficeTextToken : IEquatable<OfficeTextToken> {
    /// <summary>Creates a token.</summary>
    public OfficeTextToken(string text, int start, OfficeTextTokenKind kind) {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        if (start < 0) throw new ArgumentOutOfRangeException(nameof(start));
        Start = start;
        Kind = kind;
    }

    /// <summary>Token text.</summary>
    public string Text { get; }

    /// <summary>Zero-based UTF-16 offset of the token in the source text.</summary>
    public int Start { get; }

    /// <summary>Length in UTF-16 code units.</summary>
    public int Length => Text?.Length ?? 0;

    /// <summary>Token kind.</summary>
    public OfficeTextTokenKind Kind { get; }

    /// <inheritdoc />
    public bool Equals(OfficeTextToken other) => string.Equals(Text, other.Text, StringComparison.Ordinal) && Start == other.Start && Kind == other.Kind;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeTextToken other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => unchecked((StringComparer.Ordinal.GetHashCode(Text ?? string.Empty) * 397) ^ Start);

    /// <inheritdoc />
    public override string ToString() => Text ?? string.Empty;
}

/// <summary>Splits text into lines, words or user-perceived characters for diffing.</summary>
/// <remarks>Tokens cover the source text exactly: concatenating their <see cref="OfficeTextToken.Text"/> values gives the input back.</remarks>
public static class OfficeTextTokenizer {
    /// <summary>Splits text into lines. Each token keeps its line terminator (<c>\r\n</c>, <c>\n</c> or <c>\r</c>).</summary>
    public static IReadOnlyList<OfficeTextToken> Lines(string? text) {
        var tokens = new List<OfficeTextToken>();
        if (string.IsNullOrEmpty(text)) return tokens;
        string value = text!;
        int start = 0;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current != '\n' && current != '\r') continue;
            if (current == '\r' && index + 1 < value.Length && value[index + 1] == '\n') index++;
            tokens.Add(new OfficeTextToken(value.Substring(start, index + 1 - start), start, OfficeTextTokenKind.Line));
            start = index + 1;
        }

        if (start < value.Length) tokens.Add(new OfficeTextToken(value.Substring(start), start, OfficeTextTokenKind.Line));
        return tokens;
    }

    /// <summary>
    /// Splits text into words, whitespace runs and single punctuation marks. Grapheme clusters are never split, and each
    /// ideograph or kana is its own word because those scripts do not separate words with spaces.
    /// </summary>
    public static IReadOnlyList<OfficeTextToken> Words(string? text) {
        var tokens = new List<OfficeTextToken>();
        if (string.IsNullOrEmpty(text)) return tokens;
        string value = text!;
        int runStart = -1;
        OfficeTextTokenKind runKind = OfficeTextTokenKind.Word;
        int offset = 0;
        foreach (string element in OfficeTextElements.Enumerate(value)) {
            OfficeTextTokenKind kind = Classify(element, out bool standalone);
            if (runStart >= 0 && (kind != runKind || standalone || kind == OfficeTextTokenKind.Punctuation)) {
                tokens.Add(new OfficeTextToken(value.Substring(runStart, offset - runStart), runStart, runKind));
                runStart = -1;
            }

            if (standalone || kind == OfficeTextTokenKind.Punctuation) {
                tokens.Add(new OfficeTextToken(element, offset, kind));
            } else if (runStart < 0) {
                runStart = offset;
                runKind = kind;
            }

            offset += element.Length;
        }

        if (runStart >= 0) tokens.Add(new OfficeTextToken(value.Substring(runStart), runStart, runKind));
        return tokens;
    }

    /// <summary>Splits text into user-perceived characters (grapheme clusters), so combining sequences, surrogate pairs and emoji stay whole.</summary>
    public static IReadOnlyList<OfficeTextToken> Characters(string? text) {
        var tokens = new List<OfficeTextToken>();
        if (string.IsNullOrEmpty(text)) return tokens;
        int offset = 0;
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            tokens.Add(new OfficeTextToken(element, offset, OfficeTextTokenKind.Character));
            offset += element.Length;
        }

        return tokens;
    }

    private static OfficeTextTokenKind Classify(string element, out bool standalone) {
        standalone = false;
        int codePoint = char.IsSurrogatePair(element, 0) ? char.ConvertToUtf32(element[0], element[1]) : element[0];
        if (element[0] == '_') return OfficeTextTokenKind.Word;
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(element, 0);
        switch (category) {
            case UnicodeCategory.SpaceSeparator:
            case UnicodeCategory.LineSeparator:
            case UnicodeCategory.ParagraphSeparator:
                return OfficeTextTokenKind.Whitespace;
            case UnicodeCategory.Control:
                return char.IsWhiteSpace(element[0]) ? OfficeTextTokenKind.Whitespace : OfficeTextTokenKind.Punctuation;
            case UnicodeCategory.UppercaseLetter:
            case UnicodeCategory.LowercaseLetter:
            case UnicodeCategory.TitlecaseLetter:
            case UnicodeCategory.ModifierLetter:
            case UnicodeCategory.OtherLetter:
            case UnicodeCategory.DecimalDigitNumber:
            case UnicodeCategory.LetterNumber:
            case UnicodeCategory.OtherNumber:
            case UnicodeCategory.NonSpacingMark:
            case UnicodeCategory.SpacingCombiningMark:
            case UnicodeCategory.EnclosingMark:
                standalone = IsIdeographicOrKana(codePoint);
                return OfficeTextTokenKind.Word;
            default:
                return OfficeTextTokenKind.Punctuation;
        }
    }

    private static bool IsIdeographicOrKana(int codePoint) =>
        (codePoint >= 0x3040 && codePoint <= 0x30FF) ||
        (codePoint >= 0x31F0 && codePoint <= 0x31FF) ||
        (codePoint >= 0x3400 && codePoint <= 0x4DBF) ||
        (codePoint >= 0x4E00 && codePoint <= 0x9FFF) ||
        (codePoint >= 0xF900 && codePoint <= 0xFAFF) ||
        (codePoint >= 0x20000 && codePoint <= 0x3134F);
}
