using System.Globalization;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.TextAnalysis;
#else
namespace OfficeIMO.TextAnalysis;
#endif

/// <summary>
/// Word and extended grapheme cluster boundaries as defined by Unicode Standard Annex #29 (Unicode Text
/// Segmentation), using property tables generated from the pinned Unicode Character Database version.
/// </summary>
/// <remarks>
/// Offsets are UTF-16 code unit indexes into the input string. Boundaries never fall inside a surrogate pair; an
/// unpaired surrogate is treated as one code point. The default (untailored) rules are used, so scripts that need a
/// dictionary for word segmentation, such as Thai, Lao, Khmer or Myanmar, break per grapheme cluster, and each Han
/// ideograph is a separate segment.
/// </remarks>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
static class UnicodeTextSegmenter {
    /// <summary>Unicode version the segmentation tables were generated from.</summary>
    public static string UnicodeVersion => UnicodeSegmentationData.UnicodeVersion;

    /// <summary>Returns every word boundary offset, including 0 and the text length, for non-empty text.</summary>
    public static IReadOnlyList<int> GetWordBoundaries(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return ToOffsets(new ScalarText(text), FindWordBreaks);
    }

    /// <summary>Returns every extended grapheme cluster boundary offset, including 0 and the text length, for non-empty text.</summary>
    public static IReadOnlyList<int> GetGraphemeClusterBoundaries(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return ToOffsets(new ScalarText(text), FindGraphemeBreaks);
    }

    /// <summary>Splits text into the segments between consecutive UAX #29 word boundaries, in order.</summary>
    public static IReadOnlyList<UnicodeTextSegment> GetWordSegments(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var scalars = new ScalarText(text);
        bool[] breaks = FindWordBreaks(scalars);
        var segments = new List<UnicodeTextSegment>();
        int segmentStart = 0;
        bool wordLike = false;
        for (int index = 0; index < scalars.Count; index++) {
            if (index > 0 && breaks[index]) {
                segments.Add(new UnicodeTextSegment(scalars.Offsets[segmentStart],
                    scalars.Offsets[index] - scalars.Offsets[segmentStart], wordLike));
                segmentStart = index;
                wordLike = false;
            }
            wordLike |= IsWordLikeScalar(text, scalars, index);
        }
        if (scalars.Count > 0) {
            segments.Add(new UnicodeTextSegment(scalars.Offsets[segmentStart],
                text.Length - scalars.Offsets[segmentStart], wordLike));
        }
        return segments.AsReadOnly();
    }

    /// <summary>Splits text into extended grapheme clusters, in order.</summary>
    public static IReadOnlyList<UnicodeTextSegment> GetGraphemeClusters(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        IReadOnlyList<int> offsets = GetGraphemeClusterBoundaries(text);
        var clusters = new UnicodeTextSegment[Math.Max(0, offsets.Count - 1)];
        for (int index = 0; index < clusters.Length; index++) {
            clusters[index] = new UnicodeTextSegment(offsets[index], offsets[index + 1] - offsets[index], false);
        }
        return clusters;
    }

    private static IReadOnlyList<int> ToOffsets(ScalarText scalars, Func<ScalarText, bool[]> rules) {
        if (scalars.Count == 0) return Array.Empty<int>();
        bool[] breaks = rules(scalars);
        var offsets = new List<int> { 0 };
        for (int index = 1; index < scalars.Count; index++) {
            if (breaks[index]) offsets.Add(scalars.Offsets[index]);
        }
        offsets.Add(scalars.TextLength);
        return offsets.AsReadOnly();
    }

    private static bool IsWordLikeScalar(string text, ScalarText scalars, int index) {
        switch (scalars.Word[index]) {
            case WordBreak.ALetter:
            case WordBreak.HebrewLetter:
            case WordBreak.Numeric:
            case WordBreak.Katakana:
                return true;
        }
        switch (CharUnicodeInfo.GetUnicodeCategory(text, scalars.Offsets[index])) {
            case UnicodeCategory.UppercaseLetter:
            case UnicodeCategory.LowercaseLetter:
            case UnicodeCategory.TitlecaseLetter:
            case UnicodeCategory.ModifierLetter:
            case UnicodeCategory.OtherLetter:
            case UnicodeCategory.DecimalDigitNumber:
            case UnicodeCategory.LetterNumber:
            case UnicodeCategory.OtherNumber:
                return true;
            default:
                return false;
        }
    }

    // breaks[i] reports whether there is a boundary before scalar i (i > 0); sot and eot always break (WB1, WB2).
    private static bool[] FindWordBreaks(ScalarText s) {
        int count = s.Count;
        var breaks = new bool[count];
        for (int i = 1; i < count; i++) breaks[i] = IsWordBreak(s, i);
        return breaks;
    }

    private static bool IsWordBreak(ScalarText s, int i) {
        WordBreak before = s.Word[i - 1];
        WordBreak after = s.Word[i];
        if (before == WordBreak.CR && after == WordBreak.LF) return false; // WB3
        if (IsNewline(before)) return true; // WB3a
        if (IsNewline(after)) return true; // WB3b
        if (before == WordBreak.Zwj && s.Pictographic[i]) return false; // WB3c
        if (before == WordBreak.WSegSpace && after == WordBreak.WSegSpace) return false; // WB3d
        if (IsIgnorable(after)) return false; // WB4

        int previous = SkipIgnorableBackward(s, i - 1);
        WordBreak left = previous >= 0 ? s.Word[previous] : WordBreak.Other;
        int beforePrevious = previous > 0 ? SkipIgnorableBackward(s, previous - 1) : -1;
        WordBreak leftLeft = beforePrevious >= 0 ? s.Word[beforePrevious] : WordBreak.Other;
        int next = SkipIgnorableForward(s, i + 1);
        WordBreak right = next < s.Count ? s.Word[next] : WordBreak.Other;

        if (IsAHLetter(left) && IsAHLetter(after)) return false; // WB5
        if (IsAHLetter(left) && (after == WordBreak.MidLetter || IsMidNumLetQ(after)) && IsAHLetter(right)) return false; // WB6
        if (IsAHLetter(leftLeft) && (left == WordBreak.MidLetter || IsMidNumLetQ(left)) && IsAHLetter(after)) return false; // WB7
        if (left == WordBreak.HebrewLetter && after == WordBreak.SingleQuote) return false; // WB7a
        if (left == WordBreak.HebrewLetter && after == WordBreak.DoubleQuote && right == WordBreak.HebrewLetter) return false; // WB7b
        if (leftLeft == WordBreak.HebrewLetter && left == WordBreak.DoubleQuote && after == WordBreak.HebrewLetter) return false; // WB7c
        if (left == WordBreak.Numeric && after == WordBreak.Numeric) return false; // WB8
        if (IsAHLetter(left) && after == WordBreak.Numeric) return false; // WB9
        if (left == WordBreak.Numeric && IsAHLetter(after)) return false; // WB10
        if (leftLeft == WordBreak.Numeric && (left == WordBreak.MidNum || IsMidNumLetQ(left)) && after == WordBreak.Numeric) return false; // WB11
        if (left == WordBreak.Numeric && (after == WordBreak.MidNum || IsMidNumLetQ(after)) && right == WordBreak.Numeric) return false; // WB12
        if (left == WordBreak.Katakana && after == WordBreak.Katakana) return false; // WB13
        if ((IsAHLetter(left) || left == WordBreak.Numeric || left == WordBreak.Katakana || left == WordBreak.ExtendNumLet)
            && after == WordBreak.ExtendNumLet) return false; // WB13a
        if (left == WordBreak.ExtendNumLet && (IsAHLetter(after) || after == WordBreak.Numeric || after == WordBreak.Katakana)) return false; // WB13b
        if (left == WordBreak.RegionalIndicator && after == WordBreak.RegionalIndicator) {
            int regionalIndicators = 0;
            for (int index = previous; index >= 0; index = index > 0 ? SkipIgnorableBackward(s, index - 1) : -1) {
                if (s.Word[index] != WordBreak.RegionalIndicator) break;
                regionalIndicators++;
            }
            if (regionalIndicators % 2 == 1) return false; // WB15, WB16
        }
        return true; // WB999
    }

    // WB4: an ignorable run after sot, CR, LF or Newline is not absorbed, so it stays visible to the later rules.
    private static int SkipIgnorableBackward(ScalarText s, int index) {
        int start = index;
        while (index >= 0 && IsIgnorable(s.Word[index])) index--;
        if (index < 0 || IsNewline(s.Word[index])) return start;
        return index;
    }

    private static int SkipIgnorableForward(ScalarText s, int index) {
        while (index < s.Count && IsIgnorable(s.Word[index])) index++;
        return index;
    }

    private static bool IsNewline(WordBreak value) =>
        value == WordBreak.CR || value == WordBreak.LF || value == WordBreak.Newline;

    private static bool IsIgnorable(WordBreak value) =>
        value == WordBreak.Extend || value == WordBreak.Format || value == WordBreak.Zwj;

    private static bool IsAHLetter(WordBreak value) =>
        value == WordBreak.ALetter || value == WordBreak.HebrewLetter;

    private static bool IsMidNumLetQ(WordBreak value) =>
        value == WordBreak.MidNumLet || value == WordBreak.SingleQuote;

    private static bool[] FindGraphemeBreaks(ScalarText s) {
        int count = s.Count;
        var breaks = new bool[count];
        int regionalIndicatorRun = s.Grapheme.Length > 0 && s.Grapheme[0] == GraphemeBreak.RegionalIndicator ? 1 : 0;
        for (int i = 1; i < count; i++) {
            breaks[i] = IsGraphemeBreak(s, i, regionalIndicatorRun);
            regionalIndicatorRun = s.Grapheme[i] == GraphemeBreak.RegionalIndicator ? regionalIndicatorRun + 1 : 0;
        }
        return breaks;
    }

    private static bool IsGraphemeBreak(ScalarText s, int i, int regionalIndicatorsBefore) {
        GraphemeBreak before = s.Grapheme[i - 1];
        GraphemeBreak after = s.Grapheme[i];
        if (before == GraphemeBreak.CR && after == GraphemeBreak.LF) return false; // GB3
        if (before == GraphemeBreak.Control || before == GraphemeBreak.CR || before == GraphemeBreak.LF) return true; // GB4
        if (after == GraphemeBreak.Control || after == GraphemeBreak.CR || after == GraphemeBreak.LF) return true; // GB5
        if (before == GraphemeBreak.L && (after == GraphemeBreak.L || after == GraphemeBreak.V ||
                                          after == GraphemeBreak.LV || after == GraphemeBreak.LVT)) return false; // GB6
        if ((before == GraphemeBreak.LV || before == GraphemeBreak.V) &&
            (after == GraphemeBreak.V || after == GraphemeBreak.T)) return false; // GB7
        if ((before == GraphemeBreak.LVT || before == GraphemeBreak.T) && after == GraphemeBreak.T) return false; // GB8
        if (after == GraphemeBreak.Extend || after == GraphemeBreak.Zwj) return false; // GB9
        if (after == GraphemeBreak.SpacingMark) return false; // GB9a
        if (before == GraphemeBreak.Prepend) return false; // GB9b
        if (s.Conjunct[i] == IndicConjunctBreak.Consonant && FollowsConjunctLinker(s, i)) return false; // GB9c
        if (before == GraphemeBreak.Zwj && s.Pictographic[i]) {
            int index = i - 2;
            while (index >= 0 && s.Grapheme[index] == GraphemeBreak.Extend) index--;
            if (index >= 0 && s.Pictographic[index]) return false; // GB11
        }
        if (before == GraphemeBreak.RegionalIndicator && after == GraphemeBreak.RegionalIndicator &&
            regionalIndicatorsBefore % 2 == 1) return false; // GB12, GB13
        return true; // GB999
    }

    // GB9c: \p{InCB=Consonant} [\p{InCB=Extend}\p{InCB=Linker}]* \p{InCB=Linker} [\p{InCB=Extend}\p{InCB=Linker}]* × \p{InCB=Consonant}
    private static bool FollowsConjunctLinker(ScalarText s, int i) {
        bool linker = false;
        int index = i - 1;
        for (; index >= 0; index--) {
            IndicConjunctBreak value = s.Conjunct[index];
            if (value == IndicConjunctBreak.Linker) linker = true;
            else if (value != IndicConjunctBreak.Extend) break;
        }
        return linker && index >= 0 && s.Conjunct[index] == IndicConjunctBreak.Consonant;
    }

    private sealed class ScalarText {
        internal ScalarText(string text) {
            TextLength = text.Length;
            var offsets = new List<int>(text.Length);
            for (int index = 0; index < text.Length; index++) {
                offsets.Add(index);
                if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) index++;
            }
            Count = offsets.Count;
            Offsets = offsets.ToArray();
            Word = new WordBreak[Count];
            Grapheme = new GraphemeBreak[Count];
            Conjunct = new IndicConjunctBreak[Count];
            Pictographic = new bool[Count];
            for (int index = 0; index < Count; index++) {
                int offset = Offsets[index];
                char value = text[offset];
                int codePoint = char.IsHighSurrogate(value) && offset + 1 < text.Length && char.IsLowSurrogate(text[offset + 1])
                    ? char.ConvertToUtf32(value, text[offset + 1])
                    : value;
                ushort packed = UnicodeSegmentationData.Lookup(codePoint);
                Word[index] = (WordBreak)(packed & 0x1F);
                Grapheme[index] = (GraphemeBreak)((packed >> 5) & 0xF);
                Conjunct[index] = (IndicConjunctBreak)((packed >> 9) & 0x3);
                Pictographic[index] = (packed & 0x800) != 0;
            }
        }

        internal int Count { get; }
        internal int TextLength { get; }
        internal int[] Offsets { get; }
        internal WordBreak[] Word { get; }
        internal GraphemeBreak[] Grapheme { get; }
        internal IndicConjunctBreak[] Conjunct { get; }
        internal bool[] Pictographic { get; }
    }

    // Order matches WORD_BREAK in Build/Generate-UnicodeSegmentationData.py.
    private enum WordBreak : byte {
        Other, CR, LF, Newline, Extend, Zwj, RegionalIndicator, Format, Katakana, HebrewLetter, ALetter,
        SingleQuote, DoubleQuote, MidNumLet, MidLetter, MidNum, Numeric, ExtendNumLet, WSegSpace
    }

    // Order matches GRAPHEME_BREAK in Build/Generate-UnicodeSegmentationData.py.
    private enum GraphemeBreak : byte {
        Other, CR, LF, Control, Extend, Zwj, RegionalIndicator, Prepend, SpacingMark, L, V, T, LV, LVT
    }

    private enum IndicConjunctBreak : byte { None, Linker, Consonant, Extend }
}

/// <summary>One segment of text between two consecutive UAX #29 boundaries.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
readonly struct UnicodeTextSegment : IEquatable<UnicodeTextSegment> {
    /// <summary>Creates a segment.</summary>
    public UnicodeTextSegment(int start, int length, bool isWordLike) {
        Start = start;
        Length = length;
        IsWordLike = isWordLike;
    }

    /// <summary>Zero-based UTF-16 offset of the segment.</summary>
    public int Start { get; }

    /// <summary>Segment length in UTF-16 code units.</summary>
    public int Length { get; }

    /// <summary>Offset just past the segment.</summary>
    public int End => Start + Length;

    /// <summary>
    /// True when the segment contains a letter or number (Word_Break ALetter, Hebrew_Letter, Numeric or Katakana, or a
    /// letter or number general category). Whitespace, punctuation, symbols and emoji segments are not word-like.
    /// </summary>
    public bool IsWordLike { get; }

    /// <summary>Returns the segment text from the string it was computed for.</summary>
    public string GetText(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return text.Substring(Start, Length);
    }

    /// <inheritdoc />
    public bool Equals(UnicodeTextSegment other) =>
        Start == other.Start && Length == other.Length && IsWordLike == other.IsWordLike;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is UnicodeTextSegment other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => unchecked((Start * 397) ^ (Length * 31) ^ (IsWordLike ? 1 : 0));

    /// <inheritdoc />
    public override string ToString() => "[" + Start + ", " + End + ")" + (IsWordLike ? " word" : string.Empty);
}
