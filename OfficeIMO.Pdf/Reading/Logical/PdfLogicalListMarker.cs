namespace OfficeIMO.Pdf;

/// <summary>Numbering style shown by a PDF list item's marker.</summary>
public enum PdfLogicalListNumberFormat {
    /// <summary>A bullet or other symbol, or no visible marker: an unordered item.</summary>
    Bullet,

    /// <summary>Arabic numbers such as "1." or "2)".</summary>
    Number,

    /// <summary>Lowercase letters such as "a." or "(b)".</summary>
    LowerLetter,

    /// <summary>Uppercase letters such as "A." or "(B)".</summary>
    UpperLetter,

    /// <summary>Lowercase Roman numerals such as "iv." or "(ii)".</summary>
    LowerRoman,

    /// <summary>Uppercase Roman numerals such as "IV." or "(II)".</summary>
    UpperRoman
}

/// <summary>Classifies list markers ("•", "3.", "(b)", "iv)") into a numbering style, ordinal and punctuation.</summary>
internal readonly struct PdfLogicalListMarker {
    private PdfLogicalListMarker(PdfLogicalListNumberFormat format, int? ordinal, string prefix, string suffix, bool includesParentNumbers = false, int[]? parentNumbers = null) {
        IncludesParentNumbers = includesParentNumbers;
        ParentNumbers = parentNumbers ?? Array.Empty<int>();
        Format = format;
        Ordinal = ordinal;
        Prefix = prefix;
        Suffix = suffix;
    }

    internal PdfLogicalListNumberFormat Format { get; }

    /// <summary>Position of the item in its list ("c)" is 3); null for bullets.</summary>
    internal int? Ordinal { get; }

    /// <summary>Punctuation before the number, such as "(".</summary>
    internal string Prefix { get; }

    /// <summary>Punctuation after the number, such as "." or ")".</summary>
    internal string Suffix { get; }

    internal bool IsOrdered => Format != PdfLogicalListNumberFormat.Bullet;

    /// <summary>True for outline numbers such as "1.2" that repeat the parent item's number.</summary>
    internal bool IncludesParentNumbers { get; }

    /// <summary>The parent numbers an outline marker shows: 2 for "2.5"; empty when the marker shows none.</summary>
    internal IReadOnlyList<int> ParentNumbers { get; }

    internal static PdfLogicalListMarker Parse(string? marker) {
        string trimmed = marker?.Trim() ?? string.Empty;
        int start = 0;
        int end = trimmed.Length;
        while (start < end && trimmed[start] is '(' or '[' or '\uFF08') start++;
        while (end > start && trimmed[end - 1] is '.' or ')' or ']' or ':' or '\uFF0E' or '\uFF09' or '\u3001') end--;
        if (end <= start) return Bullet;

        string prefix = trimmed.Substring(0, start);
        string suffix = trimmed.Substring(end);
        string core = trimmed.Substring(start, end - start);
        // "1.2.3" is the third item of a nested list.
        int lastSeparator = core.Length - 1;
        while (lastSeparator >= 0 && core[lastSeparator] is not ('.' or '\uFF0E')) lastSeparator--;
        bool includesParentNumbers = lastSeparator >= 0;
        int[]? parentNumbers = null;
        if (includesParentNumbers) {
            parentNumbers = ParseParentNumbers(core.Substring(0, lastSeparator));
            core = core.Substring(lastSeparator + 1);
            if (core.Length == 0) return Bullet;
        }

        if (TryParseDecimal(core, out int number)) return new PdfLogicalListMarker(PdfLogicalListNumberFormat.Number, number, prefix, suffix, includesParentNumbers, parentNumbers);

        // Letters and numerals need punctuation; a bare "o" or "v" is a bullet glyph or a word.
        if (prefix.Length == 0 && suffix.Length == 0) return Bullet;
        if (TryParseRoman(core, out int roman) && (core.Length > 1 || core is "i" or "I")) {
            return new PdfLogicalListMarker(IsLower(core) ? PdfLogicalListNumberFormat.LowerRoman : PdfLogicalListNumberFormat.UpperRoman, roman, prefix, suffix);
        }

        if (TryParseLetters(core, out int letters)) {
            return new PdfLogicalListMarker(IsLower(core) ? PdfLogicalListNumberFormat.LowerLetter : PdfLogicalListNumberFormat.UpperLetter, letters, prefix, suffix);
        }
        return Bullet;
    }

    /// <summary>
    /// Reads "a" to "z", and labels past "z": a repeated letter counts on as Word numbers them ("aa" 27, "bb" 28),
    /// other pairs count in base 26 ("ab" 28).
    /// </summary>
    private static bool TryParseLetters(string value, out int number) {
        number = 0;
        if (value.Length is 0 or > 3) return false;
        char first = value[0];
        bool lower = first is >= 'a' and <= 'z';
        if (!lower && first is not (>= 'A' and <= 'Z')) return false;
        char baseLetter = lower ? 'a' : 'A';
        bool repeated = true;
        foreach (char ch in value) {
            if (ch < baseLetter || ch > baseLetter + 25) return false;
            repeated &= ch == first;
        }
        if (repeated) {
            number = (value.Length - 1) * 26 + first - baseLetter + 1;
            return true;
        }
        foreach (char ch in value) number = number * 26 + ch - baseLetter + 1;
        return true;
    }

    /// <summary>
    /// Reads a single-character marker in the light of a neighbouring item at the same level: "(v)" after "(iv)" is
    /// Roman 5 rather than the letter v, and "(i)" after "(h)" is the letter i rather than Roman 1.
    /// </summary>
    internal PdfLogicalListMarker ResolveBeside(PdfLogicalListMarker neighbour, string? marker) {
        bool letter = Format is PdfLogicalListNumberFormat.LowerLetter or PdfLogicalListNumberFormat.UpperLetter;
        bool roman = Format is PdfLogicalListNumberFormat.LowerRoman or PdfLogicalListNumberFormat.UpperRoman;
        bool neighbourLetter = neighbour.Format is PdfLogicalListNumberFormat.LowerLetter or PdfLogicalListNumberFormat.UpperLetter;
        bool neighbourRoman = neighbour.Format is PdfLogicalListNumberFormat.LowerRoman or PdfLogicalListNumberFormat.UpperRoman;
        if (!(letter && neighbourRoman) && !(roman && neighbourLetter)) return this;

        string core = (marker ?? string.Empty).Trim();
        core = core.Substring(Prefix.Length, core.Length - Prefix.Length - Suffix.Length);
        bool lower = IsLower(core);
        bool sameCase = lower == (neighbour.Format is PdfLogicalListNumberFormat.LowerLetter or PdfLogicalListNumberFormat.LowerRoman);
        if (!sameCase) return this;
        if (letter && TryParseRoman(core, out int value)) {
            return new PdfLogicalListMarker(lower ? PdfLogicalListNumberFormat.LowerRoman : PdfLogicalListNumberFormat.UpperRoman, value, Prefix, Suffix);
        }
        if (roman && core.Length == 1) {
            return new PdfLogicalListMarker(lower ? PdfLogicalListNumberFormat.LowerLetter : PdfLogicalListNumberFormat.UpperLetter,
                char.ToLowerInvariant(core[0]) - 'a' + 1, Prefix, Suffix);
        }
        return this;
    }

    private static int[]? ParseParentNumbers(string value) {
        string[] parts = value.Split('.', '．');
        var numbers = new int[parts.Length];
        for (int index = 0; index < parts.Length; index++) {
            if (!TryParseDecimal(parts[index], out numbers[index])) return null;
        }
        return numbers;
    }

    private static PdfLogicalListMarker Bullet => new PdfLogicalListMarker(PdfLogicalListNumberFormat.Bullet, null, string.Empty, string.Empty);

    private static bool TryParseDecimal(string value, out int number) {
        number = 0;
        if (value.Length == 0) return false;
        // Decimal digits outside the Basic Multilingual Plane ("𝟙") are surrogate pairs.
        int digits = 0;
        for (int index = 0; index < value.Length;) {
            if (++digits > 9 || !ContentStructureExtractor.TryGetDecimalDigit(value, index, out int digit, out int consumed)) return false;
            number = number * 10 + digit;
            index += consumed;
        }
        return true;
    }

    private static bool IsLower(string value) {
        foreach (char ch in value) {
            if (char.IsUpper(ch)) return false;
        }
        return true;
    }

    private static bool TryParseRoman(string value, out int number) {
        number = 0;
        if (value.Length == 0 || value.Length > 12) return false;
        bool lower = char.IsLower(value[0]);
        int previous = 0;
        for (int i = value.Length - 1; i >= 0; i--) {
            char ch = value[i];
            if (char.IsLower(ch) != lower) return false;
            int digit = char.ToUpperInvariant(ch) switch {
                'I' => 1,
                'V' => 5,
                'X' => 10,
                'L' => 50,
                'C' => 100,
                'D' => 500,
                'M' => 1000,
                _ => 0
            };
            if (digit == 0) return false;
            if (digit < previous) {
                number -= digit;
            } else {
                number += digit;
                previous = digit;
            }
        }

        return number > 0 && string.Equals(ToRoman(number), value.ToUpperInvariant(), StringComparison.Ordinal);
    }

    private static string ToRoman(int number) {
        var builder = new StringBuilder();
        (int Value, string Text)[] numerals = {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };
        foreach ((int value, string text) in numerals) {
            while (number >= value) {
                builder.Append(text);
                number -= value;
            }
        }
        return builder.ToString();
    }
}
