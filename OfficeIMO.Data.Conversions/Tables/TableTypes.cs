namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Type inference for untyped text (CSV fields, XML text). Only canonical spellings count, so <c>007</c>, <c>+1</c>,
/// <c>1,000</c>, <c>yes</c> and <c>10/11/2026</c> stay text: identifiers, ZIP codes and phone numbers survive.
/// </summary>
internal static class TableTypes {
    /// <summary>Classifies one value; null for empty text.</summary>
    public static DataScalarKind? Infer(ReadOnlySpan<char> text, char decimalSeparator) {
        if (text.Length == 0) return null;
        if (YamlScalarClassifier.IsCanonicalInteger(text)) return DataScalarKind.Integer;
        if (IsCanonicalDecimal(text, decimalSeparator)) return DataScalarKind.Float;
        if (text.Equals("true".AsSpan(), StringComparison.OrdinalIgnoreCase) || text.Equals("false".AsSpan(), StringComparison.OrdinalIgnoreCase)) {
            return DataScalarKind.Boolean;
        }

        if (TryDate(text, out DataScalarKind kind)) return kind;
        return DataScalarKind.String;
    }

    /// <summary>Whether text written to CSV would read back as something other than text.</summary>
    public static bool LooksTyped(ReadOnlySpan<char> text) {
        if (text.Length == 0) return false;
        char c = text[0];
        if (!((c >= '0' && c <= '9') || c == '-' || c == 't' || c == 'T' || c == 'f' || c == 'F')) return false;
        return Infer(text, '.') != DataScalarKind.String;
    }

    /// <summary>Normalizes an inferred number to JSON grammar.</summary>
    public static string NumberText(string text, char decimalSeparator) =>
        decimalSeparator == '.' ? text : text.Replace(decimalSeparator, '.');

    private static bool IsCanonicalDecimal(ReadOnlySpan<char> s, char separator) {
        int i = 0;
        if (s[0] == '-') i++;
        int integerStart = i;
        while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
        int integerDigits = i - integerStart;
        if (integerDigits == 0 || (integerDigits > 1 && s[integerStart] == '0')) return false;
        bool fraction = false;
        if (i < s.Length && s[i] == separator) {
            i++;
            int start = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (i == start) return false;
            fraction = true;
        }

        bool exponent = false;
        if (i < s.Length && (s[i] == 'e' || s[i] == 'E')) {
            i++;
            if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
            int start = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (i == start) return false;
            exponent = true;
        }

        return i == s.Length && (fraction || exponent);
    }

    // ISO 8601 / RFC 3339 shapes only: 2026-11-10, 2026-11-10T10:00:00[.fff], with Z or ±hh:mm for an offset.
    private static bool TryDate(ReadOnlySpan<char> s, out DataScalarKind kind) {
        kind = default;
        if (s.Length < 10 || !Digits(s, 0, 4) || s[4] != '-' || !Digits(s, 5, 2) || s[7] != '-' || !Digits(s, 8, 2)) return false;
        if (!DateTime.TryParseExact(s.Slice(0, 10).ToString(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _)) return false;
        if (s.Length == 10) {
            kind = DataScalarKind.LocalDate;
            return true;
        }

        if (s[10] != 'T' || s.Length < 19 || !Digits(s, 11, 2) || s[13] != ':' || !Digits(s, 14, 2) || s[16] != ':' || !Digits(s, 17, 2)) return false;
        int i = 19;
        if (i < s.Length && s[i] == '.') {
            i++;
            int start = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (i == start) return false;
        }

        if (i == s.Length) {
            kind = DataScalarKind.LocalDateTime;
            return true;
        }

        if (s[i] == 'Z' && i + 1 == s.Length) {
            kind = DataScalarKind.OffsetDateTime;
            return true;
        }

        if ((s[i] == '+' || s[i] == '-') && s.Length == i + 6 && Digits(s, i + 1, 2) && s[i + 3] == ':' && Digits(s, i + 4, 2)) {
            kind = DataScalarKind.OffsetDateTime;
            return true;
        }

        return false;
    }

    private static bool Digits(ReadOnlySpan<char> s, int start, int count) {
        for (int i = start; i < start + count; i++) {
            if (i >= s.Length || s[i] < '0' || s[i] > '9') return false;
        }

        return true;
    }
}
