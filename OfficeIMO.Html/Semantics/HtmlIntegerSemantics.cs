namespace OfficeIMO.Html;

/// <summary>Canonical parsing for HTML integer and non-negative-integer microsyntaxes.</summary>
internal static class HtmlIntegerSemantics {
    /// <summary>
    /// Applies HTML's integer parsing rules, including ASCII leading whitespace,
    /// an optional sign, a required digit prefix, and saturation to the managed range.
    /// </summary>
    internal static bool TryParseInteger(string? text, out int value) {
        value = 0;
        if (string.IsNullOrEmpty(text)) return false;

        int position = 0;
        while (position < text!.Length && IsAsciiWhitespace(text[position])) position++;
        bool negative = false;
        if (position < text.Length && (text[position] == '+' || text[position] == '-')) {
            negative = text[position] == '-';
            position++;
        }

        int firstDigit = position;
        long magnitude = 0L;
        long limit = negative ? (long)int.MaxValue + 1L : int.MaxValue;
        while (position < text.Length && IsAsciiDigit(text[position])) {
            int digit = text[position] - '0';
            magnitude = magnitude > (limit - digit) / 10L ? limit : magnitude * 10L + digit;
            position++;
        }
        if (position == firstDigit) return false;

        value = negative
            ? magnitude == (long)int.MaxValue + 1L ? int.MinValue : -(int)magnitude
            : (int)magnitude;
        return true;
    }

    /// <summary>
    /// Applies HTML's non-negative-integer parsing rules. These parsing rules are
    /// intentionally more forgiving than the valid authoring syntax.
    /// </summary>
    internal static bool TryParseNonNegativeInteger(string? text, out int value) =>
        TryParseInteger(text, out value) && value >= 0;

    /// <summary>Parses an HTML non-negative integer and requires a value greater than zero.</summary>
    internal static bool TryParsePositiveInteger(string? text, out int value) =>
        TryParseNonNegativeInteger(text, out value) && value > 0;

    /// <summary>Advances an integer without wrapping at the managed range boundaries.</summary>
    internal static int AdvanceSaturating(int value, int step) {
        if (step > 0) return value == int.MaxValue ? int.MaxValue : value + 1;
        if (step < 0) return value == int.MinValue ? int.MinValue : value - 1;
        return value;
    }

    private static bool IsAsciiDigit(char value) => value >= '0' && value <= '9';

    private static bool IsAsciiWhitespace(char value) =>
        value == '\t' || value == '\n' || value == '\f' || value == '\r' || value == ' ';
}
