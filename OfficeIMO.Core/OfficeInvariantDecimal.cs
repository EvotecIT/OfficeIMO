using System.Globalization;

namespace OfficeIMO;

/// <summary>Exact invariant decimal conversion for data transport boundaries that cannot silently round.</summary>
internal static class OfficeInvariantDecimal {
    internal static bool TryParseExact(string? text, bool allowExponent, out decimal value) {
        value = 0;
        if (string.IsNullOrEmpty(text) || text!.Length > 4096) return false;
        var styles = NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint;
        if (allowExponent) styles |= NumberStyles.AllowExponent;
        if (!decimal.TryParse(text, styles, CultureInfo.InvariantCulture, out var parsed)) return false;
        if (!Canonical(text, out string digits, out long power) ||
            !Canonical(parsed.ToString(CultureInfo.InvariantCulture), out string actual, out long actualPower) ||
            digits != actual || power != actualPower) return false;
        value = parsed;
        return true;
    }

    private static bool Canonical(string text, out string digits, out long power) {
        digits = ""; power = 0;
        int exponentIndex = text.IndexOfAny(new[] { 'e', 'E' });
        if (exponentIndex >= 0) {
            if (!int.TryParse(text.Substring(exponentIndex + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int exponent)) return false;
            power = exponent;
            text = text.Substring(0, exponentIndex);
        }
        bool negative = text[0] == '-';
        if (negative || text[0] == '+') text = text.Substring(1);
        int point = text.IndexOf('.');
        if (point >= 0) { power -= text.Length - point - 1; text = text.Remove(point, 1); }
        text = text.TrimStart('0');
        if (text.Length == 0) { digits = "0"; power = 0; return true; }
        string trimmed = text.TrimEnd('0'); power += text.Length - trimmed.Length;
        digits = (negative ? "-" : "") + trimmed;
        return true;
    }
}
