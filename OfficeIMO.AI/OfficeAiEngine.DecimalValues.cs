using System.Globalization;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    // TryParse may succeed after rounding excess precision or underflowing to zero.
    // Compare exact decimal digit/scale representations before admitting normalization.
    private static bool IsExactDecimal(string raw, decimal value, NumberFormatInfo format) =>
        CanonicalDecimal(raw, format) == CanonicalDecimal(value.ToString(CultureInfo.InvariantCulture), NumberFormatInfo.InvariantInfo);

    private static string CanonicalDecimal(string value, NumberFormatInfo format) {
        value = value.Trim();
        bool negative = value.StartsWith(format.NegativeSign, StringComparison.Ordinal)
            || value.EndsWith(format.NegativeSign, StringComparison.Ordinal);
        if (value.StartsWith(format.NegativeSign, StringComparison.Ordinal)) value = value[format.NegativeSign.Length..];
        else if (negative) value = value[..^format.NegativeSign.Length];
        else if (format.PositiveSign.Length > 0 && value.StartsWith(format.PositiveSign, StringComparison.Ordinal))
            value = value[format.PositiveSign.Length..];
        else if (format.PositiveSign.Length > 0 && value.EndsWith(format.PositiveSign, StringComparison.Ordinal))
            value = value[..^format.PositiveSign.Length];
        if (format.NumberGroupSeparator is "\u00a0" or "\u202f") value = value.Replace(" ", format.NumberGroupSeparator, StringComparison.Ordinal);
        if (format.NumberGroupSeparator.Length > 0 && format.NumberGroupSeparator != format.NumberDecimalSeparator)
            value = value.Replace(format.NumberGroupSeparator, string.Empty, StringComparison.Ordinal);
        int point = value.IndexOf(format.NumberDecimalSeparator, StringComparison.Ordinal);
        int scale = point < 0 ? 0 : value.Length - point - format.NumberDecimalSeparator.Length;
        string digits = point < 0 ? value : value.Remove(point, format.NumberDecimalSeparator.Length);
        digits = digits.TrimStart('0');
        if (digits.Length == 0) return "0";
        int end = digits.Length;
        while (scale > 0 && digits[end - 1] == '0') { end--; scale--; }
        return (negative ? "-" : string.Empty) + digits[..end] + ":" + scale.ToString(CultureInfo.InvariantCulture);
    }
}
