using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfNumberFormatter {
    // Gradient coordinates may be tiny before an affine map magnifies them.
    // Preserve the double without exponent notation, which PDF reals forbid.
    internal static string Precise(double value) {
        if (value == 0D) return "0";
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(nameof(value), "PDF numbers must be finite.");
        }
        string text = value.ToString("R", CultureInfo.InvariantCulture);
        int exponentAt = text.IndexOf('E');
        if (exponentAt < 0) exponentAt = text.IndexOf('e');
        if (exponentAt < 0) return text;
        bool negative = text[0] == '-';
        int start = negative ? 1 : 0;
        string mantissa = text.Substring(start, exponentAt - start);
        int point = mantissa.IndexOf('.');
        if (point < 0) point = mantissa.Length;
        string digits = mantissa.Replace(".", string.Empty);
#if NET6_0_OR_GREATER
        point += int.Parse(text.AsSpan(exponentAt + 1), NumberStyles.Integer, CultureInfo.InvariantCulture);
#else
        point += int.Parse(text.Substring(exponentAt + 1), NumberStyles.Integer, CultureInfo.InvariantCulture);
#endif
        string expanded = point <= 0 ? "0." + new string('0', -point) + digits
            : point >= digits.Length ? digits + new string('0', point - digits.Length)
            : digits.Insert(point, ".");
        return negative ? "-" + expanded : expanded;
    }
}
