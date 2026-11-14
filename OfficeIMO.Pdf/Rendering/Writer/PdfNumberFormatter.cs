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

    private static readonly double[] DecimalScales = { 1D, 10D, 100D, 1000D, 10000D, 100000D, 1000000D };

    /// <summary>
    /// Formats <paramref name="value"/> exactly as <c>value.ToString("0.###…", InvariantCulture)</c> with
    /// <paramref name="decimals"/> optional digits. A value that is the nearest double to a decimal with at most
    /// that many fraction digits and 15 significant digits is written from its integer scaled form; other values
    /// use the framework formatter. .NET Framework formats doubles slowly, and content streams are full of them.
    /// </summary>
    internal static string FixedDecimals(double value, int decimals) {
        if (BitConverter.DoubleToInt64Bits(value) == 0L) return "0";
        double scale = DecimalScales[decimals];
        double scaled = value * scale;
        if (scaled > -1E15D && scaled < 1E15D) {
            long units = (long)Math.Round(scaled);
            if (units != 0 && units / scale == value) {
                return FormatScaled(units, decimals);
            }
        }
        return value.ToString(FixedDecimalFormats[decimals], CultureInfo.InvariantCulture);
    }

    private static readonly string[] FixedDecimalFormats = { "0", "0.#", "0.##", "0.###", "0.####", "0.#####", "0.######" };

    private static string FormatScaled(long units, int decimals) {
        bool negative = units < 0;
        ulong magnitude = negative ? (ulong)(-units) : (ulong)units;
        int fraction = decimals;
        while (fraction > 0 && magnitude % 10UL == 0UL) {
            magnitude /= 10UL;
            fraction--;
        }

        // 15 integer digits, a point, up to six fraction digits, a sign and a leading zero.
        char[] buffer = new char[24];
        int position = buffer.Length;
        for (int digit = 0; digit < fraction; digit++) {
            buffer[--position] = (char)('0' + (int)(magnitude % 10UL));
            magnitude /= 10UL;
        }
        if (fraction > 0) buffer[--position] = '.';
        do {
            buffer[--position] = (char)('0' + (int)(magnitude % 10UL));
            magnitude /= 10UL;
        } while (magnitude != 0UL);
        if (negative) buffer[--position] = '-';
        return new string(buffer, position, buffer.Length - position);
    }
}
