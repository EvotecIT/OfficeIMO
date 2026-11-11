#if NET8_0_OR_GREATER
using System.Globalization;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    private static readonly double[] DecimalScales = { 1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7 };
    private const int ShortDecimalMaximumLength = 24;

    /// <summary>
    /// Formats with the invariant culture and default format. Doubles that are short exact decimals are written
    /// by <see cref="TryFormatShortDecimal"/>, which gives the same text without the general shortest-digit search.
    /// </summary>
    internal static bool TryFormatInvariant<T>(T value, Span<char> destination, out int written)
        where T : ISpanFormattable
    {
        if (typeof(T) == typeof(double) && TryFormatShortDecimal((double)(object)value, destination, out written))
        {
            return true;
        }

        return value.TryFormat(destination, out written, default, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Writes the invariant shortest round-trip text of a double that is a decimal of at most seven places and
    /// fifteen significant digits, between 0.001 and 10^15, where that text has no exponent. For the smallest scale
    /// k with m = round(|v| * 10^k) and m / 10^k == |v|, m and 10^k are exact doubles and IEEE division rounds
    /// correctly, so m * 10^-k parses back to v. Distinct decimals of at most fifteen significant digits never parse
    /// to the same double, so no shorter text parses to v either: this is the text the runtime writes. Returns false
    /// for any other value, which the caller formats as usual.
    /// </summary>
    internal static bool TryFormatShortDecimal(double value, Span<char> destination, out int written)
    {
        written = 0;
        double magnitude = Math.Abs(value);
        if (!(magnitude >= 0.001 && magnitude < 1e15) || destination.Length < ShortDecimalMaximumLength)
        {
            return false;
        }

        for (var scale = 0; scale < DecimalScales.Length; scale++)
        {
            double digits = Math.Round(magnitude * DecimalScales[scale]);
            if (digits >= 1e15)
            {
                return false;
            }

            if (digits / DecimalScales[scale] == magnitude)
            {
                written = WriteScaledDecimal(value < 0, (long)digits, scale, destination);
                return true;
            }
        }

        return false;
    }

    /// <summary>Writes digits * 10^-scale; the smallest scale leaves no trailing fraction zeros.</summary>
    private static int WriteScaledDecimal(bool negative, long digits, int scale, Span<char> destination)
    {
        var digitCount = 1;
        for (long bound = 10; digitCount < 15 && digits >= bound; bound *= 10)
        {
            digitCount++;
        }

        int integerDigits = digitCount - scale;
        int length = (negative ? 1 : 0) + (scale == 0 ? digitCount : integerDigits > 0 ? digitCount + 1 : scale + 2);
        int position = length;
        for (var place = 0; place < scale; place++)
        {
            destination[--position] = (char)('0' + (int)(digits % 10));
            digits /= 10;
        }

        if (scale > 0)
        {
            destination[--position] = '.';
        }

        if (integerDigits > 0)
        {
            do
            {
                destination[--position] = (char)('0' + (int)(digits % 10));
                digits /= 10;
            }
            while (digits != 0);
        }
        else
        {
            destination[--position] = '0';
        }

        if (negative)
        {
            destination[--position] = '-';
        }

        return length;
    }
}
#endif
