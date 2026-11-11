#if NET8_0_OR_GREATER
using System.Globalization;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    private static readonly double[] DecimalScales = { 1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7 };
    private const int ShortDecimalMaximumLength = 24;

    /// <summary>
    /// Formats with the invariant culture and default format. Doubles that are short exact decimals and decimals
    /// with a 64-bit coefficient are written by <see cref="TryFormatShortDecimal"/> and
    /// <see cref="TryFormatSmallDecimal"/>, which give the same text with less work.
    /// </summary>
    internal static bool TryFormatInvariant<T>(T value, Span<char> destination, out int written)
        where T : ISpanFormattable
    {
        if (typeof(T) == typeof(double) && (TryFormatShortDecimal((double)(object)value, destination, out written)
            || CsvShortestDouble.TryFormat((double)(object)value, destination, out written)))
        {
            return true;
        }

        if (typeof(T) == typeof(decimal) && TryFormatSmallDecimal((decimal)(object)value, destination, out written))
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

    /// <summary>
    /// Writes the invariant text of a non-zero decimal whose 96-bit coefficient fits in 64 bits: the coefficient's
    /// digits with the point placed by the scale, trailing zeros kept, as the runtime writes it. Returns false
    /// for zero and larger coefficients, which the caller formats as usual.
    /// </summary>
    internal static bool TryFormatSmallDecimal(decimal value, Span<char> destination, out int written)
    {
        written = 0;
        Span<int> bits = stackalloc int[4];
        decimal.GetBits(value, bits);
        ulong coefficient = ((ulong)(uint)bits[1] << 32) | (uint)bits[0];
        if (bits[2] != 0 || coefficient == 0 || destination.Length < ShortDecimalMaximumLength + 8)
        {
            return false;
        }

        int scale = (bits[3] >> 16) & 0xFF;
        int position = 0;
        if (bits[3] < 0)
        {
            destination[position++] = '-';
        }

        // The coefficient's digits are written in place, then the fraction moves right for the point.
        coefficient.TryFormat(destination[position..], out int digitCount, default, CultureInfo.InvariantCulture);
        int integerDigits = digitCount - scale;
        if (scale == 0)
        {
            written = position + digitCount;
        }
        else if (integerDigits > 0)
        {
            int point = position + integerDigits;
            destination.Slice(point, scale).CopyTo(destination[(point + 1)..]);
            destination[point] = '.';
            written = position + digitCount + 1;
        }
        else
        {
            int shift = 2 - integerDigits;
            destination.Slice(position, digitCount).CopyTo(destination[(position + shift)..]);
            destination[position] = '0';
            destination[position + 1] = '.';
            destination.Slice(position + 2, -integerDigits).Fill('0');
            written = position + shift + digitCount;
        }

        return true;
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
