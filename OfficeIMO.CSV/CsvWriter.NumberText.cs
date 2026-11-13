#if NET8_0_OR_GREATER
using System.Globalization;
using System.Runtime.InteropServices;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    private static readonly double[] DecimalScales = { 1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7 };
    private const int ShortDecimalMaximumLength = 24;
    internal const int IntegerMaximumLength = 20;

    /// <summary>
    /// Formats with the invariant culture and default format into UTF-16 or UTF-8 units. Integers, doubles that are
    /// short exact decimals and decimals with a 64-bit coefficient are written here, the remaining doubles by
    /// <see cref="CsvShortestDouble"/>; each gives the runtime's text with less work.
    /// </summary>
    internal static bool TryFormatInvariant<T, TChar>(T value, Span<TChar> destination, out int written)
        where T : ISpanFormattable, IUtf8SpanFormattable
        where TChar : unmanaged
    {
        if (typeof(T) == typeof(int) && destination.Length >= IntegerMaximumLength)
        {
            written = FormatInt32((int)(object)value, destination);
            return true;
        }

        if (typeof(T) == typeof(long) && destination.Length >= IntegerMaximumLength)
        {
            written = FormatInt64((long)(object)value, destination);
            return true;
        }

        if (typeof(T) == typeof(double) && destination.Length >= CsvShortestDouble.MaximumLength)
        {
            written = FormatDouble((double)(object)value, destination);
            if (written > 0)
            {
                return true;
            }
        }

        if (typeof(T) == typeof(decimal) && TryFormatSmallDecimal((decimal)(object)value, destination, out written))
        {
            return true;
        }

        return typeof(TChar) == typeof(byte)
            ? value.TryFormat(MemoryMarshal.Cast<TChar, byte>(destination), out written, default, CultureInfo.InvariantCulture)
            : value.TryFormat(MemoryMarshal.Cast<TChar, char>(destination), out written, default, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Writes the invariant text of an integer into at least <see cref="IntegerMaximumLength"/> UTF-16 or UTF-8 units.
    /// The runtime writes non-negative values without consulting the culture; a negative value takes its minus sign
    /// here and its magnitude from the same unsigned path.
    /// </summary>
    internal static int FormatInt64<TChar>(long value, Span<TChar> destination)
        where TChar : unmanaged
    {
        int sign = 0;
        ulong magnitude = (ulong)value;
        if (value < 0)
        {
            destination[0] = CsvShortestDouble.Unit<TChar>('-');
            magnitude = 0 - magnitude;
            sign = 1;
        }

        int digits;
        if (typeof(TChar) == typeof(byte))
        {
            magnitude.TryFormat(MemoryMarshal.Cast<TChar, byte>(destination.Slice(sign)), out digits, default, CultureInfo.InvariantCulture);
        }
        else
        {
            magnitude.TryFormat(MemoryMarshal.Cast<TChar, char>(destination.Slice(sign)), out digits, default, CultureInfo.InvariantCulture);
        }

        return sign + digits;
    }

    /// <inheritdoc cref="FormatInt64{TChar}(long, Span{TChar})"/>
    internal static int FormatInt32<TChar>(int value, Span<TChar> destination)
        where TChar : unmanaged
    {
        int sign = 0;
        uint magnitude = (uint)value;
        if (value < 0)
        {
            destination[0] = CsvShortestDouble.Unit<TChar>('-');
            magnitude = 0 - magnitude;
            sign = 1;
        }

        int digits;
        if (typeof(TChar) == typeof(byte))
        {
            magnitude.TryFormat(MemoryMarshal.Cast<TChar, byte>(destination.Slice(sign)), out digits, default, CultureInfo.InvariantCulture);
        }
        else
        {
            magnitude.TryFormat(MemoryMarshal.Cast<TChar, char>(destination.Slice(sign)), out digits, default, CultureInfo.InvariantCulture);
        }

        return sign + digits;
    }

    /// <summary>
    /// Writes the invariant text of a finite non-zero double into at least <see cref="CsvShortestDouble.MaximumLength"/>
    /// units and returns its length, or 0 for zero, infinities and NaN, which the caller formats as usual.
    /// </summary>
    internal static int FormatDouble<TChar>(double value, Span<TChar> destination)
        where TChar : unmanaged
    {
        if (TryFormatShortDecimal(value, destination, out int written) || CsvShortestDouble.TryFormat(value, destination, out written))
        {
            return written;
        }

        return 0;
    }

    /// <summary>
    /// Writes the invariant shortest round-trip text of a double that is a decimal of at most seven places and
    /// fifteen significant digits, between 0.001 and 10^15, where that text has no exponent. For the smallest scale
    /// k with m = round(|v| * 10^k) and m / 10^k == |v|, m and 10^k are exact doubles and IEEE division rounds
    /// correctly, so m * 10^-k parses back to v. Distinct decimals of at most fifteen significant digits never parse
    /// to the same double, so no shorter text parses to v either: this is the text the runtime writes. Returns false
    /// for any other value, which the caller formats as usual.
    /// </summary>
    internal static bool TryFormatShortDecimal<TChar>(double value, Span<TChar> destination, out int written)
        where TChar : unmanaged
    {
        written = 0;
        double magnitude = Math.Abs(value);
        if (!(magnitude >= 0.001 && magnitude < 1e15) || destination.Length < ShortDecimalMaximumLength)
        {
            return false;
        }

        if (magnitude < 1e8)
        {
            // Below 10^8 the product with 10^7 is within 0.15 of the scaled decimal, so one check at the largest
            // scale answers for every smaller one: k places pass exactly when m * 10^(7-k) passes at seven places,
            // and stripping its trailing zeros gives the smallest scale.
            double scaled = Math.Round(magnitude * 1e7);
            if (scaled / 1e7 != magnitude)
            {
                return false;
            }

            long fixedDigits = (long)scaled;
            int fixedScale = DecimalScales.Length - 1;
            while (fixedScale > 0)
            {
                long quotient = fixedDigits / 10;
                if (fixedDigits - quotient * 10 != 0)
                {
                    break;
                }

                fixedDigits = quotient;
                fixedScale--;
            }

            written = WriteFixed(value < 0, (ulong)fixedDigits, fixedScale, destination);
            return true;
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
                written = WriteFixed(value < 0, (ulong)digits, scale, destination);
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
    internal static bool TryFormatSmallDecimal<TChar>(decimal value, Span<TChar> destination, out int written)
        where TChar : unmanaged
    {
        written = 0;
        Span<int> bits = stackalloc int[4];
        decimal.GetBits(value, bits);
        ulong coefficient = ((ulong)(uint)bits[1] << 32) | (uint)bits[0];
        if (bits[2] != 0 || coefficient == 0 || destination.Length < ShortDecimalMaximumLength + 8)
        {
            return false;
        }

        written = WriteFixed(bits[3] < 0, coefficient, (bits[3] >> 16) & 0xFF, destination);
        return true;
    }

    /// <summary>Writes digits * 10^-scale with every fraction digit, as "0.0dd" below one.</summary>
    private static int WriteFixed<TChar>(bool negative, ulong digits, int scale, Span<TChar> destination)
        where TChar : unmanaged
    {
        int position = 0;
        if (negative)
        {
            destination[position++] = CsvShortestDouble.Unit<TChar>('-');
        }

        int digitCount = CsvShortestDouble.CountDigits(digits);
        if (digitCount > scale)
        {
            return CsvShortestDouble.WriteDigits(destination, position, digits, digitCount, scale);
        }

        destination[position++] = CsvShortestDouble.Unit<TChar>('0');
        destination[position++] = CsvShortestDouble.Unit<TChar>('.');
        int zeros = scale - digitCount;
        destination.Slice(position, zeros).Fill(CsvShortestDouble.Unit<TChar>('0'));
        return CsvShortestDouble.WriteDigits(destination, position + zeros, digits, digitCount, 0);
    }
}
#endif
