namespace OfficeIMO.Project;

/// <summary>Runtime-independent conversion of native floating-point values to decimal.</summary>
internal static class ProjectNumber {
    private static readonly decimal[] PowersOfTwo = CreatePowers(2m, 96);
    private static readonly decimal[] PowersOfTen = CreatePowers(10m, 29);

    /// <summary>
    /// Converts a double to the decimal it stands for: its exact binary value rounded to
    /// 15 significant digits, ties to even. A plain cast rounds to 15 digits up to .NET 10
    /// but keeps the full binary value from .NET 11, and G15 formatting rounds ties
    /// differently on .NET Framework, so neither gives the same value on every runtime.
    /// </summary>
    internal static decimal ToDecimal(double value) {
        // the same failure as the cast, so callers handle a non-finite or out-of-range value unchanged;
        // (double)decimal.MaxValue rounds up to 2^96, which is already out of range
        if (double.IsNaN(value) || double.IsInfinity(value) || value >= (double)decimal.MaxValue || value <= (double)decimal.MinValue)
            throw new OverflowException("Value was either too large or too small for a Decimal.");
        return RoundToFifteenDigits(ExactValue(value));
    }

    // mantissa * 2^exponent in decimal arithmetic; a division rounds only beyond the 28th digit
    private static decimal ExactValue(double value) {
        long bits = BitConverter.DoubleToInt64Bits(value);
        int exponent = (int)((bits >> 52) & 0x7FF);
        long mantissa = bits & 0xFFFFFFFFFFFFFL;
        if (exponent == 0) exponent = 1; else mantissa |= 1L << 52;
        exponent -= 1075;
        decimal result = mantissa;
        if (exponent > 0) {
            result *= PowersOfTwo[exponent];
        } else {
            for (int shift = -exponent; shift > 0 && result != 0m;) {
                int step = Math.Min(shift, PowersOfTwo.Length - 1);
                result /= PowersOfTwo[step];
                shift -= step;
            }
        }
        return bits < 0 ? -result : result;
    }

    private static decimal RoundToFifteenDigits(decimal value) {
        if (value == 0m) return 0m;
        decimal magnitude = Math.Abs(value);
        int exponent; // 10^exponent <= magnitude < 10^(exponent + 1)
        if (magnitude >= 1m) {
            exponent = 0;
            while (exponent < 28 && magnitude >= PowersOfTen[exponent + 1]) exponent++;
        } else {
            exponent = -1;
            while (exponent > -28 && magnitude * PowersOfTen[-exponent] < 1m) exponent--;
        }
        int decimals = 14 - exponent;
        if (decimals >= 0) {
            value = decimal.Round(value, Math.Min(decimals, 28), MidpointRounding.ToEven);
        } else {
            decimal scale = PowersOfTen[-decimals];
            value = decimal.Round(value / scale, 0, MidpointRounding.ToEven) * scale;
        }
        // drop trailing zeros, as the cast did
        return value / 1.0000000000000000000000000000m;
    }

    private static decimal[] CreatePowers(decimal radix, int count) {
        var powers = new decimal[count];
        powers[0] = 1m;
        for (int i = 1; i < count; i++) powers[i] = powers[i - 1] * radix;
        return powers;
    }
}
