#if NET8_0_OR_GREATER
using System.Globalization;
using System.Numerics;

namespace OfficeIMO.CSV;

/// <summary>
/// Writes the invariant text the runtime gives a finite, non-zero double by default: the fewest significant
/// digits that parse back to the same double (the closest such digits when several qualify), in fixed notation
/// for decimal exponents from -4 to 16 and as <c>d.dddE+XX</c> otherwise. The digits come from the Ryu
/// algorithm (Ulf Adams, "Ryū: fast float-to-string conversion", PLDI 2018), which finds them with 128-bit
/// multiplications by powers of five; the tables of those powers are computed once, on first use.
/// </summary>
internal static class CsvShortestDouble
{
    private const int MantissaBits = 52;
    private const int ExponentBias = 1023;
    private const int Pow5Bits = 125;
    private const int Pow5InverseBits = 125;
    private const int Pow5Count = 326;
    private const int Pow5InverseCount = 342;
    internal const int MaximumLength = 25;
    private const string TwoDigits = "00010203040506070809101112131415161718192021222324252627282930313233343536373839404142434445464748495051525354555657585960616263646566676869707172737475767778798081828384858687888990919293949596979899";
    private static readonly ulong[] PowersOfTen = CreatePowersOfTen();

    /// <summary>5^i, normalised to its top <see cref="Pow5Bits"/> bits.</summary>
    private static readonly UInt128[] Pow5 = CreatePow5Table();
    /// <summary>2^j / 5^i rounded up, where j keeps <see cref="Pow5InverseBits"/> significant bits.</summary>
    private static readonly UInt128[] Pow5Inverse = CreatePow5InverseTable();

    internal static bool TryFormat(double value, Span<char> destination, out int written)
    {
        written = 0;
        if (!double.IsFinite(value) || value == 0 || destination.Length < MaximumLength)
        {
            return false;
        }

        ulong bits = BitConverter.DoubleToUInt64Bits(value);
        ToDecimal(bits & ((1UL << MantissaBits) - 1), (int)(bits >> MantissaBits) & 0x7FF, out ulong digits, out int exponent);
        while (digits % 10 == 0)
        {
            digits /= 10;
            exponent++;
        }

        int position = 0;
        if ((long)bits < 0)
        {
            destination[position++] = '-';
        }

        int digitCount = CountDigits(digits);
        int scientificExponent = exponent + digitCount - 1;
        if (scientificExponent is < -4 or >= 17)
        {
            position = WriteDigits(destination, position, digits, digitCount, digitCount - 1);
            destination[position++] = 'E';
            destination[position++] = scientificExponent < 0 ? '-' : '+';
            int magnitude = Math.Abs(scientificExponent);
            if (magnitude >= 100)
            {
                destination[position++] = (char)('0' + magnitude / 100);
                magnitude %= 100;
            }

            destination[position++] = (char)('0' + magnitude / 10);
            destination[position++] = (char)('0' + magnitude % 10);
            written = position;
            return true;
        }

        if (exponent >= 0)
        {
            position = WriteDigits(destination, position, digits, digitCount, 0);
            destination.Slice(position, exponent).Fill('0');
            written = position + exponent;
            return true;
        }

        int integerDigits = digitCount + exponent;
        if (integerDigits > 0)
        {
            written = WriteDigits(destination, position, digits, digitCount, -exponent);
            return true;
        }

        destination[position++] = '0';
        destination[position++] = '.';
        destination.Slice(position, -integerDigits).Fill('0');
        written = WriteDigits(destination, position - integerDigits, digits, digitCount, 0);
        return true;
    }

    /// <summary>
    /// Writes the <paramref name="digitCount"/> digits of <paramref name="value"/> from <paramref name="start"/>, with a
    /// point before the last <paramref name="fractionDigits"/> of them, two digits at a time from the end, and
    /// returns the position after them.
    /// </summary>
    private static int WriteDigits(Span<char> destination, int start, ulong value, int digitCount, int fractionDigits)
    {
        bool point = fractionDigits > 0;
        int end = start + digitCount + (point ? 1 : 0);
        int position = end;
        for (; fractionDigits >= 2; fractionDigits -= 2)
        {
            position -= 2;
            value = WritePair(destination, position, value);
        }

        if (fractionDigits == 1)
        {
            ulong quotient = value / 10;
            destination[--position] = (char)('0' + (int)(value - quotient * 10));
            value = quotient;
        }

        if (point)
        {
            destination[--position] = '.';
        }

        while (position - start >= 2)
        {
            position -= 2;
            value = WritePair(destination, position, value);
        }

        if (position > start)
        {
            destination[start] = (char)('0' + (int)value);
        }

        return end;
    }

    private static ulong WritePair(Span<char> destination, int position, ulong value)
    {
        ulong quotient = value / 100;
        int pair = (int)(value - quotient * 100) * 2;
        destination[position] = TwoDigits[pair];
        destination[position + 1] = TwoDigits[pair + 1];
        return quotient;
    }

    private static int CountDigits(ulong value)
    {
        int estimate = ((64 - BitOperations.LeadingZeroCount(value | 1)) * 1233) >> 12;
        return estimate + (value >= PowersOfTen[estimate] ? 1 : 0);
    }

    /// <summary>The shortest decimal <paramref name="digits"/> * 10^<paramref name="exponent"/> that parses back to the double.</summary>
    private static void ToDecimal(ulong ieeeMantissa, int ieeeExponent, out ulong digits, out int exponent)
    {
        int e2;
        ulong m2;
        if (ieeeExponent == 0)
        {
            e2 = 1 - ExponentBias - MantissaBits - 2;
            m2 = ieeeMantissa;
        }
        else
        {
            e2 = ieeeExponent - ExponentBias - MantissaBits - 2;
            m2 = (1UL << MantissaBits) | ieeeMantissa;
        }

        // The interval of decimals that parse to this double includes its ends when the mantissa is even.
        bool acceptBounds = (m2 & 1) == 0;
        ulong mv = 4 * m2;
        uint mmShift = ieeeMantissa != 0 || ieeeExponent <= 1 ? 1u : 0u;
        ulong vr, vp, vm;
        int e10;
        bool vmIsTrailingZeros = false, vrIsTrailingZeros = false;
        if (e2 >= 0)
        {
            int q = Log10Pow2(e2) - (e2 > 3 ? 1 : 0);
            e10 = q;
            int shift = -e2 + q + Pow5InverseBits + Pow5BitLength(q) - 1;
            UInt128 multiplier = Pow5Inverse[q];
            vr = MultiplyShift(mv, multiplier, shift);
            vp = MultiplyShift(mv + 2, multiplier, shift);
            vm = MultiplyShift(mv - 1 - mmShift, multiplier, shift);
            if (q <= 21)
            {
                if (mv % 5 == 0)
                {
                    vrIsTrailingZeros = IsMultipleOfPowerOf5(mv, q);
                }
                else if (acceptBounds)
                {
                    vmIsTrailingZeros = IsMultipleOfPowerOf5(mv - 1 - mmShift, q);
                }
                else if (IsMultipleOfPowerOf5(mv + 2, q))
                {
                    vp--;
                }
            }
        }
        else
        {
            int q = Log10Pow5(-e2) - (-e2 > 1 ? 1 : 0);
            e10 = q + e2;
            int i = -e2 - q;
            int shift = q - (Pow5BitLength(i) - Pow5Bits);
            UInt128 multiplier = Pow5[i];
            vr = MultiplyShift(mv, multiplier, shift);
            vp = MultiplyShift(mv + 2, multiplier, shift);
            vm = MultiplyShift(mv - 1 - mmShift, multiplier, shift);
            if (q <= 1)
            {
                vrIsTrailingZeros = true;
                if (acceptBounds)
                {
                    vmIsTrailingZeros = mmShift == 1;
                }
                else
                {
                    vp--;
                }
            }
            else if (q < 63)
            {
                vrIsTrailingZeros = (mv & ((1UL << q) - 1)) == 0;
            }
        }

        // Remove digits while the interval still holds a shorter decimal, then round the last one kept.
        int removed = 0;
        uint lastRemovedDigit = 0;
        ulong output;
        if (vmIsTrailingZeros || vrIsTrailingZeros)
        {
            while (vp / 10 > vm / 10)
            {
                vmIsTrailingZeros &= vm % 10 == 0;
                vrIsTrailingZeros &= lastRemovedDigit == 0;
                lastRemovedDigit = (uint)(vr % 10);
                vr /= 10;
                vp /= 10;
                vm /= 10;
                removed++;
            }

            if (vmIsTrailingZeros)
            {
                while (vm % 10 == 0)
                {
                    vrIsTrailingZeros &= lastRemovedDigit == 0;
                    lastRemovedDigit = (uint)(vr % 10);
                    vr /= 10;
                    vp /= 10;
                    vm /= 10;
                    removed++;
                }
            }

            if (vrIsTrailingZeros && lastRemovedDigit == 5 && vr % 2 == 0)
            {
                // An exact half rounds to even.
                lastRemovedDigit = 4;
            }

            output = vr + ((vr == vm && (!acceptBounds || !vmIsTrailingZeros)) || lastRemovedDigit >= 5 ? 1UL : 0UL);
        }
        else
        {
            bool roundUp = false;
            if (vp / 100 > vm / 100)
            {
                roundUp = vr % 100 >= 50;
                vr /= 100;
                vp /= 100;
                vm /= 100;
                removed += 2;
            }

            while (vp / 10 > vm / 10)
            {
                roundUp = vr % 10 >= 5;
                vr /= 10;
                vp /= 10;
                vm /= 10;
                removed++;
            }

            output = vr + (vr == vm || roundUp ? 1UL : 0UL);
        }

        digits = output;
        exponent = e10 + removed;
    }

    /// <summary>The 192-bit product of <paramref name="value"/> and a 128-bit multiplier, shifted right by at least 64.</summary>
    private static ulong MultiplyShift(ulong value, UInt128 multiplier, int shift)
    {
        UInt128 low = (UInt128)value * (ulong)multiplier;
        UInt128 high = (UInt128)value * (ulong)(multiplier >> 64);
        return (ulong)(((low >> 64) + high) >> (shift - 64));
    }

    private static bool IsMultipleOfPowerOf5(ulong value, int power)
    {
        int count = 0;
        while (value % 5 == 0)
        {
            value /= 5;
            count++;
        }

        return count >= power;
    }

    /// <summary>The bit length of 5^e, which is 1 for e = 0.</summary>
    private static int Pow5BitLength(int e) => (int)(((uint)e * 1217359) >> 19) + 1;

    private static int Log10Pow2(int e) => (int)(((uint)e * 78913) >> 18);

    private static int Log10Pow5(int e) => (int)(((uint)e * 732923) >> 20);

    private static ulong[] CreatePowersOfTen()
    {
        var powers = new ulong[20];
        powers[0] = 1;
        for (int index = 1; index < powers.Length; index++) powers[index] = powers[index - 1] * 10;
        return powers;
    }

    private static UInt128[] CreatePow5Table()
    {
        var table = new UInt128[Pow5Count];
        BigInteger power = BigInteger.One;
        for (int index = 0; index < table.Length; index++)
        {
            int shift = (int)power.GetBitLength() - Pow5Bits;
            table[index] = (UInt128)(shift >= 0 ? power >> shift : power << -shift);
            power *= 5;
        }

        return table;
    }

    private static UInt128[] CreatePow5InverseTable()
    {
        var table = new UInt128[Pow5InverseCount];
        BigInteger power = BigInteger.One;
        for (int index = 0; index < table.Length; index++)
        {
            int shift = (int)power.GetBitLength() - 1 + Pow5InverseBits;
            table[index] = (UInt128)((BigInteger.One << shift) / power + 1);
            power *= 5;
        }

        return table;
    }
}
#endif
