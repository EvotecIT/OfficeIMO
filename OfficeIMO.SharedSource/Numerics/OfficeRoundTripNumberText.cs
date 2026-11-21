using System.Globalization;
#if !NETCOREAPP3_0_OR_GREATER
using System.Numerics;
using System.Text;
#endif

namespace OfficeIMO;

/// <summary>
/// The shortest text that parses back to a <see cref="double"/> or <see cref="float"/>, written the same way on
/// every target framework. .NET Core 3.0 and later write it with <c>"R"</c>. .NET Framework's <c>"R"</c> is longer
/// than needed, and its parser accepts some texts that a correct parser reads as a neighbouring value, so there the
/// digits are found with exact integer arithmetic and laid out the way .NET Core lays out <c>"R"</c>.
/// </summary>
internal static class OfficeRoundTripNumberText {
    internal static string Format(double value, CultureInfo culture) {
#if NETCOREAPP3_0_OR_GREATER
        return value.ToString("R", culture);
#else
        NumberFormatInfo format = NumberFormatInfo.GetInstance(culture);
        if (double.IsNaN(value)) return format.NaNSymbol;
        if (double.IsPositiveInfinity(value)) return format.PositiveInfinitySymbol;
        if (double.IsNegativeInfinity(value)) return format.NegativeInfinitySymbol;
        long bits = BitConverter.DoubleToInt64Bits(value);
        bool negative = bits < 0;
        if (value == 0D) return negative ? format.NegativeSign + "0" : "0";

        // A whole number below 2^53 is exact, and no shorter decimal lies within half a unit of it.
        double magnitude = Math.Abs(value);
        if (magnitude < 9007199254740992D && Math.Floor(magnitude) == magnitude) {
            return ((long)value).ToString(culture);
        }

        RyuDouble.ToDecimal((ulong)bits & 0xFFFFFFFFFFFFFUL, (int)((bits >> 52) & 0x7FF), out ulong digits, out int exponent);
        return Layout(negative, digits, exponent, 17, format);
#endif
    }

    internal static string Format(float value, CultureInfo culture) {
#if NETCOREAPP3_0_OR_GREATER
        return value.ToString("R", culture);
#else
        NumberFormatInfo format = NumberFormatInfo.GetInstance(culture);
        if (float.IsNaN(value)) return format.NaNSymbol;
        if (float.IsPositiveInfinity(value)) return format.PositiveInfinitySymbol;
        if (float.IsNegativeInfinity(value)) return format.NegativeInfinitySymbol;
        int bits = BitConverter.ToInt32(BitConverter.GetBytes(value), 0);
        bool negative = bits < 0;
        if (value == 0F) return negative ? format.NegativeSign + "0" : "0";

        // Whole numbers below 2^24, as for doubles above.
        float magnitude = Math.Abs(value);
        if (magnitude < 16777216F && Math.Floor(magnitude) == magnitude) {
            return ((int)value).ToString(culture);
        }

        int biased = (bits >> 23) & 0xFF;
        uint fraction = (uint)bits & 0x7FFFFFU;
        uint mantissa = biased == 0 ? fraction : fraction | (1U << 23);
        int binaryExponent = biased == 0 ? -149 : biased - 150;
        ExactSingle.Shortest(mantissa, binaryExponent, fraction == 0 && biased > 1, out ulong digits, out int exponent);
        return Layout(negative, digits, exponent, 9, format);
#endif
    }

#if !NETCOREAPP3_0_OR_GREATER
    // Lays out digits * 10^exponent as .NET Core's "R" does: fixed notation unless the leading digit's exponent
    // reaches the round-trip digit count or falls below -4, then d.dddE+XX with at least two exponent digits.
    private static string Layout(bool negative, ulong digits, int exponent, int roundTripDigits, NumberFormatInfo format) {
        while (digits % 10UL == 0UL) {
            digits /= 10UL;
            exponent++;
        }
        string text = digits.ToString(CultureInfo.InvariantCulture);
        int leading = exponent + text.Length - 1;
        var builder = new StringBuilder(text.Length + 8);
        if (negative) builder.Append(format.NegativeSign);
        if (leading >= roundTripDigits || leading < -4) {
            builder.Append(text[0]);
            if (text.Length > 1) builder.Append(format.NumberDecimalSeparator).Append(text, 1, text.Length - 1);
            builder.Append('E').Append(leading < 0 ? format.NegativeSign : format.PositiveSign);
            int absolute = Math.Abs(leading);
            if (absolute < 10) builder.Append('0');
            builder.Append(absolute.ToString(CultureInfo.InvariantCulture));
        } else if (leading < 0) {
            builder.Append('0').Append(format.NumberDecimalSeparator).Append('0', -leading - 1).Append(text);
        } else if (exponent >= 0) {
            builder.Append(text).Append('0', exponent);
        } else {
            int whole = leading + 1;
            builder.Append(text, 0, whole).Append(format.NumberDecimalSeparator).Append(text, whole, text.Length - whole);
        }
        return builder.ToString();
    }

    /// <summary>
    /// The shortest decimal that parses back to a double, closest to it when several qualify and rounding an exact
    /// half to even, from the Ryu algorithm (Ulf Adams, "Ryū: fast float-to-string conversion", PLDI 2018), which
    /// finds the digits with 128-bit multiplications by powers of five. OfficeIMO.CSV has the same algorithm on
    /// <c>UInt128</c> for .NET 8; here the 128-bit values are pairs of 64-bit words.
    /// </summary>
    private static class RyuDouble {
        private const int MantissaBits = 52;
        private const int ExponentBias = 1023;
        private const int Pow5Bits = 125;
        private const int Pow5InverseBits = 125;

        // 5^i normalised to its top Pow5Bits bits, and 2^j / 5^i rounded up with Pow5InverseBits significant bits,
        // as high and low 64-bit words.
        private static readonly ulong[] Pow5High, Pow5Low, Pow5InverseHigh, Pow5InverseLow;

        static RyuDouble() {
            CreateTable(326, inverse: false, out Pow5High, out Pow5Low);
            CreateTable(342, inverse: true, out Pow5InverseHigh, out Pow5InverseLow);
        }

        internal static void ToDecimal(ulong ieeeMantissa, int ieeeExponent, out ulong digits, out int exponent) {
            int e2;
            ulong m2;
            if (ieeeExponent == 0) {
                e2 = 1 - ExponentBias - MantissaBits - 2;
                m2 = ieeeMantissa;
            } else {
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
            if (e2 >= 0) {
                int q = Log10Pow2(e2) - (e2 > 3 ? 1 : 0);
                e10 = q;
                int shift = -e2 + q + Pow5InverseBits + Pow5BitLength(q) - 1;
                ulong high = Pow5InverseHigh[q], low = Pow5InverseLow[q];
                vr = MultiplyShift(mv, high, low, shift);
                vp = MultiplyShift(mv + 2, high, low, shift);
                vm = MultiplyShift(mv - 1 - mmShift, high, low, shift);
                if (q <= 21) {
                    if (mv % 5 == 0) {
                        vrIsTrailingZeros = IsMultipleOfPowerOf5(mv, q);
                    } else if (acceptBounds) {
                        vmIsTrailingZeros = IsMultipleOfPowerOf5(mv - 1 - mmShift, q);
                    } else if (IsMultipleOfPowerOf5(mv + 2, q)) {
                        vp--;
                    }
                }
            } else {
                int q = Log10Pow5(-e2) - (-e2 > 1 ? 1 : 0);
                e10 = q + e2;
                int i = -e2 - q;
                int shift = q - (Pow5BitLength(i) - Pow5Bits);
                ulong high = Pow5High[i], low = Pow5Low[i];
                vr = MultiplyShift(mv, high, low, shift);
                vp = MultiplyShift(mv + 2, high, low, shift);
                vm = MultiplyShift(mv - 1 - mmShift, high, low, shift);
                if (q <= 1) {
                    vrIsTrailingZeros = true;
                    if (acceptBounds) {
                        vmIsTrailingZeros = mmShift == 1;
                    } else {
                        vp--;
                    }
                } else if (q < 63) {
                    vrIsTrailingZeros = (mv & ((1UL << q) - 1)) == 0;
                }
            }

            // Remove digits while the interval still holds a shorter decimal, then round the last one kept.
            int removed = 0;
            uint lastRemovedDigit = 0;
            ulong output;
            if (vmIsTrailingZeros || vrIsTrailingZeros) {
                while (vp / 10 > vm / 10) {
                    vmIsTrailingZeros &= vm % 10 == 0;
                    vrIsTrailingZeros &= lastRemovedDigit == 0;
                    lastRemovedDigit = (uint)(vr % 10);
                    vr /= 10;
                    vp /= 10;
                    vm /= 10;
                    removed++;
                }
                if (vmIsTrailingZeros) {
                    while (vm % 10 == 0) {
                        vrIsTrailingZeros &= lastRemovedDigit == 0;
                        lastRemovedDigit = (uint)(vr % 10);
                        vr /= 10;
                        vp /= 10;
                        vm /= 10;
                        removed++;
                    }
                }
                // An exact half rounds to even.
                if (vrIsTrailingZeros && lastRemovedDigit == 5 && vr % 2 == 0) lastRemovedDigit = 4;
                output = vr + ((vr == vm && (!acceptBounds || !vmIsTrailingZeros)) || lastRemovedDigit >= 5 ? 1UL : 0UL);
            } else {
                bool roundUp = false;
                if (vp / 100 > vm / 100) {
                    roundUp = vr % 100 >= 50;
                    vr /= 100;
                    vp /= 100;
                    vm /= 100;
                    removed += 2;
                }
                while (vp / 10 > vm / 10) {
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

        // The 192-bit product of value and the 128-bit multiplier high:low, shifted right by shift (at least 64).
        private static ulong MultiplyShift(ulong value, ulong high, ulong low, int shift) {
            ulong lowProductHigh = MultiplyHigh(value, low, out _);
            ulong highProductHigh = MultiplyHigh(value, high, out ulong highProductLow);
            ulong sumLow = highProductLow + lowProductHigh;
            ulong sumHigh = highProductHigh + (sumLow < lowProductHigh ? 1UL : 0UL);
            int bits = shift - 64;
            if (bits == 0) return sumLow;
            return bits < 64 ? (sumHigh << (64 - bits)) | (sumLow >> bits) : sumHigh >> (bits - 64);
        }

        // The high 64 bits of a * b, with the low 64 bits in low.
        private static ulong MultiplyHigh(ulong a, ulong b, out ulong low) {
            ulong aLow = (uint)a, aHigh = a >> 32, bLow = (uint)b, bHigh = b >> 32;
            ulong lowLow = aLow * bLow;
            ulong highLow = aHigh * bLow;
            ulong lowHigh = aLow * bHigh;
            ulong highHigh = aHigh * bHigh;
            ulong middle = (lowLow >> 32) + (uint)highLow + (uint)lowHigh;
            low = (middle << 32) | (uint)lowLow;
            return highHigh + (highLow >> 32) + (lowHigh >> 32) + (middle >> 32);
        }

        private static bool IsMultipleOfPowerOf5(ulong value, int power) {
            int count = 0;
            while (value % 5 == 0) {
                value /= 5;
                count++;
            }
            return count >= power;
        }

        // The bit length of 5^e, which is 1 for e = 0.
        private static int Pow5BitLength(int e) => (int)(((uint)e * 1217359) >> 19) + 1;

        private static int Log10Pow2(int e) => (int)(((uint)e * 78913) >> 18);

        private static int Log10Pow5(int e) => (int)(((uint)e * 732923) >> 20);

        private static void CreateTable(int count, bool inverse, out ulong[] high, out ulong[] low) {
            high = new ulong[count];
            low = new ulong[count];
            BigInteger mask = (BigInteger.One << 64) - 1;
            BigInteger power = BigInteger.One;
            for (int index = 0; index < count; index++) {
                int length = BitLength(power);
                BigInteger entry = inverse
                    ? (BigInteger.One << (length - 1 + Pow5InverseBits)) / power + 1
                    : length >= Pow5Bits ? power >> (length - Pow5Bits) : power << (Pow5Bits - length);
                high[index] = (ulong)(entry >> 64);
                low[index] = (ulong)(entry & mask);
                power *= 5;
            }
        }

        private static int BitLength(BigInteger value) {
            byte[] bytes = value.ToByteArray();
            int top = bytes.Length - 1;
            while (top > 0 && bytes[top] == 0) top--;
            int length = top * 8;
            for (int b = bytes[top]; b != 0; b >>= 1) length++;
            return length;
        }
    }

    /// <summary>
    /// The shortest decimal that parses back to a float, found with exact integer arithmetic: the decimal nearest
    /// the float at each length, rounding halves to even, is checked against the interval of reals that round to it.
    /// Floats are rare enough here that this simpler search is used instead of a float variant of Ryu.
    /// </summary>
    private static class ExactSingle {
        private static readonly BigInteger[] PowersOfTen = CreatePowersOfTen(100);

        internal static void Shortest(uint mantissa, int binaryExponent, bool narrowLowerGap, out ulong digits, out int exponent) {
            int leading = (int)Math.Floor(Math.Log10(mantissa) + binaryExponent * 0.30102999566398120D);
            // Nine digits always identify a float; the closest decimal at a length reads back only if the closest
            // at every greater length does, so the shortest length is found by bisection.
            int low = 1, high = 9;
            digits = Closest(mantissa, binaryExponent, 9, ref leading, out exponent);
            while (low < high) {
                int middle = (low + high) / 2;
                ulong candidate = Closest(mantissa, binaryExponent, middle, ref leading, out int candidateExponent);
                if (Contains(mantissa, binaryExponent, narrowLowerGap, candidate, candidateExponent)) {
                    high = middle;
                    digits = candidate;
                    exponent = candidateExponent;
                } else {
                    low = middle + 1;
                }
            }
        }

        private static ulong Closest(uint mantissa, int binaryExponent, int precision, ref int leading, out int exponent) {
            while (true) {
                exponent = leading - precision + 1;
                BigInteger numerator = new BigInteger(mantissa);
                BigInteger denominator = BigInteger.One;
                if (binaryExponent > 0) numerator <<= binaryExponent;
                else denominator <<= -binaryExponent;
                if (exponent < 0) numerator *= PowersOfTen[-exponent];
                else denominator *= PowersOfTen[exponent];
                BigInteger rounded = BigInteger.DivRem(numerator, denominator, out BigInteger remainder);
                int half = (remainder << 1).CompareTo(denominator);
                if (half > 0 || (half == 0 && !rounded.IsEven)) rounded += BigInteger.One;
                // A low estimate, or digits that carry as 9.99 does to 10.0, move the leading digit up one place.
                if (rounded >= PowersOfTen[precision]) {
                    leading++;
                    continue;
                }
                if (rounded < PowersOfTen[precision - 1]) {
                    leading--;
                    continue;
                }
                return (ulong)rounded;
            }
        }

        // Whether digits * 10^exponent lies within half a unit in the last place of the float (a quarter below a
        // power of two), ends included when the mantissa is even, since a halfway text rounds to the even neighbour.
        private static bool Contains(uint mantissa, int binaryExponent, bool narrowLowerGap, ulong digits, int exponent) {
            BigInteger left = new BigInteger(digits);
            BigInteger scale = BigInteger.One;
            if (exponent > 0) left *= PowersOfTen[exponent];
            else if (exponent < 0) scale *= PowersOfTen[-exponent];
            int shift = binaryExponent - 2;
            if (shift < 0) left <<= -shift;
            else scale <<= shift;
            BigInteger quadruple = new BigInteger(mantissa) << 2;
            int toLower = left.CompareTo((quadruple - (narrowLowerGap ? 1 : 2)) * scale);
            int toUpper = left.CompareTo((quadruple + 2) * scale);
            return (mantissa & 1U) == 0U ? toLower >= 0 && toUpper <= 0 : toLower > 0 && toUpper < 0;
        }

        private static BigInteger[] CreatePowersOfTen(int count) {
            var powers = new BigInteger[count];
            powers[0] = BigInteger.One;
            for (int index = 1; index < count; index++) powers[index] = powers[index - 1] * 10;
            return powers;
        }
    }
#endif
}
