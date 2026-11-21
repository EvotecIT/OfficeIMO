using System.Globalization;
using System.Numerics;

namespace OfficeIMO;

/// <summary>
/// Correctly rounded number parsing on every framework. .NET Core 3.0 and later parse correctly, so there these
/// delegate to the runtime. .NET Framework's parser can return a neighbouring double for long digit strings, reports
/// an overflow where .NET Core returns an infinity, loses the sign of negative zero and accepts infinity and NaN only
/// in their exact spelling. Elsewhere plain decimal text is parsed here: a single exact operation when the digits and
/// the power of ten are small enough, the Eisel-Lemire algorithm (Daniel Lemire, "Number Parsing at a Gigabyte per
/// Second", 2021) for up to 19 significant digits, and exact big-integer arithmetic when neither settles the result.
/// Other forms go to .NET Framework's parser: currency symbols, parentheses, and text read with currency allowed that
/// contains a currency separator differing from the culture's number separator.
/// </summary>
internal static partial class OfficeRoundTripNumberText {
    /// <summary>Parses a double as .NET Core's <see cref="double.TryParse(string, NumberStyles, IFormatProvider, out double)"/> does.</summary>
    internal static bool TryParse(string? text, NumberStyles style, IFormatProvider? provider, out double value) {
#if NETCOREAPP3_0_OR_GREATER
        return double.TryParse(text, style, provider, out value);
#else
        value = 0D;
        if (text == null) return false;
        if ((style & NumberStyles.AllowHexSpecifier) != 0) return double.TryParse(text, style, provider, out value);
        NumberFormatInfo format = NumberFormatInfo.GetInstance(provider);
        if (DecimalText.TryScan(new TextSource(text), style, format, out DecimalText scanned)) {
            value = scanned.ToDouble(new TextSource(text));
            return true;
        }
        if (TryParseSymbol(text, style, format, out value)) return true;
        return double.TryParse(text, style, format, out value);
#endif
    }

    /// <summary>Parses a float as .NET Core's <see cref="float.TryParse(string, NumberStyles, IFormatProvider, out float)"/> does.</summary>
    internal static bool TryParse(string? text, NumberStyles style, IFormatProvider? provider, out float value) {
#if NETCOREAPP3_0_OR_GREATER
        return float.TryParse(text, style, provider, out value);
#else
        value = 0F;
        if (text == null) return false;
        if ((style & NumberStyles.AllowHexSpecifier) != 0) return float.TryParse(text, style, provider, out value);
        NumberFormatInfo format = NumberFormatInfo.GetInstance(provider);
        if (DecimalText.TryScan(new TextSource(text), style, format, out DecimalText scanned)) {
            value = scanned.ToSingle(new TextSource(text));
            return true;
        }
        if (TryParseSymbol(text, style, format, out double symbol)) {
            value = (float)symbol;
            return true;
        }
        return float.TryParse(text, style, format, out value);
#endif
    }

    /// <summary>
    /// Parses a double like <see cref="TryParse(string, NumberStyles, IFormatProvider, out double)"/>, but treats
    /// text that reads as infinity or NaN (an overflow such as <c>1e400</c>, or a spelled-out <c>NaN</c>) as not a
    /// number, for values that a file format can store only as finite numbers.
    /// </summary>
    internal static bool TryParseFinite(string? text, NumberStyles style, IFormatProvider? provider, out double value) {
        if (TryParse(text, style, provider, out value) && !double.IsNaN(value) && !double.IsInfinity(value)) return true;
        value = 0D;
        return false;
    }

#if NETCOREAPP3_0_OR_GREATER
    // The runtime parser is correct here; these let span call sites share the same owner.
    internal static bool TryParse(ReadOnlySpan<char> text, NumberStyles style, IFormatProvider? provider, out double value) =>
        double.TryParse(text, style, provider, out value);

    internal static bool TryParse(ReadOnlySpan<char> text, NumberStyles style, IFormatProvider? provider, out float value) =>
        float.TryParse(text, style, provider, out value);

    internal static double ParseDouble(ReadOnlySpan<char> text, NumberStyles style, IFormatProvider? provider) =>
        double.Parse(text, style, provider);

    internal static float ParseSingle(ReadOnlySpan<char> text, NumberStyles style, IFormatProvider? provider) =>
        float.Parse(text, style, provider);
#endif

    /// <summary>Parses a double as .NET Core's <see cref="double.Parse(string, NumberStyles, IFormatProvider)"/> does.</summary>
    internal static double ParseDouble(string? text, NumberStyles style, IFormatProvider? provider) {
#if NETCOREAPP3_0_OR_GREATER
        return double.Parse(text!, style, provider);
#else
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (TryParse(text, style, provider, out double value)) return value;
        throw new FormatException("The input string '" + text + "' was not in a correct format.");
#endif
    }

    /// <summary>Parses a float as .NET Core's <see cref="float.Parse(string, NumberStyles, IFormatProvider)"/> does.</summary>
    internal static float ParseSingle(string? text, NumberStyles style, IFormatProvider? provider) {
#if NETCOREAPP3_0_OR_GREATER
        return float.Parse(text!, style, provider);
#else
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (TryParse(text, style, provider, out float value)) return value;
        throw new FormatException("The input string '" + text + "' was not in a correct format.");
#endif
    }

    /// <summary>
    /// Parses ASCII text such as UTF-8 cell values: an optional sign, digits with an optional '.' fraction and an
    /// optional exponent, nothing else, read as .NET Core's parser reads it with the invariant culture.
    /// </summary>
    internal static bool TryParseAscii(byte[] buffer, int offset, int count, out double value) {
        value = 0D;
        var source = new TextSource(buffer, offset, count);
        if (!DecimalText.TryScan(source, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
                NumberFormatInfo.InvariantInfo, out DecimalText scanned)) return false;
        value = scanned.ToDouble(source);
        return true;
    }

    /// <summary>
    /// Converts a value as <see cref="Convert.ToDouble(object, IFormatProvider)"/> does, parsing text as .NET Core does.
    /// </summary>
    internal static double ToDouble(object? value, IFormatProvider? provider) =>
        value is string text ? ParseDouble(text, NumberStyles.Float | NumberStyles.AllowThousands, provider) : Convert.ToDouble(value, provider);

    /// <summary>
    /// Converts a value as <see cref="Convert.ToSingle(object, IFormatProvider)"/> does, parsing text as .NET Core does.
    /// </summary>
    internal static float ToSingle(object? value, IFormatProvider? provider) =>
        value is string text ? ParseSingle(text, NumberStyles.Float | NumberStyles.AllowThousands, provider) : Convert.ToSingle(value, provider);

    private const NumberStyles XmlNumberStyle = NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint |
        NumberStyles.AllowExponent | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite;

    /// <summary>Reads an <c>xsd:double</c> as .NET Core's <see cref="System.Xml.XmlConvert.ToDouble(string)"/> does.</summary>
    internal static double ParseXmlDouble(string? text) {
#if NETCOREAPP3_0_OR_GREATER
        return System.Xml.XmlConvert.ToDouble(text!);
#else
        if (text == null) throw new ArgumentNullException(nameof(text));
        string trimmed = text.Trim(XmlWhitespace);
        if (trimmed == "-INF") return double.NegativeInfinity;
        if (trimmed == "INF") return double.PositiveInfinity;
        return ParseDouble(trimmed, XmlNumberStyle, NumberFormatInfo.InvariantInfo);
#endif
    }

    /// <summary>Reads an <c>xsd:float</c> as .NET Core's <see cref="System.Xml.XmlConvert.ToSingle(string)"/> does.</summary>
    internal static float ParseXmlSingle(string? text) {
#if NETCOREAPP3_0_OR_GREATER
        return System.Xml.XmlConvert.ToSingle(text!);
#else
        if (text == null) throw new ArgumentNullException(nameof(text));
        string trimmed = text.Trim(XmlWhitespace);
        if (trimmed == "-INF") return float.NegativeInfinity;
        if (trimmed == "INF") return float.PositiveInfinity;
        return ParseSingle(trimmed, XmlNumberStyle, NumberFormatInfo.InvariantInfo);
#endif
    }

#if !NETCOREAPP3_0_OR_GREATER
    private static readonly char[] XmlWhitespace = { ' ', '\t', '\n', '\r' };

    // .NET Core accepts its infinity and NaN symbols in any letter case, optionally signed.
    private static bool TryParseSymbol(string text, NumberStyles style, NumberFormatInfo format, out double value) {
        value = 0D;
        int start = 0, end = text.Length;
        if ((style & NumberStyles.AllowLeadingWhite) != 0) while (start < end && IsWhite(text[start])) start++;
        if ((style & NumberStyles.AllowTrailingWhite) != 0) while (end > start && IsWhite(text[end - 1])) end--;
        if (start == end) return false;
        string body = text.Substring(start, end - start);
        if (Equal(body, format.PositiveInfinitySymbol) || Equal(body, format.PositiveSign + format.PositiveInfinitySymbol)) {
            value = double.PositiveInfinity;
            return true;
        }
        if (Equal(body, format.NegativeInfinitySymbol) || Equal(body, format.NegativeSign + format.PositiveInfinitySymbol)) {
            value = double.NegativeInfinity;
            return true;
        }
        if (Equal(body, format.NaNSymbol) || Equal(body, format.PositiveSign + format.NaNSymbol) || Equal(body, format.NegativeSign + format.NaNSymbol)) {
            value = double.NaN;
            return true;
        }
        return false;
    }

    private static bool Equal(string text, string symbol) =>
        symbol.Length > 0 && string.Equals(text, symbol, StringComparison.OrdinalIgnoreCase);

#endif

    private static bool IsWhite(char c) => c == ' ' || (c >= '\t' && c <= '\r');

    /// <summary>Characters of a string, or ASCII bytes read as characters.</summary>
    private readonly struct TextSource {
        private readonly string? _text;
        private readonly byte[]? _bytes;
        private readonly int _offset;

        internal TextSource(string text) {
            _text = text;
            _bytes = null;
            _offset = 0;
            Length = text.Length;
        }

        internal TextSource(byte[] bytes, int offset, int count) {
            _text = null;
            _bytes = bytes;
            _offset = offset;
            Length = count;
        }

        internal int Length { get; }

        internal char this[int index] => _text != null ? _text[index] : (char)_bytes![_offset + index];

        internal bool StartsWith(int index, string value) {
            if (value.Length == 0 || index + value.Length > Length) return false;
            for (int i = 0; i < value.Length; i++) {
                if (this[index + i] != value[i]) return false;
            }
            return true;
        }

        internal bool Contains(string value) {
            for (int index = 0; index + value.Length <= Length; index++) {
                if (StartsWith(index, value)) return true;
            }
            return false;
        }
    }

    /// <summary>A plain decimal number found in a text: sign, up to 19 leading significant digits and a power of ten.</summary>
    private struct DecimalText {
        private const int MaximumExponent = 100000;
        private const int MaximumDigits = 19;

        private bool _negative;
        private ulong _significand;
        private int _exponent;
        private bool _truncated;
        // Where the digits and the separator are, for the exact path.
        private int _digitsStart, _digitsEnd, _pointAt, _pointLength, _groupLength;
        private int _writtenExponent;

        internal static bool TryScan(TextSource text, NumberStyles style, NumberFormatInfo format, out DecimalText result) {
            result = default;
            // With currency allowed, the runtime also takes the currency separators, ahead of the number ones. Text that
            // contains a currency separator differing from its number counterpart is left to the runtime parser.
            if ((style & NumberStyles.AllowCurrencySymbol) != 0
                && ((format.CurrencyDecimalSeparator != format.NumberDecimalSeparator && text.Contains(format.CurrencyDecimalSeparator))
                    || (format.CurrencyGroupSeparator != format.NumberGroupSeparator && text.Contains(format.CurrencyGroupSeparator)))) {
                return false;
            }
            int index = 0, length = text.Length;
            if ((style & NumberStyles.AllowLeadingWhite) != 0) while (index < length && IsWhite(text[index])) index++;
            if ((style & NumberStyles.AllowLeadingSign) != 0) {
                if (text.StartsWith(index, format.NegativeSign)) {
                    result._negative = true;
                    index += format.NegativeSign.Length;
                } else if (text.StartsWith(index, format.PositiveSign)) {
                    index += format.PositiveSign.Length;
                }
            }

            string point = format.NumberDecimalSeparator, group = format.NumberGroupSeparator;
            bool allowPoint = (style & NumberStyles.AllowDecimalPoint) != 0 && point.Length > 0;
            bool allowGroup = (style & NumberStyles.AllowThousands) != 0 && group.Length > 0 && group != point;
            result._digitsStart = index;
            result._pointAt = -1;
            result._pointLength = point.Length;
            result._groupLength = group.Length;
            bool anyDigit = false, inFraction = false;
            int kept = 0, scale = 0;
            ulong significand = 0;
            while (index < length) {
                char c = text[index];
                if (c >= '0' && c <= '9') {
                    anyDigit = true;
                    if (significand == 0 && c == '0') {
                        if (inFraction) scale--;
                    } else if (kept < MaximumDigits) {
                        significand = significand * 10 + (ulong)(c - '0');
                        kept++;
                        if (inFraction) scale--;
                    } else {
                        if (c != '0') result._truncated = true;
                        if (!inFraction) scale++;
                    }
                    index++;
                } else if (!inFraction && allowPoint && text.StartsWith(index, point)) {
                    inFraction = true;
                    result._pointAt = index;
                    index += point.Length;
                } else if (!inFraction && allowGroup && anyDigit && text.StartsWith(index, group)) {
                    index += group.Length;
                } else {
                    break;
                }
            }
            if (!anyDigit) return false;
            result._digitsEnd = index;

            int written = 0;
            if (index < length && (text[index] == 'e' || text[index] == 'E') && (style & NumberStyles.AllowExponent) != 0) {
                int at = index + 1;
                bool negativeExponent = false;
                if (text.StartsWith(at, format.NegativeSign)) {
                    negativeExponent = true;
                    at += format.NegativeSign.Length;
                } else if (text.StartsWith(at, format.PositiveSign)) {
                    at += format.PositiveSign.Length;
                }
                int digits = 0;
                while (at < length && text[at] >= '0' && text[at] <= '9') {
                    if (written < MaximumExponent) written = written * 10 + (text[at] - '0');
                    digits++;
                    at++;
                }
                // An 'e' without digits is left to the runtime parser.
                if (digits == 0) return false;
                if (negativeExponent) written = -written;
                index = at;
            }
            if ((style & NumberStyles.AllowTrailingWhite) != 0) while (index < length && IsWhite(text[index])) index++;
            if (index != length) return false;

            result._significand = significand;
            result._writtenExponent = written;
            result._exponent = significand == 0 ? 0 : scale + written;
            return true;
        }

        internal double ToDouble(TextSource text) {
            double magnitude = Magnitude(text);
            return _negative ? -magnitude : magnitude;
        }

        internal float ToSingle(TextSource text) {
            double magnitude = Magnitude(text);
            float single = (float)magnitude;
            if (IsSingleMidpoint(magnitude, single)) {
                ExactBinary.Round(Digits(text, out int exponent), exponent, 24, -149, 104, out ulong mantissa, out int binaryExponent, out bool infinite);
                single = infinite ? float.PositiveInfinity : SingleFromParts(mantissa, binaryExponent);
            }
            return _negative ? -single : single;
        }

        private double Magnitude(TextSource text) {
            if (_significand == 0) return 0D;
            ulong w = _significand;
            int q = _exponent;
            if (!_truncated && w <= (1UL << 53)) {
                // Both operands are exact, so one correctly rounded operation gives the correctly rounded result.
                if (q >= 0 && q <= 22) return w * PowersOfTen[q];
                if (q < 0 && q >= -22) return w / PowersOfTen[-q];
                if (q > 22 && q <= 22 + 15) {
                    double scaled = w * PowersOfTen[q - 22];
                    if (scaled <= 9007199254740992D) return scaled * PowersOfTen[22];
                }
            }
            if (EiselLemire.TryCompute(q, w, out double value)) {
                if (!_truncated) return value;
                // The dropped digits lie between w and w + 1; when both round alike, so does the exact value.
                if (EiselLemire.TryCompute(q, w + 1, out double upper) && upper.Equals(value)) return value;
            }
            ExactBinary.Round(Digits(text, out int exponent), exponent, 53, -1074, 971, out ulong mantissa, out int binaryExponent, out bool infinite);
            return infinite ? double.PositiveInfinity : DoubleFromParts(mantissa, binaryExponent);
        }

        // Every significant digit as an integer, with the power of ten that scales it; digits past 800 only matter
        // when non-zero, so they are folded into one final digit 1.
        private BigInteger Digits(TextSource text, out int exponent) {
            var builder = new System.Text.StringBuilder(Math.Min(_digitsEnd - _digitsStart, 802));
            int scale = 0;
            bool inFraction = false, sticky = false;
            for (int index = _digitsStart; index < _digitsEnd;) {
                if (index == _pointAt) {
                    inFraction = true;
                    index += _pointLength;
                    continue;
                }
                char c = text[index];
                if (c < '0' || c > '9') {
                    index += _groupLength;
                    continue;
                }
                if (builder.Length == 0 && c == '0') {
                    if (inFraction) scale--;
                } else if (builder.Length < 800) {
                    builder.Append(c);
                    if (inFraction) scale--;
                } else {
                    if (c != '0') sticky = true;
                    if (!inFraction) scale++;
                }
                index++;
            }
            if (sticky) {
                builder.Append('1');
                scale--;
            }
            exponent = scale + _writtenExponent;
            return builder.Length == 0 ? BigInteger.Zero : BigInteger.Parse(builder.ToString(), CultureInfo.InvariantCulture);
        }

        // (float)double rounds once more; that only differs from rounding the exact decimal when the double sits
        // exactly halfway between two floats.
        private static bool IsSingleMidpoint(double magnitude, float single) {
            if (double.IsInfinity(magnitude) || (double)single == magnitude) return false;
            float lower, upper;
            if ((double)single > magnitude) {
                upper = single;
                lower = NextSingle(single, down: true);
            } else {
                lower = single;
                upper = NextSingle(single, down: false);
            }
            double middle = float.IsInfinity(upper) ? 3.4028235677973366E+38 : ((double)lower + upper) / 2D;
            return middle == magnitude;
        }

        private static float NextSingle(float value, bool down) {
            int bits = BitConverter.ToInt32(BitConverter.GetBytes(value), 0) + (down ? -1 : 1);
            return BitConverter.ToSingle(BitConverter.GetBytes(bits), 0);
        }
    }

    private static readonly double[] PowersOfTen = {
        1E0, 1E1, 1E2, 1E3, 1E4, 1E5, 1E6, 1E7, 1E8, 1E9, 1E10, 1E11, 1E12, 1E13, 1E14, 1E15, 1E16, 1E17, 1E18, 1E19,
        1E20, 1E21, 1E22
    };

    private static double DoubleFromParts(ulong mantissa, int binaryExponent) {
        // mantissa * 2^binaryExponent, with mantissa below 2^53; below 2^52 only at the subnormal exponent.
        long bits = mantissa < (1UL << 52) ? (long)mantissa : ((long)(binaryExponent + 1075) << 52) | (long)(mantissa & ((1UL << 52) - 1));
        return BitConverter.Int64BitsToDouble(bits);
    }

    private static float SingleFromParts(ulong mantissa, int binaryExponent) {
        int bits = mantissa < (1UL << 23) ? (int)mantissa : ((binaryExponent + 150) << 23) | (int)(mantissa & ((1UL << 23) - 1));
        return BitConverter.ToSingle(BitConverter.GetBytes(bits), 0);
    }

    /// <summary>Rounds digits * 10^exponent to the nearest binary value with exact integer arithmetic.</summary>
    private static class ExactBinary {
        internal static void Round(BigInteger digits, int exponent, int precision, int minimumExponent, int maximumExponent,
            out ulong mantissa, out int binaryExponent, out bool infinite) {
            mantissa = 0;
            binaryExponent = minimumExponent;
            infinite = false;
            if (digits.IsZero) return;
            // Far outside the range the result is zero or infinite; this bounds the size of the integers below.
            int digitCount = digits.ToString(CultureInfo.InvariantCulture).Length;
            if (exponent + digitCount < -400) return;
            if (exponent + digitCount > 400) {
                infinite = true;
                return;
            }
            BigInteger numerator = exponent >= 0 ? digits * BigInteger.Pow(10, exponent) : digits;
            BigInteger denominator = exponent >= 0 ? BigInteger.One : BigInteger.Pow(10, -exponent);
            int e2 = BitLength(numerator) - BitLength(denominator) - precision;
            BigInteger quotient, remainder, divisor;
            while (true) {
                if (e2 < minimumExponent) e2 = minimumExponent;
                BigInteger n = e2 >= 0 ? numerator : numerator << -e2;
                divisor = e2 >= 0 ? denominator << e2 : denominator;
                quotient = BigInteger.DivRem(n, divisor, out remainder);
                if (quotient >= (BigInteger.One << precision)) {
                    e2++;
                    continue;
                }
                if (quotient < (BigInteger.One << (precision - 1)) && e2 > minimumExponent) {
                    e2--;
                    continue;
                }
                break;
            }
            int half = (remainder << 1).CompareTo(divisor);
            if (half > 0 || (half == 0 && !quotient.IsEven)) quotient += BigInteger.One;
            if (quotient == (BigInteger.One << precision)) {
                quotient >>= 1;
                e2++;
            }
            if (e2 > maximumExponent) {
                infinite = true;
                return;
            }
            mantissa = (ulong)quotient;
            binaryExponent = e2;
        }

        private static int BitLength(BigInteger value) {
            byte[] bytes = value.ToByteArray();
            int top = bytes.Length - 1;
            while (top > 0 && bytes[top] == 0) top--;
            int bits = top * 8;
            for (int b = bytes[top]; b != 0; b >>= 1) bits++;
            return bits;
        }
    }

    /// <summary>
    /// The Eisel-Lemire algorithm for doubles: the significand times a 128-bit approximation of the power of five
    /// gives the correctly rounded result for every significand below 10^19 (Mushtak and Lemire, "Fast number parsing
    /// without fallback", 2023), except that it does not decide the rare cases it reports.
    /// </summary>
    private static class EiselLemire {
        private const int SmallestPowerOfTen = -342;
        private const int LargestPowerOfTen = 308;
        private static readonly ulong[] PowersOfFive = CreatePowersOfFive();

        internal static bool TryCompute(int q, ulong w, out double value) {
            value = 0D;
            if (w == 0 || q < SmallestPowerOfTen) return true;
            if (q > LargestPowerOfTen) {
                value = double.PositiveInfinity;
                return true;
            }
            int leadingZeros = LeadingZeros(w);
            w <<= leadingZeros;
            int index = 2 * (q - SmallestPowerOfTen);
            ulong high = MultiplyHigh(w, PowersOfFive[index], out ulong low);
            if ((high & 0x1FF) == 0x1FF) {
                ulong secondHigh = MultiplyHigh(w, PowersOfFive[index + 1], out _);
                low += secondHigh;
                if (secondHigh > low) high++;
                if (low == ulong.MaxValue && (q < -27 || q > 55)) return false;
            }
            int upperBit = (int)(high >> 63);
            int shift = upperBit + 64 - 52 - 3;
            ulong mantissa = high >> shift;
            int power2 = ((217706 * q) >> 16) + 63 + upperBit - leadingZeros + 1023;
            if (power2 <= 0) {
                if (-power2 + 1 >= 64) return true;
                mantissa >>= -power2 + 1;
                mantissa += mantissa & 1;
                mantissa >>= 1;
                power2 = mantissa < (1UL << 52) ? 0 : 1;
                value = BitConverter.Int64BitsToDouble(((long)power2 << 52) | (long)(mantissa & ((1UL << 52) - 1)));
                return true;
            }
            // A product that is exactly halfway rounds to even.
            if (low <= 1 && q >= -4 && q <= 23 && (mantissa & 3) == 1 && (mantissa << shift) == high) mantissa &= ~1UL;
            mantissa += mantissa & 1;
            mantissa >>= 1;
            if (mantissa >= (2UL << 52)) {
                mantissa = 1UL << 52;
                power2++;
            }
            if (power2 >= 0x7FF) {
                value = double.PositiveInfinity;
                return true;
            }
            value = BitConverter.Int64BitsToDouble(((long)power2 << 52) | (long)(mantissa & ((1UL << 52) - 1)));
            return true;
        }

        private static int LeadingZeros(ulong value) {
            int count = 0;
            if ((value >> 32) == 0) { count += 32; value <<= 32; }
            if ((value >> 48) == 0) { count += 16; value <<= 16; }
            if ((value >> 56) == 0) { count += 8; value <<= 8; }
            if ((value >> 60) == 0) { count += 4; value <<= 4; }
            if ((value >> 62) == 0) { count += 2; value <<= 2; }
            if ((value >> 63) == 0) count++;
            return count;
        }

        private static ulong MultiplyHigh(ulong a, ulong b, out ulong low) {
            ulong aLow = (uint)a, aHigh = a >> 32, bLow = (uint)b, bHigh = b >> 32;
            ulong ll = aLow * bLow, lh = aLow * bHigh, hl = aHigh * bLow, hh = aHigh * bHigh;
            ulong middle = (ll >> 32) + (uint)lh + (uint)hl;
            low = (middle << 32) | (uint)ll;
            return hh + (lh >> 32) + (hl >> 32) + (middle >> 32);
        }

        // For each q from -342 to 308, 5^q as a 128-bit value with its top bit set: truncated for q >= 0, and for
        // q < 0 the reciprocal rounded up, as the algorithm's published table is built.
        private static ulong[] CreatePowersOfFive() {
            var table = new ulong[2 * (LargestPowerOfTen - SmallestPowerOfTen + 1)];
            BigInteger two128 = BigInteger.One << 128, mask = (BigInteger.One << 64) - 1;
            for (int q = SmallestPowerOfTen; q <= LargestPowerOfTen; q++) {
                BigInteger value;
                if (q >= 0) {
                    value = BigInteger.Pow(5, q);
                    int bits = BitLength(value);
                    value = bits < 128 ? value << (128 - bits) : value >> (bits - 128);
                } else {
                    BigInteger power = BigInteger.Pow(5, -q);
                    int z = BitLength(power);
                    if ((BigInteger.One << (z - 1)) == power) z--;
                    if (q >= -27) {
                        value = (BigInteger.One << (z + 127)) / power + 1;
                    } else {
                        value = (BigInteger.One << (2 * z + 128)) / power + 1;
                        while (value >= two128) value >>= 1;
                    }
                }
                int index = 2 * (q - SmallestPowerOfTen);
                table[index] = (ulong)(value >> 64);
                table[index + 1] = (ulong)(value & mask);
            }
            return table;
        }

        private static int BitLength(BigInteger value) {
            int bits = 0;
            while (value > ulong.MaxValue) {
                value >>= 64;
                bits += 64;
            }
            for (ulong rest = (ulong)value; rest != 0; rest >>= 1) bits++;
            return bits;
        }
    }
}
