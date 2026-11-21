using System.Buffers.Text;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("0")]
    [InlineData("-0")]
    [InlineData("+0")]
    [InlineData("-0.0")]
    [InlineData("-0.00")]
    [InlineData("+0.00")]
    [InlineData("-.25")]
    [InlineData("+.25")]
    [InlineData("99999999999999.9")]
    [InlineData("9999999999999.99")]
    [InlineData("100000000000000.0")]
    [InlineData("10000000000000.00")]
    [InlineData("42")]
    [InlineData("-42")]
    [InlineData(".5")]
    [InlineData("1.")]
    [InlineData("000000000000000123")]
    [InlineData("0.0000000000000000000001")]
    [InlineData("123456789012345")]
    [InlineData("1.23456789012345")]
    [InlineData("999999999999999")]
    [InlineData("0.999999999999999")]
    [InlineData("0.100000000000005")]
    [InlineData("-0.0000000000000000000001")]
    [InlineData("000000000000000000000000000000001.25")]
    public void Utf8NumberParser_CommonFixedPointMatchesGeneralParser(string text) =>
        AssertParserMatchesGeneralParser(text);

    [Fact]
    public void Utf8NumberParser_AllSupportedScalesMatchGeneralParserBitForBit() {
        var random = new Random(0x51CA1E);
        for (int scale = 1; scale <= 22; scale++) {
            for (int sample = 0; sample < 1_000; sample++) {
                int digitCount = random.Next(1, 16);
                long lowerBound = digitCount == 1
                    ? 1
                    : (long)Math.Pow(10, digitCount - 1);
                long upperBound = digitCount == 15
                    ? 1_000_000_000_000_000L
                    : (long)Math.Pow(10, digitCount);
                long mantissa = NextInt64(random, lowerBound, upperBound);
                string digits = mantissa.ToString(CultureInfo.InvariantCulture);
                string unsigned = scale < digits.Length
                    ? digits.Insert(digits.Length - scale, ".")
                    : "0." + new string('0', scale - digits.Length) + digits;
                string text = (sample & 1) == 0 ? unsigned : "-" + unsigned;
                AssertParserMatchesGeneralParser(text);
            }
        }
    }

    [Theory]
    [InlineData("1234567890123456")]
    [InlineData("0.00000000000000000000001")]
    [InlineData("1.000000000000000")]
    public void Utf8NumberParser_AdjacentPrecisionBoundariesMatchGeneralParser(string text) {
        byte[] utf8 = Encoding.UTF8.GetBytes(text);
        Assert.True(ExcelUtf8NumberParser.TryParseDouble(utf8, out double actual));
        Assert.True(Utf8Parser.TryParse(utf8, out double expected, out int consumed));
        Assert.Equal(utf8.Length, consumed);
        Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));
    }

    /// <summary>
    /// Compares with .NET's correctly rounded parser. .NET Framework's parser can be one unit
    /// off, so there the result is checked against the exact value; zero stays unsigned there,
    /// as Framework's parsers return it.
    /// </summary>
    private static void AssertParserMatchesGeneralParser(string text) {
        byte[] utf8 = Encoding.UTF8.GetBytes(text);
        Assert.True(ExcelUtf8NumberParser.TryParseDouble(utf8, out double actual));
#if NETFRAMEWORK
        AssertNearestDouble(text, actual);
#else
        Assert.True(Utf8Parser.TryParse(utf8, out double expected, out int consumed));
        Assert.Equal(utf8.Length, consumed);
        Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));
#endif
    }

#if NETFRAMEWORK
    /// <summary>Checks that a plain decimal literal parsed to its nearest double, ties to even.</summary>
    private static void AssertNearestDouble(string text, double actual) {
        string unsigned = text.TrimStart('-', '+');
        int point = unsigned.IndexOf('.');
        int scale = point < 0 ? 0 : unsigned.Length - point - 1;
        System.Numerics.BigInteger numerator = System.Numerics.BigInteger.Parse(
            "0" + unsigned.Replace(".", string.Empty), CultureInfo.InvariantCulture);
        Assert.Equal(text[0] == '-' && !numerator.IsZero, BitConverter.DoubleToInt64Bits(actual) < 0);
        System.Numerics.BigInteger denominator = System.Numerics.BigInteger.Pow(10, scale);
        long bits = BitConverter.DoubleToInt64Bits(Math.Abs(actual));
        if (numerator.IsZero) {
            Assert.Equal(0, bits);
            return;
        }

        // The exact value must lie between the midpoints to both neighbours of the result.
        (System.Numerics.BigInteger lowNumerator, System.Numerics.BigInteger lowDenominator) = Midpoint(bits - 1, bits);
        (System.Numerics.BigInteger highNumerator, System.Numerics.BigInteger highDenominator) = Midpoint(bits, bits + 1);
        int low = (numerator * lowDenominator).CompareTo(lowNumerator * denominator);
        int high = (numerator * highDenominator).CompareTo(highNumerator * denominator);
        bool even = (bits & 1) == 0;
        Assert.True(low > 0 || low == 0 && even, text + " parsed above its nearest double.");
        Assert.True(high < 0 || high == 0 && even, text + " parsed below its nearest double.");
    }

    private static (System.Numerics.BigInteger Numerator, System.Numerics.BigInteger Denominator) Midpoint(long first, long second) {
        // Positive doubles are mantissa * 2^exponent; scale both to the smaller exponent.
        static (System.Numerics.BigInteger Mantissa, int Exponent) Exact(long bits) {
            int exponent = (int)(bits >> 52);
            long fraction = bits & 0xFFFFFFFFFFFFFL;
            return exponent == 0 ? (fraction, -1074) : (fraction | (1L << 52), exponent - 1075);
        }

        (System.Numerics.BigInteger a, int aExponent) = Exact(Math.Max(first, 0));
        (System.Numerics.BigInteger b, int bExponent) = Exact(second);
        int exponent = Math.Min(aExponent, bExponent) - 1;
        System.Numerics.BigInteger sum = (a << (aExponent - exponent - 1)) + (b << (bExponent - exponent - 1));
        return exponent >= 0
            ? (sum << exponent, System.Numerics.BigInteger.One)
            : (sum, System.Numerics.BigInteger.One << -exponent);
    }
#endif

    private static long NextInt64(Random random, long minimum, long maximum) =>
#if NET6_0_OR_GREATER
        random.NextInt64(minimum, maximum);
#else
        minimum + (long)(random.NextDouble() * (maximum - minimum));
#endif

    [Theory]
    [InlineData("1234567890123456")]
    [InlineData("1.234567890123456")]
    [InlineData("1E+100")]
    [InlineData("5e-324")]
    public void Utf8NumberParser_UncommonFormsMatchGeneralParser(string text) {
        byte[] utf8 = Encoding.UTF8.GetBytes(text);
        Assert.True(Utf8Parser.TryParse(utf8, out double expected, out int consumed));
        Assert.Equal(utf8.Length, consumed);

        Assert.True(ExcelUtf8NumberParser.TryParseDouble(utf8, out double actual));
        Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));
    }

    [Theory]
    [InlineData("")]
    [InlineData("+")]
    [InlineData(".")]
    [InlineData("1.2.3")]
    [InlineData("12x")]
    [InlineData(" 12")]
    [InlineData("12 ")]
    public void Utf8NumberParser_RejectsIncompleteOrPartiallyConsumedInput(string text) {
        Assert.False(ExcelUtf8NumberParser.TryParseDouble(Encoding.UTF8.GetBytes(text), out _));
    }

    [Fact]
    public void Utf8NumberParser_RandomCommonValuesMatchGeneralParserBitForBit() {
        var random = new Random(0x5EED);
        for (int index = 0; index < 10_000; index++) {
            int whole = random.Next(0, 1_000_000_000);
            int fraction = random.Next(0, 1_000_000);
            string sign = (index & 1) == 0 ? string.Empty : "-";
            string text = string.Format(
                CultureInfo.InvariantCulture,
                "{0}{1}.{2:D6}",
                sign,
                whole,
                fraction);
            AssertParserMatchesGeneralParser(text);
        }
    }
}
