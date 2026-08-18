using System.Buffers.Text;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
#if NET8_0_OR_GREATER
    [Theory]
    [InlineData("0")]
    [InlineData("-0")]
    [InlineData("+0")]
    [InlineData("42")]
    [InlineData("-42")]
    [InlineData(".5")]
    [InlineData("1.")]
    [InlineData("000000000000000123")]
    [InlineData("0.0000000000000000000001")]
    [InlineData("123456789012345")]
    [InlineData("1.23456789012345")]
    public void Utf8NumberParser_CommonFixedPointMatchesGeneralParser(string text) {
        byte[] utf8 = Encoding.UTF8.GetBytes(text);
        Assert.True(Utf8Parser.TryParse(utf8, out double expected, out int consumed));
        Assert.Equal(utf8.Length, consumed);

        Assert.True(ExcelUtf8NumberParser.TryParseCommonFixedPoint(utf8, out double actual));
        Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));
    }
#endif

    [Theory]
    [InlineData("1234567890123456")]
    [InlineData("1.234567890123456")]
    [InlineData("1E+100")]
    [InlineData("5e-324")]
    public void Utf8NumberParser_UncommonFormsRetainGeneralParserFallback(string text) {
        byte[] utf8 = Encoding.UTF8.GetBytes(text);
#if NET8_0_OR_GREATER
        Assert.False(ExcelUtf8NumberParser.TryParseCommonFixedPoint(utf8, out _));
#endif
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
            byte[] utf8 = Encoding.UTF8.GetBytes(text);

            Assert.True(Utf8Parser.TryParse(utf8, out double expected, out int consumed));
            Assert.Equal(utf8.Length, consumed);
#if NET8_0_OR_GREATER
            Assert.True(ExcelUtf8NumberParser.TryParseCommonFixedPoint(utf8, out double actual));
#else
            Assert.True(ExcelUtf8NumberParser.TryParseDouble(utf8, out double actual));
#endif
            Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));
        }
    }
}
