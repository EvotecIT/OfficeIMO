using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeIndependentNumberTests {
    private static double NextUp(double value) => BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(value) + 1);

    private static string Text(decimal value) => value.ToString(CultureInfo.InvariantCulture);

    // Expected text is what (decimal)value produced on .NET Framework and .NET up to 10,
    // including the scale, so the conversion reads the same on every runtime.
    public static TheoryData<double, string> Doubles => new() {
        // A native Project cost of 500000 stored one ulp high, as in a real .mpp.
        { NextUp(500000d), "500000" },
        { 0.1, "0.1" },
        { 0.29, "0.29" },
        { -1.1, "-1.1" },
        { 1.5, "1.5" },
        { 1234.56, "1234.56" },
        { -0.001, "-0.001" },
        { 1.4142135623730951, "1.4142135623731" },
        { 0d, "0" },
        // Exact ties at the fifteenth digit round to even.
        { 1000000000000.125, "1000000000000.12" },
        { -1000000000000.375, "-1000000000000.38" },
        { 100000000000000.5, "100000000000000" },
        { 123456789.12345679, "123456789.123457" },
        { 1e20, "100000000000000000000" },
        { 79228162514264328797450928128d, "79228162514264300000000000000" },
        { 1.23456789012345e-20, "0.0000000000000000000123456789" },
        { 1e-28, "0.0000000000000000000000000001" },
        { 1e-29, "0" },
        { -0d, "0" },
        { double.Epsilon, "0" },
    };

    [Theory]
    [MemberData(nameof(Doubles))]
    public void DoublesRoundToFifteenSignificantDigits(double value, string expected) {
        Assert.Equal(expected, Text(OfficeDecimalConversion.FromDouble(value)));
        Assert.Equal(expected, Text(OfficeDecimalConversion.FromObject(value, CultureInfo.InvariantCulture)));
        Assert.Equal(expected, Text((decimal)OfficeDecimalConversion.ChangeType(value, typeof(decimal), CultureInfo.InvariantCulture)!));
    }

    [Theory]
    [InlineData(0.1f, "0.1")]
    [InlineData(1.1f, "1.1")]
    [InlineData(16777217f, "16777220")]
    public void FloatsRoundToSevenSignificantDigits(float value, string expected) {
        Assert.Equal(expected, Text(OfficeDecimalConversion.FromSingle(value)));
        Assert.Equal(expected, Text(OfficeDecimalConversion.FromObject(value, CultureInfo.InvariantCulture)));
    }

    // Expected values are what Math.Round produced on .NET Framework and .NET up to 10;
    // .NET 11 rounds the exact binary value and returns 2.67, 31267539659173.9 and 15.01.
    [Theory]
    [InlineData(2.675, 2, MidpointRounding.ToEven, 2.68)]
    [InlineData(2.675, 2, MidpointRounding.AwayFromZero, 2.68)]
    [InlineData(31267539659173.95, 1, MidpointRounding.ToEven, 31267539659174d)]
    [InlineData(15.005, 2, MidpointRounding.ToEven, 15d)]
    [InlineData(15.005, 2, MidpointRounding.AwayFromZero, 15.01)]
    [InlineData(0.125, 2, MidpointRounding.ToEven, 0.12)]
    [InlineData(0.125, 2, MidpointRounding.AwayFromZero, 0.13)]
    public void RoundingScalesBeforeRoundingAsOlderRuntimesDid(double value, int digits, MidpointRounding mode, double expected) =>
        Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(OfficeRounding.Round(value, digits, mode)));

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(79228162514264337593543950336d)]
    [InlineData(-79228162514264337593543950336d)]
    [InlineData(1e30)]
    public void ValuesOutsideTheDecimalRangeOverflowLikeTheCast(double value) {
        Assert.Throws<OverflowException>(() => OfficeDecimalConversion.FromDouble(value));
        Assert.False(OfficeDecimalConversion.TryFromDouble(value, out _));
    }
}
