namespace OfficeIMO.Project.Tests;

public sealed class ProjectNumberTests {
    private static double NextUp(double value) => BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(value) + 1);

    [Fact]
    public void RepresentationErrorIsRoundedAway() {
        // a native cost of 500000 stored one ulp high, as in a real .mpp
        var value = ProjectNumber.ToDecimal(NextUp(500000d));
        Assert.Equal(500000m, value);
        Assert.Equal("500000", value.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(0.1, "0.1")]
    [InlineData(1234.56, "1234.56")]
    [InlineData(-0.001, "-0.001")]
    [InlineData(1.4142135623730951, "1.41421356237310")]
    [InlineData(0d, "0")]
    public void RoundsToFifteenSignificantDigits(double input, string expected) =>
        Assert.Equal(decimal.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), ProjectNumber.ToDecimal(input));

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(1e30)]
    public void NonFiniteOrOutOfRangeThrowsLikeTheCast(double input) =>
        Assert.Throws<OverflowException>(() => ProjectNumber.ToDecimal(input));
}
