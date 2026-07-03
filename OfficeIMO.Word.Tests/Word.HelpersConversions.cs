using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private const double Tolerance = 0.0001;

    [Theory]
    [InlineData(0.1)]
    [InlineData(1)]
    [InlineData(2.54)]
    [InlineData(3.5)]
    public void Test_CentimetersEmusConversions(double centimeters) {
        int expectedEmus = (int)(centimeters * 360000);
        long expectedEmus64 = (long)(centimeters * 360000);

        var emus = Helpers.ConvertCentimetersToEmus(centimeters);
        var emus64 = Helpers.ConvertCentimetersToEmusInt64(centimeters);

        Assert.True(emus.HasValue);
        Assert.Equal(expectedEmus, emus);
        Assert.Equal(expectedEmus64, emus64);
        var cm = Helpers.ConvertEmusToCentimeters(emus.Value);
        Assert.True(cm.HasValue);
        Assert.Equal(centimeters, cm.Value, 5);
        Assert.Equal(centimeters, Helpers.ConvertEmusToCentimeters(emus64), 5);
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(1)]
    [InlineData(2.54)]
    [InlineData(3.5)]
    public void Test_CentimetersTwipsConversions(double centimeters) {
        int expectedTwips = (int)Math.Round(centimeters * 567.0);
        uint expectedTwipsUint = (uint)Math.Round(centimeters * 567.0);

        var twips = Helpers.ConvertCentimetersToTwips(centimeters);
        var twipsUint = Helpers.ConvertCentimetersToTwipsUInt32(centimeters);

        Assert.Equal(expectedTwips, twips);
        Assert.Equal(expectedTwipsUint, twipsUint);
        Assert.Equal(Math.Round(centimeters, 2), Helpers.ConvertTwipsToCentimeters(twips));
        Assert.Equal(Math.Round(centimeters, 2), Helpers.ConvertTwipsToCentimeters(twipsUint));
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(1)]
    [InlineData(2.54)]
    [InlineData(3.5)]
    public void Test_CentimetersPointsConversions(double centimeters) {
        double expectedPoints = (centimeters / 2.54) * 72;

        var points = Helpers.ConvertCentimetersToPoints(centimeters);
        Assert.Equal(expectedPoints, points, 5);
        Assert.Equal(centimeters, Helpers.ConvertPointsToCentimeters(points), 5);
    }

    [Theory]
    [InlineData(20)]
    [InlineData(567)]
    [InlineData(1000)]
    [InlineData(1440)]
    public void Test_TwipsPointsConversions(int twips) {
        double expectedPoints = Math.Round(twips / 20.0, 2);

        var points1 = Helpers.ConvertTwipsToPoints(twips);
        var points2 = Helpers.ConvertTwipsToPoints((uint)twips);

        Assert.Equal(expectedPoints, points1);
        Assert.Equal(expectedPoints, points2);
        Assert.Equal(twips, Helpers.ConvertPointsToTwips(points1));
        Assert.Equal((uint)twips, Helpers.ConvertPointsToTwipsUInt32(points1));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(50)]
    [InlineData(72)]
    public void Test_PointsEmusConversions(double points) {
        long expectedEmus = (long)(points * 12700.0);
        var emus = Helpers.ConvertPointsToEmusInt64(points);
        Assert.Equal(expectedEmus, emus);
        Assert.Equal(points, Helpers.ConvertEmusToPoints(emus), 5);
    }
}
