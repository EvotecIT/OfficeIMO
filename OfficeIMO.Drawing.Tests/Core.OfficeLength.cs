using System.Globalization;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeLengthTests {
    [Theory]
    [InlineData("en-US")]
    [InlineData("pl-PL")]
    [InlineData("de-DE")]
    public void InvariantParsingAndFormattingRoundTripAcrossCultures(string cultureName) {
        CultureInfo saved = CultureInfo.CurrentCulture;
        CultureInfo savedUi = CultureInfo.CurrentUICulture;
        try {
            var culture = CultureInfo.GetCultureInfo(cultureName);
            CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = culture;
            Assert.Equal(OfficeLength.Centimeters(2.5), OfficeLength.Parse("2.5cm"));
            Assert.Equal(OfficeLength.Points(12.5), OfficeLength.Parse("12.5"));
            foreach (long emus in new[] { 0L, 1L, -1L, 914400L, long.MaxValue, long.MinValue }) {
                OfficeLength value = OfficeLength.Emus(emus);
                foreach (string unit in new[] { "G", "pt", "cm", "mm", "in", "px", "emu" }) {
                    string formatted = value.ToString(unit, culture);
                    Assert.DoesNotContain(",", formatted);
                    Assert.Equal(value, OfficeLength.Parse(formatted));
                }
            }
            Assert.False(OfficeLength.TryParse("2,5cm", out _));
        } finally {
            CultureInfo.CurrentCulture = saved;
            CultureInfo.CurrentUICulture = savedUi;
        }
    }

    [Fact]
    public void PhysicalUnitRelationshipsAreExact() {
        OfficeLength inch = OfficeLength.Inches(1);
        Assert.Equal(914400, inch.ToEmus());
        Assert.Equal(inch, OfficeLength.Centimeters(2.54));
        Assert.Equal(inch, OfficeLength.Millimeters(25.4));
        Assert.Equal(inch, OfficeLength.Points(72));
        Assert.Equal(inch, OfficeLength.Pixels(96));
        Assert.Equal(72, inch.ToPoints());
        Assert.Equal(2.54, inch.ToCentimeters());
        Assert.Equal(25.4, inch.ToMillimeters());
        Assert.Equal(96, inch.ToPixels());
        Assert.Equal(1, inch.ToInches());
        Assert.Equal(OfficeLength.Zero, default);
    }

    [Theory]
    [InlineData(" .5 PT ", 6350)]
    [InlineData("+1in", 914400)]
    [InlineData("1e2mm", 3600000)]
    [InlineData("-0.5emu", -1)]
    [InlineData("0.5emu", 1)]
    public void SupportedInvariantSyntaxRoundsToWholeEmus(string input, long expected) => Assert.Equal(expected, OfficeLength.Parse(input).ToEmus());

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("1,000pt")]
    [InlineData("NaNpt")]
    [InlineData("Infinity")]
    [InlineData("12foo")]
    [InlineData("1 2pt")]
    [InlineData("1e100in")]
    [InlineData("9223372036854775808emu")]
    public void InvalidOrOverflowingInputFailsWithoutAValue(string? input) {
        Assert.False(OfficeLength.TryParse(input, out OfficeLength result));
        Assert.Equal(OfficeLength.Zero, result);
        if (input != null) Assert.Throws<FormatException>(() => OfficeLength.Parse(input));
        else Assert.Throws<ArgumentNullException>(() => OfficeLength.Parse(input!));
    }

    [Fact]
    public void CheckedArithmeticAndComparisonPreserveTheStoredValue() {
        var a = OfficeLength.Points(12);
        var b = OfficeLength.Points(6);
        Assert.Equal(a, b * 2);
        Assert.Equal(a, 2 * b);
        Assert.Equal(a, b + b);
        Assert.Equal(b, a - b);
        Assert.True(a > b && a >= b && b < a && b <= a && a != b);
        Assert.Equal(0, a.CompareTo((object)a));
        Assert.Equal(1, a.CompareTo(null));
        Assert.Throws<ArgumentException>(() => a.CompareTo("12pt"));
        Assert.Equal(a.GetHashCode(), OfficeLength.Parse("12pt").GetHashCode());
        Assert.Throws<OverflowException>(() => OfficeLength.Emus(long.MaxValue) + OfficeLength.Emus(1));
        Assert.Throws<OverflowException>(() => OfficeLength.Emus(long.MinValue) - OfficeLength.Emus(1));
        Assert.Throws<OverflowException>(() => OfficeLength.Emus(long.MaxValue) * 2);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLength.Points(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => a * double.PositiveInfinity);
        Assert.Throws<FormatException>(() => a.ToString("bogus", null));
    }

    [Fact]
    public void ScalarRoundingPreservesTheExactFiniteDoubleAtBoundaries() {
        Assert.Equal(OfficeLength.Zero, OfficeLength.Emus(1) * 0.49999999999999994);
        Assert.Equal(OfficeLength.Zero, OfficeLength.Emus(1) * -0.49999999999999994);
        Assert.Equal(OfficeLength.Emus(1), OfficeLength.Emus(1) * 0.5);
        Assert.Equal(OfficeLength.Emus(-1), OfficeLength.Emus(1) * -0.5);
        Assert.Equal(OfficeLength.Zero, OfficeLength.Zero * double.MaxValue);
        Assert.Equal(OfficeLength.Zero, OfficeLength.Emus(1) * double.Epsilon);
        Assert.Equal(OfficeLength.Emus(long.MinValue), OfficeLength.Emus(long.MinValue) * 1);
        Assert.Throws<OverflowException>(() => OfficeLength.Emus(long.MaxValue) * 1.0000000000000002);
        Assert.Throws<OverflowException>(() => OfficeLength.Emus(long.MinValue) * -1);
        Assert.Throws<OverflowException>(() => OfficeLength.Inches(double.MaxValue));
    }

    [Theory]
    [InlineData(0.025, 1, 0)]
    [InlineData(-0.025, -1, 0)]
    [InlineData(0.25, 5, 1)]
    [InlineData(-0.25, -5, -1)]
    public void CoarseUnitsRoundOnlyAtWriteTime(double points, long twips, long halfPoints) {
        OfficeLength value = OfficeLength.Points(points);
        Assert.Equal(twips, value.ToTwips());
        Assert.Equal(halfPoints, value.ToHalfPoints());
        Assert.Equal(value, OfficeLength.Parse(value.ToString()));
    }
}
