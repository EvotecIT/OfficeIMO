using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void AddPositionedTextRetainsNineParameterBinarySignature() {
        Type[] legacyParameters = {
            typeof(string),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(OfficeFontInfo),
            typeof(OfficeColor?),
            typeof(OfficeTextAlignment),
            typeof(double?)
        };

        Assert.NotNull(typeof(OfficeDrawing).GetMethod(
            nameof(OfficeDrawing.AddPositionedText),
            legacyParameters));
    }

    [Fact]
    public void AddPositionedTextRetainsNamedAdvanceSourceContract() {
        OfficeDrawing drawing = new OfficeDrawing(100D, 40D)
            .AddPositionedText(
                "positioned",
                0D,
                0D,
                80D,
                20D,
                textAdvanceWidth: 60D);

        Assert.Equal(
            60D,
            Assert.IsType<OfficeDrawingText>(Assert.Single(drawing.Elements)).TextAdvanceWidth);
    }
}
