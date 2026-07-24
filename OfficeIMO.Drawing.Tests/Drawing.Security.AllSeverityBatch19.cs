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
}
