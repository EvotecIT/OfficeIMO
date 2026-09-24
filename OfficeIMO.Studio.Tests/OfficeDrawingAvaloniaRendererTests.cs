using OfficeIMO.Drawing;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class OfficeDrawingAvaloniaRendererTests {
    [Theory]
    [InlineData(OfficeTextAlignment.Left)]
    [InlineData(OfficeTextAlignment.Justify)]
    [InlineData(OfficeTextAlignment.Center)]
    [InlineData(OfficeTextAlignment.Right)]
    public void WiderSubstituteRunIsCompressedIntoItsBox(OfficeTextAlignment alignment) {
        (double offsetX, double scaleX) = OfficeDrawingAvaloniaRenderer.FitSingleLine(120D, 80D, alignment);

        Assert.Equal(0D, offsetX, 6);
        Assert.Equal(80D, 120D * scaleX, 6);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Left, 0D)]
    [InlineData(OfficeTextAlignment.Center, 20D)]
    [InlineData(OfficeTextAlignment.Right, 40D)]
    public void NarrowerRunKeepsItsWidthAndHonorsAlignment(OfficeTextAlignment alignment, double expectedOffset) {
        (double offsetX, double scaleX) = OfficeDrawingAvaloniaRenderer.FitSingleLine(60D, 100D, alignment);

        Assert.Equal(1D, scaleX);
        Assert.Equal(expectedOffset, offsetX, 6);
    }

    [Theory]
    [InlineData(0D, 50D)]
    [InlineData(50D, 0D)]
    [InlineData(double.NaN, 50D)]
    [InlineData(50D, double.PositiveInfinity)]
    public void DegenerateMeasurementsLeaveTheRunUntransformed(double measuredWidth, double boxWidth) {
        Assert.Equal((0D, 1D), OfficeDrawingAvaloniaRenderer.FitSingleLine(measuredWidth, boxWidth, OfficeTextAlignment.Right));
    }
}
