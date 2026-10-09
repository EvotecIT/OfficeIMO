using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterComparisonTests {
    [Fact]
    public void InvisibleRgbDoesNotCountAsAVisibleDifference() {
        var left = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(255, 0, 0, 0));
        var right = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(0, 0, 255, 0));
        OfficeRasterComparisonResult result = OfficeRasterComparison.Compare(left, right);
        Assert.Equal(1D, result.Similarity); Assert.Equal(0L, result.ChangedPixels);
        Assert.Equal(OfficeColor.Black, result.DifferenceImage.GetPixel(0, 0));
    }

    [Fact]
    public void PixelDifferenceReportsNormalizedMagnitudeAndAlphaChanges() {
        var black = new OfficeRasterImage(1, 1, OfficeColor.Black);
        var white = new OfficeRasterImage(1, 1, OfficeColor.White);
        OfficeRasterComparisonResult result = OfficeRasterComparison.Compare(black, white);
        Assert.Equal(.75D, result.MeanAbsoluteDifference); Assert.Equal(.25D, result.Similarity);
        Assert.Equal(1L, result.ChangedPixels); Assert.Equal(255, result.MaximumChannelDifference);
        Assert.Equal(OfficeColor.White, result.DifferenceImage.GetPixel(0, 0));
        var transparent = new OfficeRasterImage(1, 1);
        Assert.Equal(1D, OfficeRasterComparison.Compare(transparent, white).MeanAbsoluteDifference);
        Assert.Throws<ArgumentException>(() => OfficeRasterComparison.Compare(black, new OfficeRasterImage(2, 1)));
    }
}
