using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingRasterFrameRetainedBudgetTests {
    [Theory]
    [InlineData(256L * 1024 * 1024)]
    [InlineData(long.MaxValue)]
    public void AuxiliaryRetentionRejectsBeforeMapping(long retainedBytes) {
        var image = new OfficeRasterImage(2, 1);
        image.SetPixel(0, 0, OfficeColor.Red);
        var frames = new OfficeRasterFrames(new[] { new OfficeRasterFrame(image) });
        bool mapped = false;

        Assert.Throws<ArgumentException>(() => frames.Transform(source => {
            mapped = true;
            return source.Clone();
        }, additionalRetainedBytes: retainedBytes));

        Assert.False(mapped);
        Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
    }

    [Fact]
    public void AuxiliaryRetentionPreservesCompleteFrameContract() {
        var image = new OfficeRasterImage(2, 1);
        image.SetPixel(0, 0, OfficeColor.Red);
        var duration = TimeSpan.FromMilliseconds(150);
        var frames = new OfficeRasterFrames(new[] { new OfficeRasterFrame(image, duration) }, 3);

        var result = frames.Transform(source => source.Clone(), additionalRetainedBytes: 128L * 1024 * 1024);

        Assert.Equal(3, result.PlayCount);
        Assert.Equal(duration, result[0].Duration);
        Assert.Equal(image.GetPixels(), result[0].Image.GetPixels());
        Assert.NotSame(image, result[0].Image);
        Assert.Throws<ArgumentOutOfRangeException>(() => frames.Transform(source => source.Clone(), additionalRetainedBytes: -1));
    }
}
