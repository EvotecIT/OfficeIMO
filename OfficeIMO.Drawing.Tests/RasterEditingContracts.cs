using OfficeIMO.Drawing;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RasterEditingContracts {
    [Fact]
    public void ExpansionPlanningMatchesRotatedAndSkewedPixels() {
        var source = new OfficeRasterImage(7, 4, OfficeColor.Red);
        var rotationSize = OfficeRasterTransforms.GetRotatedSize(source, 35);
        var rotation = OfficeRasterTransforms.Rotate(source, 35);
        Assert.Equal((rotation.Width, rotation.Height), rotationSize);
        var skewSize = OfficeRasterTransforms.GetSkewedSize(source, 12, -9);
        var skew = OfficeRasterTransforms.Skew(source, 12, -9);
        Assert.Equal((skew.Width, skew.Height), skewSize);
        Assert.Equal(OfficeColor.Red, source.GetPixel(0, 0));
    }

    [Theory]
    [InlineData(1, 3, 2, "123456")]
    [InlineData(2, 3, 2, "321654")]
    [InlineData(3, 3, 2, "654321")]
    [InlineData(4, 3, 2, "456123")]
    [InlineData(5, 2, 3, "142536")]
    [InlineData(6, 2, 3, "415263")]
    [InlineData(7, 2, 3, "635241")]
    [InlineData(8, 2, 3, "362514")]
    public void ExifOrientationPreservesExactPixelsAndIndependentOwnership(int orientation, int width, int height, string expected) {
        var source = new OfficeRasterImage(3, 2);
        for (int y = 0; y < 2; y++) {
            for (int x = 0; x < 3; x++) { source.SetPixel(x, y, OfficeColor.FromRgba((byte)(y * 3 + x + 1), 10, 20, 128)); }
        }
        var oriented = OfficeRasterTransforms.AutoOrient(source, orientation);
        Assert.Equal(width, oriented.Width);
        Assert.Equal(height, oriented.Height);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Equal(OfficeColor.FromRgba((byte)(expected[index] - '0'), 10, 20, 128), oriented.GetPixel(index % width, index / width));
        }
        oriented.SetPixel(0, 0, OfficeColor.White);
        Assert.Equal((byte)1, source.GetPixel(0, 0).R);
    }

    [Fact]
    public void CropAndQuarterRotationRetainPixelCoordinates() {
        var source = new OfficeRasterImage(4, 3, OfficeColor.Red);
        source.SetPixel(1, 1, OfficeColor.Blue);
        var crop = OfficeRasterTransforms.Crop(source, 1, 1, 2, 2);
        Assert.Equal(OfficeColor.Blue, crop.GetPixel(0, 0));
        var rotate = OfficeRasterTransforms.Rotate(crop, 90);
        Assert.Equal(OfficeColor.Blue, rotate.GetPixel(1, 0));
        crop.SetPixel(0, 0, OfficeColor.White);
        Assert.Equal(OfficeColor.Blue, source.GetPixel(1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterTransforms.Crop(source, 3, 0, 2, 2));
    }

    [Fact]
    public void MasksMultiplyExistingAlphaAndRetainUnmaskedColors() {
        var source = new OfficeRasterImage(21, 21, OfficeColor.FromRgba(40, 80, 120, 128));
        var ellipse = OfficeRasterTransforms.MaskEllipse(source, 0, 0, 21, 21);
        Assert.Equal((byte)0, ellipse.GetPixel(0, 0).A);
        Assert.Equal(source.GetPixel(10, 10), ellipse.GetPixel(10, 10));
        Assert.Contains(ellipse.GetPixels().Where((_, index) => index % 4 == 3), alpha => alpha > 0 && alpha < 128);
        var rounded = OfficeRasterTransforms.MaskRoundedRectangle(source, 5);
        Assert.Equal((byte)0, rounded.GetPixel(0, 0).A);
        Assert.Equal(source.GetPixel(10, 10), rounded.GetPixel(10, 10));
        Assert.Equal((byte)128, source.GetPixel(0, 0).A);
    }

    [Fact]
    public void TextLayoutWrapsAndExplicitClipContainsAllEffects() {
        var layout = OfficeRasterText.Measure("one two three four", 14, "Arial", 35);
        Assert.True(layout.Lines.Count > 1);
        var image = new OfficeRasterImage(100, 70, OfficeColor.White);
        OfficeRasterText.Draw(image, "one two three four", 20, 15, 35, 25, OfficeColor.Black,
            new OfficeRasterTextOptions { FontSize = 14, FontFamily = "Arial", Wrap = true, Clip = true,
                ShadowColor = OfficeColor.Red, ShadowOffsetX = 4, ShadowOffsetY = 3, OutlineColor = OfficeColor.Blue, OutlineWidth = 2 });
        bool ink = false;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                if (x < 20 || x >= 55 || y < 15 || y >= 40) { Assert.Equal(OfficeColor.White, image.GetPixel(x, y)); }
                else if (image.GetPixel(x, y) != OfficeColor.White) { ink = true; }
            }
        }
        Assert.True(ink);
    }

    [Fact]
    public void TextOutlineAndShadowExpandVisibleCoverage() {
        var plain = new OfficeRasterImage(140, 60, OfficeColor.White);
        var effects = plain.Clone();
        OfficeRasterText.Draw(plain, "Office", 10, 10, 110, 30, OfficeColor.Black, 20, "Arial");
        OfficeRasterText.Draw(effects, "Office", 10, 10, 110, 30, OfficeColor.Black,
            new OfficeRasterTextOptions { FontSize = 20, FontFamily = "Arial",
                ShadowColor = OfficeColor.Red, ShadowOffsetX = 5, ShadowOffsetY = 5, OutlineColor = OfficeColor.Blue, OutlineWidth = 2 });
        int plainInk = 0, effectsInk = 0, red = 0, blue = 0;
        for (int y = 0; y < plain.Height; y++) {
            for (int x = 0; x < plain.Width; x++) {
                if (plain.GetPixel(x, y) != OfficeColor.White) { plainInk++; }
                OfficeColor pixel = effects.GetPixel(x, y);
                if (pixel != OfficeColor.White) { effectsInk++; }
                if (pixel.R > pixel.B && pixel.R > pixel.G) { red++; }
                if (pixel.B > pixel.R && pixel.B > pixel.G) { blue++; }
            }
        }
        Assert.True(effectsInk > plainInk);
        Assert.True(red > 0);
        Assert.True(blue > 0);
    }

    [Fact]
    public void EditingCancellationAndInvalidGeometryFailBeforeMutation() {
        var source = new OfficeRasterImage(10, 10, OfficeColor.Red);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterTransforms.AutoOrient(source, 1, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => OfficeRasterTransforms.MaskEllipse(source, 0, 0, 10, 10, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => OfficeRasterText.Draw(source, "hello", 0, 0, 10, 10, OfficeColor.White, cancellationToken: cancellation.Token));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterText.Draw(source, "hello", double.NaN, 0, 10, 10, OfficeColor.White));
        Assert.Equal(OfficeColor.Red, source.GetPixel(0, 0));
    }
}
