using System;
using Xunit;

namespace OfficeIMO.Drawing.CodeGlyphX.Tests;

public sealed class CodeGlyphRasterImageCodecTests {
    [Fact]
    public void IndependentLossyWebpPreservesColorsThroughOptionalDrawingBoundary() {
        byte[] bytes = Convert.FromBase64String("UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoQABAAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=");
        var options = new OfficeRasterDecodeOptions { ImageCodec = new CodeGlyphRasterImageCodec() };
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, options, out var image, out var info));
        Assert.True(info.Succeeded);
        Assert.Equal(16, image!.Width);
        Assert.Equal(16, image.Height);
        var color = image.GetPixel(8, 8);
        Assert.InRange((int)color.R, 254, 255);
        Assert.InRange((int)color.G, 0, 2);
        Assert.InRange((int)color.B, 0, 1);
    }

    [Fact]
    public void InvalidRasterPayloadReturnsFailure() {
        Assert.False(new CodeGlyphRasterImageCodec().TryDecode(new byte[] { 1, 2, 3 }, null, out var image));
        Assert.Null(image);
    }
}
