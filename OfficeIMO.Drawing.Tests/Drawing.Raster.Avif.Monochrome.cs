using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native-produced primary monochrome items keep grayscale range separate from auxiliary-alpha semantics.</summary>
public sealed class DrawingAvifMonochromeTests {
    [Theory]
    [InlineData("avif-monochrome-full", true, false)]
    [InlineData("avif-monochrome-limited", false, false)]
    [InlineData("avif-monochrome-full-alpha", true, true)]
    [InlineData("avif-monochrome-limited-alpha", false, true)]
    public void PrimaryPlaneRangeDoesNotAssignTheAuxiliaryAlphaRole(string name, bool fullRange, bool hasAlpha) {
        byte[] bytes = File.ReadAllBytes(Asset(name));
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        var primary = container!.Color;
        Assert.True(primary.Monochrome);
        Assert.False(primary.IsAlpha);
        Assert.Equal(hasAlpha, container.Alpha != null);
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, primary, options, out var sequence));
        Assert.Equal(fullRange, sequence!.Color.FullRange);
        Assert.Equal(fullRange, primary.ColorDescription!.FullRange);
        if (hasAlpha) {
            Assert.True(container.Alpha!.IsAlpha);
            Assert.True(container.Alpha.Monochrome);
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, container.Alpha, options, out var alpha));
            Assert.True(alpha!.Color.FullRange);
        }
        // The same valid limited-range monochrome bitstream is invalid when selected as auxiliary alpha.
        var auxiliary = new OfficeAvifImageItem(primary.Id, primary.Width, primary.Height,
            primary.Offset, primary.Length, primary.Configuration, true, null);
        Assert.Equal(fullRange, OfficeAv1StillSequenceReader.TryRead(bytes, auxiliary, options, out _));
    }

    [Theory]
    [InlineData("avif-monochrome-full")]
    [InlineData("avif-opaque")]
    public void ConfigurationCannotContradictDeclaredPixelChannels(string name) {
        byte[] bytes = File.ReadAllBytes(Asset(name));
        int config = Find(bytes, "av1C") + 4;
        bytes[config + 2] ^= 16; // Single-plane <-> three-plane declaration; pixi remains unchanged.
        var codec = new AcceptingCodec();
        Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image);
        Assert.Equal(0, codec.Calls);
    }

    private static int Find(byte[] bytes, string text) {
        byte[] marker = Encoding.ASCII.GetBytes(text);
        for (int i = 0; i <= bytes.Length - marker.Length; i++)
            if (bytes.Skip(i).Take(marker.Length).SequenceEqual(marker)) return i;
        throw new InvalidOperationException("Missing native fixture property.");
    }

    private static string Asset(string name) => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", name + ".avif");

    private sealed class AcceptingCodec : IOfficeRasterImageCodec {
        internal int Calls;
        public bool TryDecode(byte[] bytes, string? mimeType, out OfficeRasterImage? image) {
            Calls++;
            image = OfficeRasterImage.FromRgba32(49, 33, new byte[49 * 33 * 4]);
            return true;
        }
    }
}
