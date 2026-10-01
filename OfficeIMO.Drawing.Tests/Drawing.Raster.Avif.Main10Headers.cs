using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent Main10 still items establish the bit-depth boundary before widened reconstruction.</summary>
public sealed class DrawingAvifMain10HeaderTests {
    [Theory]
    [InlineData("avif-main10-420-full", false, true, false)]
    [InlineData("avif-main10-420-limited", false, false, false)]
    [InlineData("avif-main10-420-full-alpha", false, true, true)]
    [InlineData("avif-main10-420-limited-alpha", false, false, true)]
    [InlineData("avif-main10-mono-full", true, true, false)]
    [InlineData("avif-main10-mono-limited", true, false, false)]
    [InlineData("avif-main10-mono-full-alpha", true, true, true)]
    [InlineData("avif-main10-mono-limited-alpha", true, false, true)]
    public void ContainerAndReducedHeadersPreserveIndependentMain10Selection(string name, bool mono, bool full, bool alpha) {
        byte[] bytes = File.ReadAllBytes(Asset(name));
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        Assert.Equal(10, container!.Color.BitDepth);
        Assert.Equal(mono, container!.Color.Monochrome);
        Assert.Equal(alpha, container.Alpha != null);
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, container.Color, options, out var sequence));
        Assert.Equal(full, sequence!.Color.FullRange);
        Assert.Equal(10, sequence.BitDepth);
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, container.Color, sequence, options, out var frame));
        Assert.Equal(49, frame!.UpscaledWidth);
        Assert.Equal(33, frame.Height);
        Assert.Equal(10, frame.BitDepth);
        if (alpha) {
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, container.Alpha!, options, out var auxiliary));
            Assert.True(auxiliary!.Monochrome);
            Assert.Equal(10, container.Alpha!.BitDepth);
            Assert.Equal(10, auxiliary.BitDepth);
            Assert.True(auxiliary.Color.FullRange);
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, container.Alpha!, auxiliary, options, out var alphaFrame));
            Assert.Equal(frame.UpscaledWidth, alphaFrame!.UpscaledWidth);
            Assert.Equal(frame.Height, alphaFrame.Height);
            Assert.Equal(10, alphaFrame.BitDepth);
        }
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, name + ".avif", out var metadata));
        Assert.Equal(49, metadata.Width); Assert.Equal(33, metadata.Height);
        Assert.False(OfficeImageReader.TryValidateContent(bytes, name + ".avif", out _));
        Assert.Equal(10, OfficeAv1FrameReconstructor.Decode(bytes, sequence, frame, options, OfficeAv1ReconstructionStage.Restored).BitDepth);
        var codec = new AcceptingCodec();
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image); Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData("avif-main10-420-full")]
    [InlineData("avif-main10-mono-full")]
    public void PixelDepthAndBitstreamDepthMustAgreeWithTheConfiguration(string name) {
        byte[] original = File.ReadAllBytes(Asset(name));
        var options = new OfficeRasterDecodeOptions();
        int config = Find(original, "av1C") + 4, pixi = Find(original, "pixi") + 8;
        byte[] mixed = (byte[])original.Clone(); mixed[pixi + mixed[pixi]] = 8;
        Assert.False(OfficeAvifContainerReader.TryRead(mixed, options, out _));
        byte[] relabeled = (byte[])original.Clone(); relabeled[config + 2] &= 0xbf;
        Assert.False(OfficeAvifContainerReader.TryRead(relabeled, options, out _));
        for (int i = 1; i <= relabeled[pixi]; i++) relabeled[pixi + i] = 8;
        Assert.True(OfficeAvifContainerReader.TryRead(relabeled, options, out var declaredEight));
        Assert.False(OfficeAv1StillSequenceReader.TryRead(relabeled, declaredEight!.Color, options, out _));
        var codec = new AcceptingCodec();
        foreach (byte[] invalid in new[] { mixed, relabeled }) {
            Assert.False(OfficeRasterImageDecoder.TryDecode(invalid, new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
            Assert.Null(image);
        }
        Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData("avif-main10-mono-full", true)]
    [InlineData("avif-main10-mono-limited", false)]
    public void TenBitPrimaryRangeDoesNotRelaxTheAuxiliaryAlphaContract(string name, bool fullRange) {
        byte[] bytes = File.ReadAllBytes(Asset(name));
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        var primary = container!.Color;
        var auxiliary = new OfficeAvifImageItem(primary.Id, primary.Width, primary.Height, primary.Offset,
            primary.Length, primary.Configuration, true, null);
        Assert.Equal(fullRange, OfficeAv1StillSequenceReader.TryRead(bytes, auxiliary, options, out _));
    }

    private static int Find(byte[] bytes, string text) {
        byte[] marker = Encoding.ASCII.GetBytes(text);
        for (int i = 0; i <= bytes.Length - marker.Length; i++)
            if (bytes.Skip(i).Take(marker.Length).SequenceEqual(marker)) return i;
        throw new InvalidOperationException("Missing native fixture property.");
    }

    private sealed class AcceptingCodec : IOfficeRasterImageCodec {
        internal int Calls;
        public bool TryDecode(byte[] bytes, string? mimeType, out OfficeRasterImage? image) {
            Calls++;
            image = OfficeRasterImage.FromRgba32(49, 33, new byte[49 * 33 * 4]);
            return true;
        }
    }

    private static string Asset(string name) => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", name + ".avif");
}
