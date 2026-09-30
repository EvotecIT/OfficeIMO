using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpAlphaTests {
    [Theory]
    [InlineData("alpha-gradient")]
    [InlineData("alpha-noise")]
    [InlineData("opaque-control")]
    [InlineData("raw-filter-0")]
    [InlineData("raw-filter-1")]
    [InlineData("raw-filter-2")]
    [InlineData("raw-filter-3")]
    [InlineData("compressed-filter-0")]
    [InlineData("compressed-filter-1")]
    [InlineData("compressed-filter-2")]
    [InlineData("compressed-filter-3")]
    public void SeparateAlphaMatchesIndependentPixels(string name) {
        byte[] bytes = ReadFixture(name + ".webp");
        byte[] expected = ReadFixture(name + ".rgba");
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var metadata));
        Assert.Equal((49, 33), (metadata.Width, metadata.Height));
        Assert.True(OfficeRasterContainerInspector.TryInspect(bytes, out _));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
        byte[] actual = image!.GetPixels();
        Assert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < actual.Length; i++) {
            Assert.InRange(Math.Abs(expected[i] - actual[i]), 0, i % 4 == 3 ? 0 : 3);
        }
    }

    [Fact]
    public void AlphaDecodePreservesLimitsAndCancellation() {
        byte[] bytes = ReadFixture("alpha-gradient.webp");
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 49 * 33 - 1 }, out _, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { MaximumEncodedBytes = bytes.Length - 1 }, out _, out _));
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { CancellationToken = new CancellationToken(true) }, out _, out _));
        Assert.False(OfficeWebpCodec.TryDecode(bytes, CancellationToken.None,
            OfficeRasterGuards.MaximumDecodedBytes, out _));
    }

    [Theory]
    [InlineData("alpha-noise", false)]
    [InlineData("alpha-noise", true)]
    [InlineData("alpha-gradient", false)]
    public void MalformedAlphaDoesNotReachCallerCodec(string name, bool extraByte) {
        byte[] source = ReadFixture(name + ".webp");
        // The complete container and VP8 keyframe remain intact; only ALPH data is damaged.
        using var output = new MemoryStream();
        output.Write(source, 0, 12);
        for (int cursor = 12; cursor < source.Length;) {
            int length = BitConverter.ToInt32(source, cursor + 4);
            int next = cursor + 8 + length + (length & 1);
            if (System.Text.Encoding.ASCII.GetString(source, cursor, 4) == "ALPH") {
                int changedLength = extraByte ? length + 1 : Math.Max(2, length / 2);
                output.Write(source, cursor, 4);
                byte[] encodedLength = BitConverter.GetBytes(changedLength);
                output.Write(encodedLength, 0, 4);
                output.Write(source, cursor + 8, Math.Min(length, changedLength));
                if (extraByte) output.WriteByte(0);
                if ((changedLength & 1) != 0) output.WriteByte(0);
            } else {
                output.Write(source, cursor, next - cursor);
            }
            cursor = next;
        }
        byte[] bytes = output.ToArray();
        Array.Copy(BitConverter.GetBytes(bytes.Length - 8), 0, bytes, 4, 4);
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out _));
        Assert.False(OfficeWebpCodec.TryDecode(bytes, out _));
        Assert.False(OfficeRasterContainerInspector.TryInspect(bytes, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
        var codec = new CountingCodec();
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out _, out _));
        Assert.Equal(0, codec.Calls);
        var drawing = new OfficeDrawing(49, 33).AddImage(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 49, 33)));
        OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions { ImageCodec = codec });
        Assert.Equal(0, codec.Calls);
        var nearest = new OfficeDrawing(49, 33).AddImageWithInterpolation(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 49, 33)), interpolate: false);
        bool nearestRejected = false;
        try {
            OfficeDrawingSvgExporter.ToSvg(nearest, 1D, OfficeSvgSizeUnit.Pixel, imageCodec: codec);
        } catch (InvalidOperationException) {
            nearestRejected = true;
        }
        // Mislabeled source metadata must not bypass the detected container's policy.
        bool dataUriRejected = !OfficeSvgImageRenderer.TryCreateDataUri("image/bmp", bytes, null, codec, out _);
        Assert.True(nearestRejected && dataUriRejected && codec.Calls == 0,
            $"Nearest rejected: {nearestRejected}; data URI rejected: {dataUriRejected}; caller calls: {codec.Calls}");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnidentifiedMalformedWebpDoesNotReachExportCodecs(bool reservedAlphaHeader) {
        byte[] bytes = ReadFixture("alpha-gradient.webp");
        if (reservedAlphaHeader) {
            for (int cursor = 12; cursor < bytes.Length;) {
                int length = BitConverter.ToInt32(bytes, cursor + 4);
                if (System.Text.Encoding.ASCII.GetString(bytes, cursor, 4) == "ALPH") {
                    bytes[cursor + 8] |= 0x80;
                    break;
                }
                cursor += 8 + length + (length & 1);
            }
        } else {
            Array.Resize(ref bytes, bytes.Length - 1);
        }
        Assert.False(OfficeImageReader.TryIdentifyByContent(bytes, null, out _));
        var codec = new CountingCodec();
        var drawing = new OfficeDrawing(49, 33).AddImageWithInterpolation(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 49, 33)), interpolate: false);
        Assert.Throws<InvalidOperationException>(() => OfficeDrawingSvgExporter.ToSvg(
            drawing, 1D, OfficeSvgSizeUnit.Pixel, imageCodec: codec));
        Assert.False(OfficeSvgImageRenderer.TryCreateDataUri("image/bmp", bytes, null, codec, out _));
        OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions { ImageCodec = codec });
        Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData("alpha-gradient", 0)]
    [InlineData("alpha-gradient", 255)]
    [InlineData("alpha-noise", 0)]
    [InlineData("alpha-noise", 255)]
    public void DrawingCompositesIndependentAlphaOnContrastingBackgrounds(string name, int background) {
        byte[] bytes = ReadFixture(name + ".webp");
        byte[] reference = ReadFixture(name + ".rgba");
        var drawing = new OfficeDrawing(49, 33).AddImage(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 49, 33)));
        var rendered = OfficeDrawingRasterRenderer.Render(drawing,
            background: new OfficeColor((byte)background, (byte)background, (byte)background));
        byte[] actual = rendered.GetPixels();
        for (int i = 0; i < actual.Length; i += 4) {
            Assert.Equal(255, actual[i + 3]);
            for (int channel = 0; channel < 3; channel++) {
                int expected = (reference[i + channel] * reference[i + 3] +
                    background * (255 - reference[i + 3]) + 127) / 255;
                Assert.InRange(Math.Abs(actual[i + channel] - expected), 0, 3);
            }
        }
    }

    private static byte[] ReadFixture(string name) {
        using Stream input = typeof(DrawingWebpAlphaTests).Assembly.GetManifestResourceStream(
            "OfficeIMO.Drawing.Tests.TestAssets.WebpAlpha." + name)!;
        Assert.NotNull(input);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private sealed class CountingCodec : IOfficeRasterImageCodec {
        internal int Calls;
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            image = new OfficeRasterImage(49, 33, OfficeColor.Red);
            return true;
        }
    }
}
