using OfficeIMO.Drawing;
using System.Threading;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingRasterJpegXrTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegXr");
    private static byte[] Fixture(string name = "rgb-19x13-frequency-overlap2-alpha0") => File.ReadAllBytes(Path.Combine(Corpus, name + ".jxr"));

    [Fact]
    public void IndependentReferenceImagesDecodeThroughThePublicApi() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
            Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var metadata), fields[0]);
            Assert.Equal(OfficeImageFormat.JpegXr, metadata.Format);
            Assert.Equal(int.Parse(fields[1]), metadata.Width); Assert.Equal(int.Parse(fields[2]), metadata.Height);
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, null, out var image, out var info), fields[0] + ": " + info.Diagnostic);
            Assert.Equal(1, info.FrameCount); Assert.False(info.FramesOrPagesDiscarded);
            Assert.Equal(File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(fields[0], ".rgba"))), image!.GetPixels());
        }
    }

    [Fact]
    public void SubsampledFirstLevelOverlapRejectsAnUndersizedImage() {
        byte[] bytes = Fixture("subsampled-yuv420-tiny-frequency-overlap0-q32-alpha0");
        int directory = Read32(bytes, 4), count = bytes[directory] | bytes[directory + 1] << 8;
        for (int i = 0; i < count; i++) {
            int entry = directory + 2 + i * 12;
            if ((bytes[entry] | bytes[entry + 1] << 8) != 0xBCC0) continue;
            int frame = Read32(bytes, entry + 8);
            bytes[frame + 9] = (byte)((bytes[frame + 9] & 0xFC) | 2);
        }
        Assert.False(OfficeImageReader.TryIdentifyByContent(bytes, null, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    [Theory]
    [InlineData("rgba-premultiplied", false)]
    [InlineData("rgba-premultiplied-separate", false)]
    [InlineData("rgba-premultiplied-separate", true)]
    public void ContainerAssociationControlsLegacyAndSeparateAlpha(string name, bool setAlphaFlag) {
        byte[] bytes = Fixture(name);
        int directory = Read32(bytes, 4), count = bytes[directory] | bytes[directory + 1] << 8;
        int primary = 0, alpha = 0, guid = 0;
        for (int i = 0; i < count; i++) {
            int entry = directory + 2 + i * 12;
            int tag = bytes[entry] | bytes[entry + 1] << 8;
            if (tag == 0xBCC0) primary = Read32(bytes, entry + 8);
            if (tag == 0xBCC2) alpha = Read32(bytes, entry + 8);
            if (tag == 0xBC01) guid = Read32(bytes, entry + 8);
        }
        bytes[primary + 10] &= 0xFD;
        if (alpha != 0) bytes[alpha + 10] = (byte)((bytes[alpha + 10] & 0xFD) | (setAlphaFlag ? 2 : 0));
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out _));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
        Assert.Equal(File.ReadAllBytes(Path.Combine(Corpus, name + ".rgba")), image!.GetPixels());
        if (alpha != 0) {
            // The primary flag is ignored when alpha occupies a separate stream.
            bytes[primary + 10] |= 2;
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out image));
            Assert.Equal(File.ReadAllBytes(Path.Combine(Corpus, name + ".rgba")), image!.GetPixels());
        }
        // A positive flag in the alpha-bearing stream contradicts straight BGRA.
        bytes[guid + 15] = 0x0F;
        bytes[(alpha == 0 ? primary : alpha) + 10] |= 2;
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    [Theory]
    [InlineData(".jxr", "image/jxr")]
    [InlineData(".wdp", "image/vnd.ms-photo")]
    [InlineData(".hdp", "image/jxr")]
    public void FormatAliasesIdentifyTheSameManagedCodec(string extension, string mime) {
        Assert.Equal(OfficeImageFormat.JpegXr, OfficeImageReader.FromExtension(extension));
        Assert.Equal(OfficeImageFormat.JpegXr, OfficeImageInfo.FromMimeType(mime));
        Assert.False(OfficeImageInfo.IsBrowserPreviewSafeContentType(mime));
    }

    [Fact]
    public void PublicLimitsAndCancellationApplyToByteAndStreamDecoding() {
        byte[] bytes = Fixture();
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { MaximumDecodedPixels = 246 }, out _, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { MaximumEncodedBytes = bytes.Length - 1 }, out _, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { FrameIndex = 1 }, out _, out _));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));
        using var stream = new MemoryStream(new byte[] { 99 }.Concat(bytes).ToArray()); stream.Position = 1;
        Assert.True(OfficeRasterImageDecoder.TryDecode(stream, out var image));
        Assert.Equal(1, stream.Position); Assert.Equal(19, image!.Width);
    }

    [Theory]
    [InlineData("rgb-19x13-frequency-overlap2-alpha0", 0x15)]
    [InlineData("u16-3c-spatial-overlap0-q0-alpha0", 0x0D)]
    [InlineData("u16-1c-spatial-overlap1-q0-alpha0", 0x08)]
    public void ContainerAndCodestreamSampleDepthMustAgree(string name, int format) {
        byte[] guid = { 0x24, 0xC3, 0xDD, 0x6F, 3, 0x4E, 0xFE, 0x4B, 0xB1, 0x85, 0x3D, 0x77, 0x76, 0x8D, 0xC9, (byte)format };
        byte[] bytes = OfficeIMO.TestAssets.JpegXrTestFixture.WithField(Fixture(name), 0xBC01, 1, guid);
        Assert.False(OfficeImageReader.TryIdentifyByContent(bytes, null, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    [Theory]
    [InlineData("extended-f16-infinity")]
    [InlineData("extended-f16-nan")]
    [InlineData("extended-f32-infinity")]
    [InlineData("extended-f32-nan")]
    public void NonfiniteSamplesFailWithoutReturningPartialPixels(string name) {
        Assert.True(OfficeImageReader.TryIdentifyByContent(Fixture(name), null, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(Fixture(name), out var image));
        Assert.Null(image);
    }

    [Fact]
    public void TruncatedOrCorruptPayloadFailsValidation() {
        byte[] bytes = Fixture();
        foreach (int removed in new[] { 1, 8, bytes.Length / 2 }) {
            byte[] truncated = bytes.Take(bytes.Length - removed).ToArray();
            Assert.False(OfficeRasterImageDecoder.TryDecode(truncated, out _));
            Assert.False(OfficeImageReader.TryValidateContent(truncated, null, out _));
        }
        int directory = Read32(bytes, 4), count = bytes[directory] | bytes[directory + 1] << 8;
        int imageOffset = 0;
        for (int i = 0; i < count; i++) {
            int entry = directory + 2 + i * 12;
            if ((bytes[entry] | bytes[entry + 1] << 8) == 0xBCC0) imageOffset = Read32(bytes, entry + 8);
        }
        Assert.True(imageOffset > 0);
        bytes[imageOffset] = 0;
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    [Fact]
    public void ContainerOrientationPermutesPixelsAndDimensions() {
        byte[] source = Fixture();
        Assert.True(OfficeRasterImageDecoder.TryDecode(source, out var original));
        for (int transform = 0; transform < 8; transform++) {
            byte[] bytes = OfficeIMO.TestAssets.JpegXrTestFixture.WithField(source, 0xBC02, 1, new[] { (byte)transform });
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
            Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var info));
            Assert.Equal(transform < 4 ? 19 : 13, image!.Width);
            Assert.Equal(transform < 4 ? 13 : 19, image.Height);
            Assert.Equal(info.Width, image.Width); Assert.Equal(info.Height, image.Height);
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
                int tx = (transform & 2) != 0 ? image.Width - 1 - x : x;
                int ty = (transform & 1) != 0 ? image.Height - 1 - y : y;
                int sx = transform < 4 ? tx : ty, sy = transform < 4 ? ty : original!.Height - 1 - tx;
                Assert.Equal(original!.GetPixel(sx, sy), image.GetPixel(x, y));
            }
        }
    }

    [Fact]
    public void AggregateCoefficientBudgetRejectsOversizedWorkingSetsBeforeDecoding() {
        byte[] bytes = Fixture();
        int directory = Read32(bytes, 4), count = bytes[directory] | bytes[directory + 1] << 8, imageOffset = 0;
        for (int i = 0; i < count; i++) {
            int entry = directory + 2 + i * 12;
            if ((bytes[entry] | bytes[entry + 1] << 8) == 0xBCC0) imageOffset = Read32(bytes, entry + 8);
        }
        Assert.NotEqual(0, bytes[imageOffset + 10] & 0x80);
        bytes[imageOffset + 12] = 0x10; bytes[imageOffset + 13] = 0x02; // 4099 columns
        bytes[imageOffset + 14] = 0x0F; bytes[imageOffset + 15] = 0xFC; // 4093 rows
        bytes = OfficeIMO.TestAssets.JpegXrTestFixture.WithField(bytes, 0xBC80, 4, new byte[] { 3, 16, 0, 0 });
        bytes = OfficeIMO.TestAssets.JpegXrTestFixture.WithField(bytes, 0xBC81, 4, new byte[] { 253, 15, 0, 0 });
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var info));
        Assert.Equal(4099, info.Width); Assert.Equal(4093, info.Height);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    private static int Read32(byte[] bytes, int offset) => bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24;
}
