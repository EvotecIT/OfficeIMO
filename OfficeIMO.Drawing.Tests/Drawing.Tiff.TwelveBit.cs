using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTiffTwelveBitTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Tiff12");

    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    [InlineData(8)]
    [InlineData(32773)]
    [InlineData(7)]
    public void IndependentTwelveBitFilesPreserveSamplesAndAlpha(int compression) {
        string[] files = Directory.GetFiles(Corpus, $"*-c{compression}-*.tif");
        Assert.NotEmpty(files);
        foreach (string file in files) {
            byte[] encoded = File.ReadAllBytes(file);
            Assert.True(OfficeImageReader.TryValidateContent(encoded, "source.tif", out _), file);
            Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var image), file);
            Assert.Equal((35, 19), (image!.Width, image.Height));
            int kind = int.Parse(Path.GetFileName(file).Substring(1, 1));
            int channels = kind < 2 ? 1 : kind == 3 || kind == 4 ? 4 : 3;
            byte[] raw = File.ReadAllBytes(file + ".raw");
            Assert.Equal(35 * 19 * channels * 2, raw.Length);
            int Sample(int at) => raw[at * 2] | raw[at * 2 + 1] << 8;
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int p = (y * 35 + x) * channels, alpha = channels == 4 ? Sample(p + 3) : 4095;
                int Component(int c) {
                    double value = kind == 3 && alpha == 0 ? 0 : Math.Min(1D, Sample(p + (channels == 1 ? 0 : c)) / (double)(kind == 3 ? alpha : 4095));
                    if (kind == 0) value = 1 - value;
                    return (int)Math.Floor(value * 255 + .5);
                }
                var actual = image.GetPixel(x, y);
                int tolerance = compression == 7 ? 3 : 0;
                Assert.True(Math.Abs(actual.R - Component(0)) <= tolerance && Math.Abs(actual.G - Component(1)) <= tolerance && Math.Abs(actual.B - Component(2)) <= tolerance,
                    $"{Path.GetFileName(file)} {x},{y}: {actual.R},{actual.G},{actual.B}; expected {Component(0)},{Component(1)},{Component(2)}");
                Assert.Equal((byte)((alpha * 255 + 2047) / 4095), actual.A);
            }
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(7)]
    public void TwelveBitDecodingHonorsResourceLimitsAndCancellation(int compression) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"k2-c{compression}-be0-p1-t0.tif"));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - bytes.Length }, out _, out _));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TwelveBitPayloadAndContainerPrecisionMustAgree(bool jpeg) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"k1-c{(jpeg ? 7 : 1)}-be0-p1-t0.tif"));
        SetShortTag(bytes, 258, 16);
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
    }

    [Fact]
    public void TwelveBitSelfContainedLegacyJpegUsesTheSamePrecisionContract() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "k2-c7-be0-p1-t0.tif"));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var modern));
        SetShortTag(bytes, 259, 6);
        Assert.True(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var legacy));
        for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++)
            Assert.Equal(modern!.GetPixel(x, y), legacy!.GetPixel(x, y));
    }

    private static void SetShortTag(byte[] bytes, int tag, int value) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int at = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, at) != tag) continue;
            Assert.Equal(3, BitConverter.ToUInt16(bytes, at + 2));
            Assert.Equal(1, BitConverter.ToInt32(bytes, at + 4));
            bytes[at + 8] = (byte)value; bytes[at + 9] = (byte)(value >> 8);
            return;
        }
        Assert.Fail($"Missing TIFF tag {tag}.");
    }

}
