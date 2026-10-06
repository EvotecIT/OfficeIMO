using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegLossless16Tests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLossless16");

    [Theory]
    [InlineData(12)]
    [InlineData(16)]
    public void NativeWordsSurvivePredictionPointTransformsAndByteOrder(int precision) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", $"TiffJpegLossless{precision}");
        int maximum = (1 << precision) - 1;
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(',');
            byte[] jpeg = File.ReadAllBytes(Path.Combine(corpus, f[0] + ".jpg"));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, f[0] + ".jpg.raw"));
            foreach (bool littleEndian in new[] { true, false }) {
                Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
                    out byte[] actual, out int width, out int height, out int channels,
                    preserveRaw16: true, samplesLittleEndian: littleEndian), f[0]);
                Assert.Equal((int.Parse(f[8]), int.Parse(f[9]), int.Parse(f[10])), (width, height, channels));
                Assert.Equal(expected.Length, actual.Length);
                for (int i = 0; i < expected.Length; i += 2) {
                    Assert.Equal(expected[i], actual[i + (littleEndian ? 0 : 1)]);
                    Assert.Equal(expected[i + 1], actual[i + (littleEndian ? 1 : 0)]);
                }
            }
            Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
                out byte[] projected, out _, out _, out int count), f[0]);
            for (int i = 0; i < projected.Length; i++) {
                int sample = expected[i * 2] | expected[i * 2 + 1] << 8;
                Assert.Equal((byte)((sample * 255L + maximum / 2) / maximum), projected[i]);
            }
            if (count == 1) {
                Assert.True(OfficeImageReader.TryValidateContent(jpeg, "source.jpg", out _));
                Assert.True(OfficeRasterContainerInspector.TryInspect(jpeg, out _));
            }
        }
    }

    [Theory]
    [InlineData(12)]
    [InlineData(16)]
    public void TiffKeepsNativeAlphaAndColorUntilRgbaAndIccProjection(int precision) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", $"TiffJpegLossless{precision}");
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(',');
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, f[0]));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, f[0] + ".rgba"));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, f[0], out _), f[0] + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), f[0] + " decode");
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int p = (y * 35 + x) * 4; var pixel = image!.GetPixel(x, y);
                Assert.Equal((expected[p], expected[p + 1], expected[p + 2], expected[p + 3]),
                    (pixel.R, pixel.G, pixel.B, pixel.A));
            }
            if (f[1] is not ("2" or "5" or "6")) continue;
            string profileName = f[1] == "5" ? "littlecms-cmyk-lut.icc" : "littlecms-rgb-matrix.icc";
            byte[] profileBytes = File.ReadAllBytes(Path.Combine(corpus, "..", "IccColorCorpus", profileName));
            Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile!, new(), out var converted));
            byte[] reference = File.ReadAllBytes(Path.Combine(corpus, f[0] + ".reference.tif"));
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(reference, profile!, new(), out var referenceImage));
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                var expectedPixel = referenceImage!.GetPixel(x, y); var actualPixel = converted!.GetPixel(x, y);
                int tolerance = precision == 12 ? 1 : 0; // Reference TIFF rescales native words to sixteen bits.
                Assert.InRange(Math.Abs(expectedPixel.R - actualPixel.R), 0, tolerance);
                Assert.InRange(Math.Abs(expectedPixel.G - actualPixel.G), 0, tolerance);
                Assert.InRange(Math.Abs(expectedPixel.B - actualPixel.B), 0, tolerance);
                Assert.Equal(expectedPixel.A, actualPixel.A);
            }
        }
    }

    [Theory]
    [InlineData(0xC1, 16)]
    [InlineData(0xC3, 1)]
    [InlineData(0xC3, 17)]
    public void UnsupportedProcessOrPrecisionDoesNotProducePixels(int marker, int precision) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "p1-d1-l4-t1-s1-r0-a1.tif.jpg"));
        int frame = Array.FindIndex(jpeg, 1, value => value == 0xC3);
        Assert.True(frame > 0 && jpeg[frame - 1] == 255);
        jpeg[frame] = (byte)marker;
        jpeg[frame + 3] = (byte)precision;
        Assert.False(OfficeImageReader.TryValidateContent(jpeg, "source.jpg", out _));
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
    }

    [Fact]
    public void RawNativeWordsPreserveEightBitSamplesAndEnforceOutputBudget() {
        byte[] eightBit = File.ReadAllBytes(Path.Combine(Corpus, "..", "TiffJpegLossless", "p2-d1-l2-t7-s1-r2.tif.jpg"));
        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(eightBit, 0, false, out byte[] bytes, out _, out _, out _));
        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(eightBit, 0, false, out byte[] words, out _, out _, out _, preserveRaw16: true));
        Assert.Equal(bytes.Length * 2, words.Length);
        for (int i = 0; i < bytes.Length; i++) Assert.Equal((int)bytes[i], words[i * 2] | words[i * 2 + 1] << 8);
        long retained = OfficeRasterGuards.MaximumDecodedBytes - 64L * 1024 - 4L * 1024 * 1024;
        Assert.True(OfficeJpegReader.TryInitializeDecodeWorkingSet(retained, 1024, 1024, 1, out _, 4, 1));
        Assert.False(OfficeJpegReader.TryInitializeDecodeWorkingSet(retained, 1024, 1024, 1, out _, 4, 2));
    }
}
