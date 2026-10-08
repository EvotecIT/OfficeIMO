using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegArithmeticLosslessTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegArithmeticLossless");

    [Theory]
    [InlineData(8, 1)]
    [InlineData(8, 3)]
    [InlineData(12, 1)]
    [InlineData(12, 3)]
    [InlineData(16, 1)]
    [InlineData(16, 3)]
    public void NativeLosslessArithmeticSamplesSurviveRowRestarts(int bits, int components) {
        int max = (1 << bits) - 1;
        foreach (int restart in new[] { 0, 19, 38 }) {
            byte[] jpeg = Read($"b{bits}-c{components}-r{restart}.jpg");
            Assert.True(OfficeImageReader.TryValidateContent(jpeg, "image.jpg", out _));
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image));
            Assert.Equal((19, 11), (image!.Width, image.Height));
            for (int y = 0; y < 11; y++) for (int x = 0; x < 19; x++) {
                var pixel = image.GetPixel(x, y);
                byte Expected(int c) {
                    int value = ((x * 193 + y * 791 + c * 3191) ^ (x * y * 53)) & max;
                    return (byte)((value * 255 + max / 2) / max);
                }
                Assert.Equal(Expected(0), pixel.R);
                Assert.Equal(Expected(components == 1 ? 0 : 1), pixel.G);
                Assert.Equal(Expected(components == 1 ? 0 : 2), pixel.B);
                Assert.Equal(255, pixel.A);
            }
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void EveryPredictorAndPrecisionSupportsPointTransforms(int components) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            if (int.Parse(fields[2]) != components) continue;
            int bits = int.Parse(fields[1]), point = int.Parse(fields[4]), max = (1 << bits) - 1;
            byte[] jpeg = Read(fields[0]);
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image), fields[0]);
            for (int y = 0; y < 11; y++) for (int x = 0; x < 19; x++) {
                var pixel = image!.GetPixel(x, y);
                byte Expected(int c) {
                    int value = (((x * 193 + y * 791 + c * 3191) ^ (x * y * 53)) & max) >> point << point;
                    return (byte)((value * 255 + max / 2) / max);
                }
                Assert.True(pixel.R == Expected(0) && pixel.G == Expected(components == 1 ? 0 : 1) &&
                    pixel.B == Expected(components == 1 ? 0 : 2) && pixel.A == 255,
                    $"{fields[0]} at {x},{y}");
            }
        }
    }

    [Fact]
    public void SubsampledArithmeticMatchesIndependentHuffmanReference() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "subsampled.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] reference = Read(name + ".nearest.rgba");
            Assert.True(OfficeJpegCodec.TryDecode(Read(name), out var image,
                new OfficeJpegDecodeOptions(highQualityChroma: false)), name);
            for (int y = 0; y < 11; y++) for (int x = 0; x < 19; x++) {
                var pixel = image!.GetPixel(x, y); int at = (y * 19 + x) * 4;
                Assert.True(pixel.R == reference[at] && pixel.G == reference[at + 1] &&
                    pixel.B == reference[at + 2] && pixel.A == reference[at + 3], $"{name} at {x},{y}");
            }
        }
    }

    [Theory]
    [InlineData(12, false)]
    [InlineData(12, true)]
    [InlineData(16, false)]
    [InlineData(16, true)]
    public void NativeSampleWordsRetainPrecisionAndByteOrder(int bits, bool littleEndian) {
        foreach (int point in new[] { 0, bits - 1 }) {
            byte[] jpeg = Read($"b{bits}-c3-p7-t{point}-r19.jpg");
            Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out var samples,
                out int width, out int height, out int components, preserveRaw16: true, samplesLittleEndian: littleEndian));
            Assert.Equal((19, 11, 3), (width, height, components));
            Assert.Equal(width * height * components * 2, samples.Length);
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) for (int c = 0; c < components; c++) {
                int expected = (((x * 193 + y * 791 + c * 3191) ^ (x * y * 53)) & ((1 << bits) - 1)) >> point << point;
                int at = ((y * width + x) * components + c) * 2;
                int actual = littleEndian ? samples[at] | samples[at + 1] << 8 : samples[at] << 8 | samples[at + 1];
                Assert.Equal(expected, actual);
            }
        }
    }

    [Theory]
    [InlineData("predictor-zero")]
    [InlineData("predictor-eight")]
    [InlineData("point-precision")]
    [InlineData("successive-approximation")]
    [InlineData("quantization")]
    [InlineData("ac-table")]
    [InlineData("unaligned-restart")]
    public void InvalidLosslessScanParametersCannotProducePixels(string defect) {
        byte[] jpeg = Read("b12-c3-r19.jpg");
        int scan = Marker(jpeg, 0xDA), spectral = scan + 5 + jpeg[scan + 4] * 2;
        switch (defect) {
            case "predictor-zero": jpeg[spectral] = 0; break;
            case "predictor-eight": jpeg[spectral] = 8; break;
            case "point-precision": jpeg[spectral + 2] = 12; break;
            case "successive-approximation": jpeg[spectral + 2] = 0x10; break;
            case "quantization": jpeg[Marker(jpeg, 0xCB) + 12] = 1; break;
            case "ac-table": jpeg[scan + 6] |= 1; break;
            case "unaligned-restart": jpeg[Marker(jpeg, 0xDD) + 5] = 1; break;
        }
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
        Assert.False(OfficeImageReader.TryValidateContent(jpeg, "invalid.jpg", out _));
    }

    private static int Marker(byte[] jpeg, byte marker) =>
        Enumerable.Range(0, jpeg.Length - 1).First(i => jpeg[i] == 255 && jpeg[i + 1] == marker);

    [Fact]
    public void MissingTerminationOrWrongRestartCannotProducePixels() {
        byte[] jpeg = Read("b16-c3-r19.jpg");
        int restart = Enumerable.Range(0, jpeg.Length - 1).First(i => jpeg[i] == 255 && jpeg[i + 1] == 0xD0);
        jpeg[restart + 1] = 0xD2;
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
        jpeg = Read("b16-c3-r0.jpg");
        Assert.False(OfficeJpegCodec.TryDecode(jpeg.Take(jpeg.Length - 2).ToArray(), out _));
    }

    [Fact]
    public void LosslessArithmeticHonorsRetainedMemoryAndCancellation() {
        byte[] jpeg = Read("b16-c3-r0.jpg");
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _,
            retainedManagedBytes: OfficeRasterGuards.MaximumDecodedBytes - jpeg.Length));
        using var cancellation = new System.Threading.CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
            out _, out _, out _, out _, cancellationToken: cancellation.Token));
    }

    private static byte[] Read(string name) => File.ReadAllBytes(Path.Combine(Corpus, name));
}
