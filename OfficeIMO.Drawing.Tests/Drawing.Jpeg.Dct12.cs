using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegDct12Tests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegDct12");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SequentialAndProgressiveTwelveBitPixelsMatchIndependentDecoder(bool fancy) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, name));
            byte[] reference = File.ReadAllBytes(Path.Combine(Corpus, name + (fancy ? ".bilinear.rgba" : ".nearest.rgba")));
            Assert.True(OfficeImageReader.TryValidateContent(jpeg, name, out _), name + " validation");
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image, new OfficeJpegDecodeOptions(highQualityChroma: fancy)), name + " decode");
            Assert.Equal((35, 19), (image!.Width, image.Height));
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
                var pixel = image.GetPixel(x, y); int at = (y * image.Width + x) * 4;
                Assert.True(Math.Abs(pixel.R - reference[at]) <= 1 && Math.Abs(pixel.G - reference[at + 1]) <= 1 &&
                    Math.Abs(pixel.B - reference[at + 2]) <= 1 && pixel.A == 255,
                    $"{name} fancy={fancy} at {x},{y}: actual {pixel.R},{pixel.G},{pixel.B}; reference {reference[at]},{reference[at+1]},{reference[at+2]}");
            }
        }
    }

    [Theory]
    [InlineData(0xC0, 12)]
    [InlineData(0xC1, 9)]
    [InlineData(0xC1, 16)]
    [InlineData(0xC2, 9)]
    [InlineData(0xC2, 16)]
    public void DctFrameMarkerAndPrecisionMustBeCompatible(int marker, int precision) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "c1-q75-p0-s0-r0.jpg"));
        int frame = HeaderSegments(jpeg).Single(s => s.Marker == 0xC1).Offset;
        jpeg[frame + 1] = (byte)marker;
        jpeg[frame + 4] = (byte)precision;
        Assert.False(OfficeImageReader.TryValidateContent(jpeg, "source.jpg", out _));
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
    }

    [Fact]
    public void TwelveBitDctRequiresExplicitHuffmanTables() {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "c1-q75-p0-s0-r0.jpg"));
        var ranges = HeaderSegments(jpeg).Where(s => s.Marker == 0xC4).OrderByDescending(s => s.Offset).ToArray();
        var bytes = jpeg.ToList();
        foreach (var range in ranges) bytes.RemoveRange(range.Offset, range.Length);
        Assert.NotEmpty(ranges);
        Assert.False(OfficeJpegCodec.TryDecode(bytes.ToArray(), out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes.ToArray(), "source.jpg", out _));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void NativeWidthDctPlanesHonorCancellationAndRetentionLimits(int progressive) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, $"c2-q75-p{progressive}-s2-r2.jpg"));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
            out _, out _, out _, out _, cancellationToken: cancellation.Token));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _,
            retainedManagedBytes: OfficeRasterGuards.MaximumDecodedBytes - jpeg.Length));
    }

    private static IEnumerable<(int Marker, int Offset, int Length)> HeaderSegments(byte[] jpeg) {
        int at = 2;
        while (at + 4 < jpeg.Length) {
            int marker = jpeg[at + 1], length = (jpeg[at + 2] << 8 | jpeg[at + 3]) + 2;
            if (marker == 0xDA || marker == 0xD9) yield break;
            yield return (marker, at, length);
            at += length;
        }
    }
}
