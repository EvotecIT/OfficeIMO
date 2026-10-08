using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegArithmeticTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegArithmetic");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SequentialArithmeticPixelsMatchIndependentDecoder(bool fancy) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, name));
            byte[] reference = File.ReadAllBytes(Path.Combine(Corpus, name + (fancy ? ".bilinear.rgba" : ".nearest.rgba")));
            Assert.True(OfficeImageReader.TryValidateContent(jpeg, name, out _), name + " validation");
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image, new OfficeJpegDecodeOptions(highQualityChroma: fancy)), name + " decode");
            Assert.Equal((35, 19), (image!.Width, image.Height));
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
                var pixel = image.GetPixel(x, y); int at = (y * image.Width + x) * 4;
                Assert.True(Math.Abs(pixel.R - reference[at]) <= 2 && Math.Abs(pixel.G - reference[at + 1]) <= 2 &&
                    Math.Abs(pixel.B - reference[at + 2]) <= 2 && pixel.A == 255,
                    $"{name} fancy={fancy} at {x},{y}: actual {pixel.R},{pixel.G},{pixel.B}; reference {reference[at]},{reference[at+1]},{reference[at+2]}");
            }
        }
    }

    [Fact]
    public void ArithmeticJpegStripsPreserveTiffPixels() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] tiff = File.ReadAllBytes(Path.Combine(Corpus, name + ".tif"));
            byte[] reference = File.ReadAllBytes(Path.Combine(Corpus, name + ".bilinear.rgba"));
            Assert.True(OfficeTiffCodec.TryDecode(tiff, out var image), name);
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                var pixel = image!.GetPixel(x, y); int at = (y * 35 + x) * 4;
                Assert.True(Math.Abs(pixel.R - reference[at]) <= 2 && Math.Abs(pixel.G - reference[at + 1]) <= 2 &&
                    Math.Abs(pixel.B - reference[at + 2]) <= 2 && pixel.A == 255, $"{name} at {x},{y}");
            }
        }
    }

    [Fact]
    public void DefaultConditioningWorksWithoutDac() {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "b8-c1-q75-s0-r0.jpg"));
        var bytes = jpeg.ToList();
        foreach (var segment in HeaderSegments(jpeg).Where(s => s.Marker == 0xCC).OrderByDescending(s => s.Offset))
            bytes.RemoveRange(segment.Offset, segment.Length);
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var expected));
        Assert.True(OfficeJpegCodec.TryDecode(bytes.ToArray(), out var actual));
        for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++)
            Assert.Equal(expected!.GetPixel(x, y), actual!.GetPixel(x, y));
    }

    [Theory]
    [InlineData(0, 0x12)] // Lower bound exceeds upper bound.
    [InlineData(16, 64)] // AC conditioning boundary exceeds coefficient range.
    [InlineData(32, 0)] // Undefined conditioning class.
    public void InvalidConditioningIsRejected(int selector, int value) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "b8-c1-q75-s0-r0.jpg"));
        int at = HeaderSegments(jpeg).First(s => s.Marker == 0xCC).Offset + 4;
        jpeg[at] = (byte)selector; jpeg[at + 1] = (byte)value;
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void RestartOrderPresenceAndIntervalAreEnforced(int corruption) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "b12-c2-q75-s2-r2.jpg"));
        int rst = Enumerable.Range(0, jpeg.Length - 1).First(i => jpeg[i] == 255 && jpeg[i + 1] == 0xD0);
        if (corruption == 0) jpeg[rst + 1] = 0xD1;
        else if (corruption == 1) jpeg[rst + 1] = 0;
        else {
            int dri = HeaderSegments(jpeg).Single(s => s.Marker == 0xDD).Offset;
            jpeg[dri + 4] = jpeg[dri + 5] = 0;
        }
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
    }

    [Theory]
    [InlineData(0xC9, 9)]
    [InlineData(0xC9, 16)]
    [InlineData(0xCA, 9)]
    [InlineData(0xCB, 17)]
    public void UnsupportedArithmeticProcessesAndPrecisionsAreRejected(int marker, int precision) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "b8-c1-q75-s0-r0.jpg"));
        int frame = HeaderSegments(jpeg).Single(s => s.Marker == 0xC9).Offset;
        jpeg[frame + 1] = (byte)marker; jpeg[frame + 4] = (byte)precision;
        Assert.False(OfficeImageReader.TryValidateContent(jpeg, "source.jpg", out _));
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
    }

    [Fact]
    public void ArithmeticPlanesHonorCancellationAndRetentionLimits() {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "b12-c2-q75-s2-r2.jpg"));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
            out _, out _, out _, out _, cancellationToken: cancellation.Token));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _,
            retainedManagedBytes: OfficeRasterGuards.MaximumDecodedBytes - jpeg.Length));
    }

    [Fact]
    public void StrictArithmeticRequiresRealTerminationButAllowsZeroLengthEntropy() {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "b8-c0-q75-s0-r0.jpg"));
        int frame = HeaderSegments(jpeg).Single(s => s.Marker == 0xC9).Offset;
        jpeg[frame + 5] = jpeg[frame + 7] = 0;
        jpeg[frame + 6] = jpeg[frame + 8] = 1;
        int scan = HeaderSegments(jpeg).Last().Offset + HeaderSegments(jpeg).Last().Length;
        int entropy = scan + 2 + (jpeg[scan + 2] << 8 | jpeg[scan + 3]);
        byte[] emptyScan = jpeg.Take(entropy).ToArray();
        Assert.False(OfficeJpegCodec.TryDecode(emptyScan, out _));
        Assert.True(OfficeJpegCodec.TryDecode(emptyScan, out _, new OfficeJpegDecodeOptions(allowTruncated: true)));
        byte[] terminated = emptyScan.Concat(new byte[] { 255, 217 }).ToArray();
        Assert.True(OfficeJpegCodec.TryDecode(terminated, out var image));
        Assert.Equal((1, 1), (image!.Width, image.Height));
        Assert.Equal(128, image.GetPixel(0, 0).R);
        byte[] commentWithoutEoi = emptyScan.Concat(new byte[] { 255, 254, 0, 2 }).ToArray();
        Assert.False(OfficeJpegCodec.TryDecode(commentWithoutEoi, out _));
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
