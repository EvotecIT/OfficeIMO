using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegArithmeticProgressiveTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegArithmeticProgressive");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProgressiveArithmeticPixelsMatchIndependentDecoder(bool fancy) {
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

    [Theory]
    [InlineData(0, 1, 0)]
    [InlineData(0, 0, 14)]
    [InlineData(1, 0, 0)]
    public void InvalidInitialScanCannotProducePixels(int spectralStart, int high, int low) {
        byte[] jpeg = Read("b8-c1-q75-s0-p1-r0.jpg");
        int scan = Segments(jpeg).First(s => s.Marker == 0xDA).Offset;
        int spectral = scan + 5 + 2 * jpeg[scan + 4];
        jpeg[spectral] = (byte)spectralStart; jpeg[spectral + 2] = (byte)((high << 4) | low);
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
        Assert.False(OfficeImageReader.TryValidateContent(jpeg, "source.jpg", out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RefinementMustFollowThePriorApproximation(bool duplicateInitial) {
        byte[] jpeg = Read("b12-c1-q75-s0-p3-r3.jpg");
        int scan = Segments(jpeg).Where(s => s.Marker == 0xDA).Select(s => s.Offset)
            .First(at => jpeg[at + 5 + 2 * jpeg[at + 4] + 2] == 0x32);
        int approximation = scan + 7 + 2 * jpeg[scan + 4];
        jpeg[approximation] = duplicateInitial ? (byte)2 : (byte)0x21;
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
    }

    [Fact]
    public void QuantizationMayArriveBeforeItsComponentFirstScan() {
        byte[] jpeg = Read("b12-c2-q75-s2-p2-r0.jpg");
        var table = Segments(jpeg).Single(s => s.Marker == 0xDB && (jpeg[s.Offset + 4] & 15) == 1);
        int secondDc = Segments(jpeg).Where(s => s.Marker == 0xDA).Skip(1).First().Offset;
        byte[] moved = jpeg.Take(table.Offset).Concat(jpeg.Skip(table.Offset + table.Length).Take(secondDc - table.Offset - table.Length))
            .Concat(jpeg.Skip(table.Offset).Take(table.Length)).Concat(jpeg.Skip(secondDc)).ToArray();
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var expected));
        Assert.True(OfficeJpegCodec.TryDecode(moved, out var actual));
        for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++)
            Assert.Equal(expected!.GetPixel(x, y), actual!.GetPixel(x, y));
    }

    [Fact]
    public void RestartOrderAndPhysicalTerminationAreRequired() {
        byte[] jpeg = Read("b8-c2-q75-s2-p1-r2.jpg");
        int rst = Enumerable.Range(0, jpeg.Length - 1).First(i => jpeg[i] == 255 && jpeg[i + 1] == 0xD0);
        jpeg[rst + 1] = 0xD2;
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
        jpeg = Read("b8-c2-q75-s2-p1-r0.jpg");
        Assert.False(OfficeJpegCodec.TryDecode(jpeg.Take(jpeg.Length - 2).ToArray(), out _));
    }

    [Fact]
    public void EveryComponentNeedsAnInitialDcScan() {
        byte[] jpeg = Read("b8-c1-q75-s0-p2-r0.jpg");
        int nextScan = Segments(jpeg).Where(s => s.Marker == 0xDA).Skip(1).First().Offset;
        byte[] incomplete = jpeg.Take(nextScan).Concat(new byte[] { 255, 217 }).ToArray();
        Assert.False(OfficeJpegCodec.TryDecode(incomplete, out _));
    }

    [Fact]
    public void RetainedProgressiveStateHonorsLimitsAndCancellation() {
        byte[] jpeg = Read("b12-c2-q75-s2-p3-r3.jpg");
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
            out _, out _, out _, out _, cancellationToken: cancellation.Token));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _,
            retainedManagedBytes: OfficeRasterGuards.MaximumDecodedBytes - jpeg.Length));
    }

    private static byte[] Read(string name) => File.ReadAllBytes(Path.Combine(Corpus, name));

    private static IEnumerable<(int Marker, int Offset, int Length)> Segments(byte[] jpeg) {
        for (int at = 2; at + 1 < jpeg.Length;) {
            if (jpeg[at] != 255 || jpeg[at + 1] == 0 || jpeg[at + 1] is >= 0xD0 and <= 0xD7) { at++; continue; }
            int marker = jpeg[at + 1];
            if (marker == 0xD9) yield break;
            int length = 2 + (jpeg[at + 2] << 8 | jpeg[at + 3]);
            yield return (marker, at, length); at += length;
        }
    }
}
