using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegLosslessTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLossless");

    [Fact]
    public void IndependentLosslessPredictorsPreserveComponentsAndTiffAlpha() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(',');
            byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, f[0] + ".jpg"));
            Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
                out byte[] components, out int jw, out int jh, out int jc), f[0]);
            Assert.Equal((int.Parse(f[8]), int.Parse(f[9]), int.Parse(f[10])), (jw, jh, jc));
            Assert.Equal(File.ReadAllBytes(Path.Combine(Corpus, f[0] + ".jpg.raw")), components);
            if (jc == 1) {
                Assert.True(OfficeImageReader.TryValidateContent(jpeg, "source.jpg", out _));
                Assert.True(OfficeRasterContainerInspector.TryInspect(jpeg, out _));
                Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var gray));
                for (int y = 0; y < jh; y++) for (int x = 0; x < jw; x++)
                    Assert.Equal(components[y * jw + x], gray!.GetPixel(x, y).R);
            }
            byte[] tiff = File.ReadAllBytes(Path.Combine(Corpus, f[0]));
            Assert.True(OfficeImageReader.TryValidateContent(tiff, f[0], out _), f[0]);
            Assert.True(OfficeTiffCodec.TryDecode(tiff, out var image), f[0]);
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, f[0] + ".rgba"));
            Assert.Equal((35, 19), (image!.Width, image.Height));
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int p = (y * 35 + x) * 4; var actual = image.GetPixel(x, y);
                Assert.Equal((expected[p], expected[p + 1], expected[p + 2], expected[p + 3]),
                    (actual.R, actual.G, actual.B, actual.A));
            }
        }
    }

    [Fact]
    public void ExtendedSequentialImagesUseTheSameValidatedContainerPath() {
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(9, 9, OfficeColor.SteelBlue));
        jpeg[FindMarker(jpeg, 192) + 1] = 193;
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out _));
        Assert.True(OfficeRasterContainerInspector.TryInspect(jpeg, out var container));
        Assert.Equal((9, 9), (container!.CanvasWidth, container.CanvasHeight));
    }

    [Fact]
    public void LosslessDecodingHonorsCancellationAndRetainedMemory() {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "p2-d1-l2-t7-s1-r2.tif.jpg"));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _,
            retainedManagedBytes: OfficeRasterGuards.MaximumDecodedBytes));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => { OfficeJpegCodec.TryDecodeColorComponents(
            jpeg, 0, false, out _, out _, out _, out _, cancellationToken: canceled.Token); });
    }

    [Theory]
    [InlineData("TiffJpegLossless")]
    [InlineData("TiffJpegLossless16")]
    public void SubsampledLosslessEdgesRetainConstantSamplesThroughHighQualityInterpolation(string folder) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", folder, "edges");
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(','); int width = int.Parse(f[1]), height = int.Parse(f[2]);
            byte[] jpeg = File.ReadAllBytes(Path.Combine(corpus, f[0] + ".jpg"));
            foreach (bool highQuality in new[] { false, true }) {
                Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image, new OfficeJpegDecodeOptions(highQualityChroma: highQuality)));
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    var pixel = image!.GetPixel(x, y);
                    Assert.Equal((128, 128, 128), ((int)pixel.R, (int)pixel.G, (int)pixel.B));
                }
            }
            // TIFF permits vertical subsampling no greater than horizontal.
            if (int.Parse(f[4]) > int.Parse(f[3])) continue;
            byte[] tiff = File.ReadAllBytes(Path.Combine(corpus, f[0] + ".tif"));
            Assert.True(OfficeTiffCodec.TryDecode(tiff, out var tiffImage), f[0]);
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                var pixel = tiffImage!.GetPixel(x, y);
                Assert.Equal((128, 128, 128), ((int)pixel.R, (int)pixel.G, (int)pixel.B));
            }
        }
    }

    [Theory]
    [InlineData("predictor")]
    [InlineData("spectral")]
    [InlineData("successive")]
    [InlineData("point")]
    [InlineData("quantization")]
    [InlineData("ac-table")]
    [InlineData("restart-row")]
    [InlineData("restart-sequence")]
    [InlineData("missing-table")]
    [InlineData("truncated")]
    public void InvalidLosslessParametersAndEntropyFailWithoutPartialPixels(string mutation) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, "p2-d1-l2-t7-s1-r2.tif.jpg"));
        int frame = FindMarker(jpeg, 195), scan = FindMarker(jpeg, 218), dri = FindMarker(jpeg, 221);
        int count = jpeg[scan + 4], parameters = scan + 5 + count * 2;
        switch (mutation) {
            case "predictor": jpeg[parameters] = 0; break;
            case "spectral": jpeg[parameters + 1] = 1; break;
            case "successive": jpeg[parameters + 2] = 16; break;
            case "point": jpeg[parameters + 2] = 8; break;
            case "quantization": jpeg[frame + 12] = 1; break;
            case "ac-table": jpeg[scan + 6] |= 1; break;
            case "restart-row": jpeg[dri + 5] = 1; break;
            case "restart-sequence": jpeg[FindMarker(jpeg, 208) + 1] = 211; break;
            case "missing-table": jpeg[FindMarker(jpeg, 196) + 1] = 254; break;
            case "truncated": jpeg = jpeg.Take(scan + 12).Concat(new byte[] { 255, 217 }).ToArray(); break;
        }
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _));
    }

    private static int FindMarker(byte[] data, byte marker) {
        for (int i = 2; i + 1 < data.Length; i++) if (data[i] == 255 && data[i + 1] == marker) return i;
        throw new InvalidOperationException($"Missing marker {marker}");
    }
}
