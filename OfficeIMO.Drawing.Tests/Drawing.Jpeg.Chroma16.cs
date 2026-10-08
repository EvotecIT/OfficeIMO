using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegChroma16Tests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegChroma16");

    [Fact]
    public void NativeSampleReconstructionMatchesIndependentDecodeAndInterpolation() {
        foreach (string file in Directory.GetFiles(Corpus, "*.jpg")) {
            byte[] jpeg = File.ReadAllBytes(file);
            foreach (bool highQuality in new[] { false, true }) {
                byte[] expected = File.ReadAllBytes(Path.ChangeExtension(file, highQuality ? ".bilinear.raw" : ".nearest.raw"));
                Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
                    out byte[] actual, out _, out _, out _,
                    options: new OfficeJpegDecodeOptions(highQualityChroma: highQuality), preserveRaw16: true), file);
                Assert.Equal(expected.Length, actual.Length);
                for (int i = 0; i < expected.Length; i += 2) {
                    int reference = expected[i] | expected[i + 1] << 8;
                    int decoded = actual[i] | actual[i + 1] << 8;
                    Assert.True(reference == decoded, $"{Path.GetFileName(file)} highQuality={highQuality} sample={i / 2}: {reference} != {decoded}");
                }
            }
        }
    }

    [Fact]
    public void TiffChromaMatchesReferenceAcrossPositioningPlanesAndPartialTiles() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(','); string name = fields[0];
            int width = int.Parse(fields[1]), height = int.Parse(fields[2]);
            byte[] tiff = File.ReadAllBytes(Path.Combine(Corpus, name + ".tif"));
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, name + ".rgba"));
            Assert.True(OfficeImageReader.TryValidateContent(tiff, name + ".tif", out _), name);
            Assert.True(OfficeTiffCodec.TryDecode(tiff, out var image), name);
            Assert.Equal((width, height), (image!.Width, image.Height));
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int p = (y * width + x) * 4; var pixel = image.GetPixel(x, y);
                Assert.True((expected[p], expected[p + 1], expected[p + 2], expected[p + 3]) ==
                    (pixel.R, pixel.G, pixel.B, pixel.A), $"{name} at {x},{y}: expected {expected[p]},{expected[p + 1]},{expected[p + 2]}; actual {pixel.R},{pixel.G},{pixel.B}");
            }
        }
    }
}
