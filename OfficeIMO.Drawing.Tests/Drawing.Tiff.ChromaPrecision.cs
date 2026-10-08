using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffChromaPrecisionTests {
    [Fact]
    public void ChromaInterpolationRetainsFractionalSamplesUntilColorProjection() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegChromaPrecision");
        foreach (string line in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] row = line.Split(',');
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, row[0] + ".tif"));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, row[0] + ".rgba"));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, row[0] + ".tif", out _), row[0]);
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), row[0]);
            Assert.Equal((int.Parse(row[1]), int.Parse(row[2])), (image!.Width, image.Height));
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
                int at = (y * image.Width + x) * 4; var p = image.GetPixel(x, y);
                Assert.True(Math.Abs(p.R - expected[at]) <= 1 && Math.Abs(p.G - expected[at + 1]) <= 1 &&
                    Math.Abs(p.B - expected[at + 2]) <= 1 && p.A == 255,
                    $"{row[0]} at {x},{y}: {p.R},{p.G},{p.B} != {expected[at]},{expected[at + 1]},{expected[at + 2]}");
            }
        }
    }
    [Fact]
    public void StandaloneJpegColorRetainsFractionalChromaAcrossPrecisions() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegChromaPrecision");
        for (int precision = 2; precision <= 16; precision++) {
            string name = $"b{precision}-h2-v1-w5-h3-l0";
            byte[] tiff = File.ReadAllBytes(Path.Combine(corpus, name + ".tif"));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, name + ".rgba"));
            // These little-endian strip fixtures contain a complete SOF3 stream.
            int offset = 0, length = 0;
            int entries = tiff[8] | tiff[9] << 8;
            for (int entry = 0; entry < entries; entry++) {
                int at = 10 + entry * 12;
                int tag = tiff[at] | tiff[at + 1] << 8;
                int value = System.Buffers.Binary.BinaryPrimitives.ReadInt32LittleEndian(tiff.AsSpan(at + 8, 4));
                if (tag == 273) offset = value;
                if (tag == 279) length = value;
            }
            Assert.True(offset > 0 && length > 0);
            byte[] jpeg = tiff.AsSpan(offset, length).ToArray();
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image, new OfficeJpegDecodeOptions(highQualityChroma: true)), name);
            for (int y = 0; y < 3; y++) for (int x = 0; x < 5; x++) {
                int at = (y * 5 + x) * 4; var p = image!.GetPixel(x, y);
                Assert.True(Math.Abs(p.R - expected[at]) <= 1 && Math.Abs(p.G - expected[at + 1]) <= 1 &&
                    Math.Abs(p.B - expected[at + 2]) <= 1 && p.A == 255,
                    $"{name} at {x},{y}: {p.R},{p.G},{p.B} != {expected[at]},{expected[at + 1]},{expected[at + 2]}");
            }
        }
    }

}
