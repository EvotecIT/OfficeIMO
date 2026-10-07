using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffArithmeticLosslessTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegArithmeticLossless");

    [Theory]
    [InlineData(8)]
    [InlineData(12)]
    [InlineData(16)]
    public void NativeContainersPreserveLosslessSamplesInBothByteOrders(int bits) {
        foreach (string line in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] row = line.Split(',');
            if (int.Parse(row[1]) != bits) continue;
            byte[] tiff = File.ReadAllBytes(Path.Combine(Corpus, row[0]));
            Assert.True(OfficeImageReader.TryValidateContent(tiff, row[0], out _), row[0] + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(tiff, out var image), row[0]);
            Assert.Equal((19, 11), (image!.Width, image.Height));
            int components = int.Parse(row[2]), point = int.Parse(row[4]), maximum = (1 << bits) - 1;
            for (int y = 0; y < 11; y++) for (int x = 0; x < 19; x++) {
                byte Expected(int channel) {
                    int sample = (((x * 193 + y * 791 + channel * 3191) ^ (x * y * 53)) & maximum) >> point << point;
                    return (byte)((sample * 255 + maximum / 2) / maximum);
                }
                var pixel = image.GetPixel(x, y);
                Assert.True(pixel.R == Expected(0) && pixel.G == Expected(components == 1 ? 0 : 1) &&
                    pixel.B == Expected(components == 1 ? 0 : 2) && pixel.A == 255, $"{row[0]} at {x},{y}");
            }
        }
    }
}
