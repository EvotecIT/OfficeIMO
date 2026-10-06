using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffLosslessAlphaPrecisionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeAlphaSurvivesUnassociationAndProfileConversion(bool profiled) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLosslessAlphaPrecision");
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(corpus, "..", "IccColorCorpus", "icc-dci-p3-matrix.icc")), out var profile));
        int cases = 0;
        bool testedColorBelowEightBitAlpha = false;
        foreach (string line in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] row = line.Split(',');
            if (profiled && row[1] is not ("2" or "6")) continue;
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, row[0]));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, row[0] + (profiled ? ".icc-rgba" : ".rgba")));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, row[0], out _), row[0]);
            OfficeRasterImage? image;
            Assert.True(profiled ? OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile!, new(), out image) :
                OfficeTiffCodec.TryDecode(bytes, out image), row[0]);
            Assert.Equal((35, 19), (image!.Width, image.Height));
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int at = (y * 35 + x) * 4;
                var pixel = image.GetPixel(x, y);
                Assert.True(pixel.A == expected[at + 3], $"{row[0]} alpha at {x},{y}");
                if (row[4] == "1" && row[1] is "2" or "6" && expected[at + 3] == 0 &&
                    (expected[at] != 0 || expected[at + 1] != 0 || expected[at + 2] != 0)) testedColorBelowEightBitAlpha = true;
                int tolerance = profiled ? 2 : 1;
                Assert.True(Math.Abs(pixel.R - expected[at]) <= tolerance && Math.Abs(pixel.G - expected[at + 1]) <= tolerance &&
                    Math.Abs(pixel.B - expected[at + 2]) <= tolerance, $"{row[0]} color at {x},{y}");
            }
            cases++;
        }
        Assert.Equal(profiled ? 96 : 192, cases);
        Assert.True(testedColorBelowEightBitAlpha);
    }
}
