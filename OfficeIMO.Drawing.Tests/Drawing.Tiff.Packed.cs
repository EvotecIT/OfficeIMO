using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class TiffPackedTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffPacked");

    [Fact]
    public void IndependentPackedGrayAndPaletteSegmentsRetainAllPixels() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
            Assert.True(OfficeTiffCodec.TryValidateAllPages(bytes), fields[0]);
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), fields[0]);
            int bits = int.Parse(fields[1]), photo = int.Parse(fields[2]), mask = (1 << bits) - 1;
            for (int y = 0; y < 17; y++) for (int x = 0; x < 19; x++) {
                int value = (x * 3 + y * 5) & mask;
                var pixel = image!.GetPixel(x, y);
                int red = (photo == 0 ? mask - value : value) * 255 / mask;
                Assert.True(pixel.R == red, $"{fields[0]} {x},{y}: {pixel.R} != {red}");
                Assert.Equal(photo == 3 ? (mask - value) * 255 / mask : red, pixel.G);
                Assert.Equal(photo == 3 ? (value * 7 & mask) * 255 / mask : red, pixel.B);
                Assert.Equal(255, pixel.A);
            }
        }
    }

    [Fact]
    public void BilevelDefaultAndPackedLimitsUseTheSameValidationContract() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "b1-p1-be0-tile0-c1.tif"));
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry) == 258) { bytes[entry] = 0xEA; bytes[entry + 1] = 0xFD; }
        }
        Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image));
        Assert.Equal(255, image!.GetPixel(1, 0).R);
        Assert.True(OfficeTiffCodec.TryValidateAllPages(bytes));
        const int packedBytes = 3 * 17; // Each 19-bit row occupies three bytes.
        Assert.True(OfficeTiffCodec.TryValidateAllPages(bytes, new(), packedBytes * 2));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(bytes, new(), packedBytes * 2 - 1));
        var options = new OfficeRasterDecodeOptions { MaximumDecodedPixels = 19 * 17 - 1 };
        Assert.False(OfficeTiffCodec.TryDecodePage(bytes, 0, options, out _));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(bytes, options));
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry) != 279) continue;
            int counts = BitConverter.ToInt32(bytes, entry + 8);
            // LibTIFF stores the four small strip byte counts as SHORT values.
            bytes[counts]--;
        }
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(bytes));
    }
}
