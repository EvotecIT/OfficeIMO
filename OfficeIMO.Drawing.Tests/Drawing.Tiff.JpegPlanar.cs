using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffJpegPlanarTests {
    [Theory]
    [InlineData("TiffJpegPlanar")]
    [InlineData("TiffJpegCosited")]
    public void IndependentlyDecodedChromaRetainsEdgesAndPositioning(string folder) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            string name = fields[0];
            int width = folder == "TiffJpegCosited" ? int.Parse(fields[7]) : 67;
            int height = folder == "TiffJpegCosited" ? int.Parse(fields[8]) : 35;
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, name));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, name + ".rgb"));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _), name + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), name + " decode");
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                var actual = image!.GetPixel(x, y); int offset = (y * width + x) * 3;
                int error = Math.Max(Math.Abs(actual.R - expected[offset]),
                    Math.Max(Math.Abs(actual.G - expected[offset + 1]), Math.Abs(actual.B - expected[offset + 2])));
                Assert.True(error <= 3, $"{name} {x},{y}: delta {error}");
                Assert.Equal(255, actual.A);
            }
        }
    }

    [Theory]
    [InlineData(221)]
    [InlineData(204)]
    public void SharedTableControlMarkersDoNotLeakIntoSegments(byte marker) {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpeg", "p2-be0-t0-q3-pl1-s1.tif"));
        int ifd = BitConverter.ToInt32(original, 4), entry = 0;
        for (int i = 0; i < BitConverter.ToUInt16(original, ifd); i++) {
            int p = ifd + 2 + 12 * i;
            if (BitConverter.ToUInt16(original, p) == 347) entry = p;
        }
        Assert.NotEqual(0, entry);
        int length = BitConverter.ToInt32(original, entry + 4), table = BitConverter.ToInt32(original, entry + 8);
        byte[] mutated = new byte[original.Length + length + 6];
        original.CopyTo(mutated, 0);
        Array.Copy(original, table, mutated, original.Length, length - 2);
        new byte[] { 255, marker, 0, 4, 0, marker == 204 ? (byte)16 : (byte)1, 255, 217 }.CopyTo(mutated, original.Length + length - 2);
        BitConverter.GetBytes(length + 6).CopyTo(mutated, entry + 4);
        BitConverter.GetBytes(original.Length).CopyTo(mutated, entry + 8);
        Assert.True(OfficeTiffCodec.TryDecode(original, out var expected));
        Assert.True(OfficeImageReader.TryValidateContent(mutated, "tables.tif", out _));
        Assert.True(OfficeTiffCodec.TryDecode(mutated, out var actual));
        for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++)
            Assert.Equal(expected!.GetPixel(x, y), actual!.GetPixel(x, y));
    }
}
