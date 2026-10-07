using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingIccMultichannelTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus");

    [Fact]
    public void MultichannelInputLutsMatchIndependentLittleCmsSwatches() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "reference-nchannel.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            double[] input = fields[1].Split(':').Select(v => double.Parse(v) / 65535D).ToArray();
            int[] expected = fields[2].Split(':').Select(int.Parse).ToArray();
            Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(Corpus, fields[0])), out var profile));
            Assert.Equal(input.Length, profile!.ComponentCount);
            Assert.True(profile.TryConvert(input, OfficeIccRenderingIntent.RelativeColorimetric, out var color));
            Assert.InRange(Math.Abs(color.R - expected[0]), 0, 2);
            Assert.InRange(Math.Abs(color.G - expected[1]), 0, 2);
            Assert.InRange(Math.Abs(color.B - expected[2]), 0, 2);
            Assert.False(profile.TryConvertToDevice(color, out _));
            Assert.False(profile.TrySoftProof(color, out _));
            Assert.False(profile.TryConvert(input.Take(input.Length - 1).ToArray(), out _));
            input[input.Length - 1] = double.NaN;
            Assert.False(profile.TryConvert(input, out _));
        }
    }

    [Fact]
    public void PackedMultichannelPixelsUseAllChannelsAndRejectIncompleteBuffers() {
        for (int channels = 3; channels <= 8; channels++) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"littlecms-{channels}clr-mab.icc"));
            Assert.True(OfficeIccColorProfile.TryCreate(bytes, out var profile));
            byte[] samples = Enumerable.Range(0, channels * 2).Select(i => (byte)(17 + i * 13)).ToArray();
            Assert.Equal(OfficeIccRasterConversionStatus.Converted,
                OfficeIccRasterConverter.TryConvertToSrgb(samples, 2, 1, bytes,
                    new OfficeIccRasterConversionOptions { RenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric }, out var image));
            for (int pixel = 0; pixel < 2; pixel++) {
                double[] input = samples.Skip(pixel * channels).Take(channels).Select(v => v / 255D).ToArray();
                Assert.True(profile!.TryConvert(input, OfficeIccRenderingIntent.RelativeColorimetric, out var expected));
                Assert.Equal(expected, image!.GetPixel(pixel, 0));
                Assert.Equal(255, image.GetPixel(pixel, 0).A);
            }
            Assert.Equal(OfficeIccRasterConversionStatus.InvalidSamples,
                OfficeIccRasterConverter.TryConvertToSrgb(samples.Take(samples.Length - 1).ToArray(), 2, 1, bytes, null, out var rejected));
            Assert.Null(rejected);
        }
    }

    [Theory]
    [InlineData("lut8")]
    [InlineData("lut16")]
    [InlineData("mab")]
    public void MultichannelClutRejectsOversizedOrMismatchedDimensions(string kind) {
        byte[] original = File.ReadAllBytes(Path.Combine(Corpus, $"littlecms-8clr-{kind}.icc"));
        int Read32(byte[] b, int p) => (b[p] << 24) | (b[p + 1] << 16) | (b[p + 2] << 8) | b[p + 3];
        int tag = 0;
        for (int i = 0; i < Read32(original, 128); i++) {
            int entry = 132 + i * 12;
            if (Read32(original, entry) == 0x41324230) tag = Read32(original, entry + 4);
        }
        Assert.NotEqual(0, tag);
        byte[] malformed = original.ToArray();
        if (kind == "mab") {
            int grid = tag + Read32(malformed, tag + 24);
            for (int i = 0; i < 8; i++) malformed[grid + i] = 33;
        } else malformed[tag + 10] = 33;
        Assert.False(OfficeIccColorProfile.TryCreate(malformed, out _));
        malformed = original.ToArray(); malformed[tag + 8] = 7;
        Assert.False(OfficeIccColorProfile.TryCreate(malformed, out _));
        malformed = original.ToArray(); malformed[16] = (byte)'9';
        Assert.False(OfficeIccColorProfile.TryCreate(malformed, out _));
    }
}
