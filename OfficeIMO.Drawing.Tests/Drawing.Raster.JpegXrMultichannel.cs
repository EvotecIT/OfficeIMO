using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingRasterJpegXrMultichannelTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegXr");

    [Fact]
    public void MultichannelSamplesPreserveSourcePrecisionAndAlphaThroughIccConversion() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "nchannel-manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] profileBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus", $"littlecms-{fields[5]}clr-mab.icc"));
            Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
            byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
            Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, out var metadata), fields[0]);
            Assert.Equal(int.Parse(fields[1]), metadata.Width); Assert.Equal(int.Parse(fields[2]), metadata.Height);
            Assert.False(OfficeRasterImageDecoder.TryDecode(encoded, out _));
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, profile!, new OfficeRasterDecodeOptions(), out var image), fields[0]);
            byte[] expected = JpegXrTestFixture.ConvertDeviceReference(
                File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(fields[0], ".nchannel"))), int.Parse(fields[3]), int.Parse(fields[4]), profile!);
            Assert.Equal(expected, image!.GetPixels());
        }
    }

    [Fact]
    public void MultichannelDecodeRejectsMismatchedChannelsAndHonorsLimits() {
        byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, "n8-16-frequency-a2.jxr"));
        string profiles = Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus");
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(profiles, "littlecms-8clr-mab.icc")), out var eight));
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(profiles, "littlecms-7clr-mab.icc")), out var seven));
        Assert.False(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, seven!, new OfficeRasterDecodeOptions(), out _));
        byte[] wrongGuid = { 0x24, 0xC3, 0xDD, 0x6F, 0x03, 0x4E, 0xFE, 0x4B, 0xB1, 0x85, 0x3D, 0x77, 0x76, 0x8D, 0xC9, 0x38 };
        Assert.False(OfficeIccRasterConverter.TryDecodeToSrgb(JpegXrTestFixture.WithField(encoded, 0xBC01, 1, wrongGuid),
            seven!, new OfficeRasterDecodeOptions(), out _));
        Assert.False(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, eight!,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1154 }, out _));
        Assert.False(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, eight!,
            new OfficeRasterDecodeOptions { MaximumEncodedBytes = encoded.Length - 1 }, out _));
        using var source = new System.Threading.CancellationTokenSource(); source.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeIccRasterConverter.TryDecodeToSrgb(encoded, eight!,
            new OfficeRasterDecodeOptions { CancellationToken = source.Token }, out _));
    }
}
