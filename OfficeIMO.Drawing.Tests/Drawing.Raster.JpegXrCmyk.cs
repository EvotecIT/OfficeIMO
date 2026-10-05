using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingRasterJpegXrCmykTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegXr");

    [Fact]
    public void CmykAndCmykDirectPreserveSourcePrecisionAndAlphaThroughIccConversion() {
        byte[] profileBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "cmyk-manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
            Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, out var metadata), fields[0]);
            Assert.Equal(int.Parse(fields[1]), metadata.Width); Assert.Equal(int.Parse(fields[2]), metadata.Height);
            Assert.False(OfficeRasterImageDecoder.TryDecode(encoded, out _));
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, profile!, new OfficeRasterDecodeOptions(), out var image), fields[0]);
            byte[] expected = JpegXrTestFixture.ConvertCmykReference(
                File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(fields[0], ".cmyk"))), int.Parse(fields[3]), int.Parse(fields[4]), profile!);
            Assert.Equal(expected, image!.GetPixels());
        }
    }

    [Fact]
    public void CmykConversionRejectsWrongProfileAndHonorsResourceLimits() {
        byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, "cmyk16-direct-frequency-q32-a2.jxr"));
        string profiles = Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus");
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(profiles, "littlecms-rgb-matrix.icc")), out var rgb));
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(profiles, "littlecms-cmyk-lut.icc")), out var cmyk));
        Assert.False(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, rgb!, new OfficeRasterDecodeOptions(), out _));
        Assert.False(OfficeIccRasterConverter.TryDecodeToSrgb(encoded, cmyk!, new OfficeRasterDecodeOptions { MaximumDecodedPixels = 246 }, out _));
        using var source = new System.Threading.CancellationTokenSource(); source.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeIccRasterConverter.TryDecodeToSrgb(encoded, cmyk!,
            new OfficeRasterDecodeOptions { CancellationToken = source.Token }, out _));
    }
}
