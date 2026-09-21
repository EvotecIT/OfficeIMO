using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingIccRasterColorCorpusTests {
    [Fact]
    public void IndependentMatrixRgbLutRgbAndLutCmykSwatchesMatchSrgbReferences() {
        string directory = Path.Combine(AppContext.BaseDirectory, "TestAssets", "IccColorCorpus");
        string[] rows = File.ReadAllLines(Path.Combine(directory, "reference-srgb.csv"));
        Assert.Equal(25, rows.Length);
        var seenProfiles = new HashSet<string>(StringComparer.Ordinal);
        foreach (string row in rows.Skip(1)) {
            string[] fields = row.Split(',');
            Assert.Equal(3, fields.Length);
            byte[] source = ParseChannels(fields[1]);
            byte[] expected = ParseChannels(fields[2]);
            seenProfiles.Add(fields[0]);
            byte[] profile = File.ReadAllBytes(Path.Combine(directory, fields[0]));

            OfficeIccRasterConversionStatus status = OfficeIccRasterConverter.TryConvertToSrgb(
                source, 1, 1, profile, options: null, out OfficeRasterImage? image);

            Assert.Equal(OfficeIccRasterConversionStatus.Converted, status);
            Assert.NotNull(image);
            OfficeColor pixel = image.GetPixel(0, 0);
            Assert.InRange(Math.Abs(pixel.R - expected[0]), 0, 2);
            Assert.InRange(Math.Abs(pixel.G - expected[1]), 0, 2);
            Assert.InRange(Math.Abs(pixel.B - expected[2]), 0, 2);
            Assert.Equal(255, pixel.A);
        }
        Assert.Equal(4, seenProfiles.Count);
    }

    [Fact]
    public void ConverterRejectsMalformedProfilesAndDeclaredResourceLimitsBeforePixelAllocation() {
        byte[] profile = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory, "TestAssets", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
        byte[] sample = { 64, 128, 192, 26 };

        Assert.Equal(OfficeIccRasterConversionStatus.UnsupportedProfile,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1,
                profile.Take(100).ToArray(), null, out OfficeRasterImage? malformed));
        Assert.Null(malformed);
        byte[] damagedTag = profile.ToArray();
        damagedTag[136] = 0x7f;
        damagedTag[137] = 0xff;
        damagedTag[138] = 0xff;
        damagedTag[139] = 0xff;
        Assert.Equal(OfficeIccRasterConversionStatus.UnsupportedProfile,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1,
                damagedTag, null, out OfficeRasterImage? invalidTag));
        Assert.Null(invalidTag);
        Assert.Equal(OfficeIccRasterConversionStatus.ProfileLimitExceeded,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1, profile,
                new OfficeIccRasterConversionOptions { MaximumProfileBytes = 1024 }, out OfficeRasterImage? oversized));
        Assert.Null(oversized);
        Assert.Equal(OfficeIccRasterConversionStatus.PixelLimitExceeded,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 2, 1, profile,
                new OfficeIccRasterConversionOptions { MaximumPixels = 1 }, out OfficeRasterImage? tooManyPixels));
        Assert.Null(tooManyPixels);
        Assert.Equal(OfficeIccRasterConversionStatus.AllocationLimitExceeded,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1, profile,
                new OfficeIccRasterConversionOptions { MaximumManagedBytes = 1024 }, out OfficeRasterImage? tooLarge));
        Assert.Null(tooLarge);
        Assert.Equal(OfficeIccRasterConversionStatus.InvalidSamples,
            OfficeIccRasterConverter.TryConvertToSrgb(sample.Take(3).ToArray(), 1, 1,
                profile, null, out OfficeRasterImage? wrongChannels));
        Assert.Null(wrongChannels);
    }

    [Fact]
    public void DeclaredLimitsAcceptExactBoundaryAndRejectOneLess() {
        byte[] profile = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory, "TestAssets", "IccColorCorpus", "icc-dci-p3-matrix.icc"));
        byte[] sample = { 64, 128, 192 };
        long budget = sample.LongLength + profile.LongLength + 4L + profile.LongLength * 32L + 4096L;
        var exact = new OfficeIccRasterConversionOptions {
            MaximumProfileBytes = profile.Length,
            MaximumPixels = 1,
            MaximumManagedBytes = budget
        };
        Assert.Equal(OfficeIccRasterConversionStatus.Converted,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1, profile, exact, out OfficeRasterImage? image));
        Assert.NotNull(image);
        exact.MaximumProfileBytes--;
        Assert.Equal(OfficeIccRasterConversionStatus.ProfileLimitExceeded,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1, profile, exact, out _));
        exact.MaximumProfileBytes++;
        exact.MaximumManagedBytes--;
        Assert.Equal(OfficeIccRasterConversionStatus.AllocationLimitExceeded,
            OfficeIccRasterConverter.TryConvertToSrgb(sample, 1, 1, profile, exact, out _));
    }

    private static byte[] ParseChannels(string value) =>
        value.Split(':').Select(byte.Parse).ToArray();
}
