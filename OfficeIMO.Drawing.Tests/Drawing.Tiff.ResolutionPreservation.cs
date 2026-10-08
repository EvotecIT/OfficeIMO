using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class TiffResolutionPreservationTests {
    [Theory]
    [InlineData(false, 3)]
    [InlineData(true, 3)]
    [InlineData(false, 7)]
    [InlineData(true, 7)]
    [InlineData(false, 1000000)]
    [InlineData(true, 1000000)]
    public void UnrelatedTiffEditPreservesExactNativeRationalFields(bool bigEndian, int denominator) {
        byte[] original = MakeTiff(bigEndian, (uint)denominator);
        Assert.True(OfficeTiffStructureValidator.TryValidate(original, 0, original.Length));
        Assert.True(OfficeRasterImageDecoder.TryDecode(original, out OfficeRasterImage? before));
        var metadata = OfficeImageMetadata.Read(original);
        Assert.Equal(1D / denominator, metadata.Resolution.Horizontal);
        metadata.SetExifValue(OfficeExifTag.Artist, "Density preserved");
        byte[] result = OfficeImageMetadata.Apply(original, metadata.Clone());
        var actual = OfficeImageMetadata.Read(result);
        Assert.Equal(1D / denominator, actual.Resolution.Horizontal);
        Assert.Equal(2D / 7D, actual.Resolution.Vertical);
        Assert.Equal(2.54D / denominator, actual.PhysicalDpiX!.Value, 14);
        Assert.Equal(RationalBytes(original, 282), RationalBytes(result, 282));
        Assert.Equal(RationalBytes(original, 283), RationalBytes(result, 283));
        Assert.Equal(OfficeImageResolutionUnit.PixelsPerCentimeter, actual.ResolutionUnits);
        Assert.Equal("Density preserved", actual.GetExifValue(OfficeExifTag.Artist)!.Value);
        Assert.True(OfficeRasterImageDecoder.TryDecode(result, out OfficeRasterImage? after));
        Assert.Equal(before!.GetPixels(), after!.GetPixels());
        // Clearing only a removable profile must not author a new density either.
        metadata.ClearExif();
        byte[] cleared = OfficeImageMetadata.Apply(original, metadata);
        Assert.Equal(RationalBytes(original, 282), RationalBytes(cleared, 282));
        Assert.Equal(RationalBytes(original, 283), RationalBytes(cleared, 283));
    }

    [Theory]
    [InlineData(0.000001)]
    [InlineData(0.3333333333333333)]
    [InlineData(4294967295D)]
    [InlineData(1D / 4294967295D)]
    [InlineData(0.000000001)]
    [InlineData(Math.PI)]
    public void AuthoredTiffAndWebpExifDensityRemainsPositiveAndAccurate(double density) {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Blue);
        foreach (byte[] original in new[] { MakeTiff(false, 3), OfficeWebpCodec.Encode(image) }) {
            var metadata = OfficeImageMetadata.Read(original);
            metadata.HorizontalResolution = density; metadata.VerticalResolution = density;
            metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerInch;
            metadata.SetExifValue(OfficeExifTag.Artist, "New density");
            var actual = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(original, metadata));
            AssertDensity(density, actual.HorizontalResolution);
            AssertDensity(density, actual.VerticalResolution);
            AssertDensity(density, actual.Resolution.Horizontal);
            AssertDensity(density, actual.PhysicalDpiX!.Value);
            Assert.Equal(OfficeImageResolutionUnit.PixelsPerInch, actual.ResolutionUnits);
        }
    }

    [Fact]
    public void WebpUnrelatedEditPreservesExactRationalsCarriedFromTiff() {
        var metadata = OfficeImageMetadata.Read(MakeTiff(true, 3));
        byte[] original = OfficeImageMetadata.Apply(OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red)), metadata);
        var read = OfficeImageMetadata.Read(original); read.SetExifValue(OfficeExifTag.Artist, "WebP edit");
        var actual = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(original, read));
        var x = Assert.IsType<OfficeRational>(actual.GetExifValue(OfficeExifTag.XResolution)!.Value);
        var y = Assert.IsType<OfficeRational>(actual.GetExifValue(OfficeExifTag.YResolution)!.Value);
        Assert.Equal(1U, x.Numerator); Assert.Equal(3U, x.Denominator);
        Assert.Equal(2U, y.Numerator); Assert.Equal(7U, y.Denominator);
    }

    [Theory]
    [InlineData(0.000001)]
    [InlineData(4294967295D)]
    public void AllTiffWritersRepresentNativeDensityWithoutThousandthsOverflowOrZero(double density) {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Blue);
        var options = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(density, density) };
        using var stream = new MemoryStream(); OfficeTiffCodec.EncodeTo(image, stream, options);
        foreach (byte[] output in new[] { OfficeTiffCodec.Encode(image, options), OfficeTiffCodec.EncodePages(new[] { image, image }, options), stream.ToArray() }) {
            var actual = OfficeImageMetadata.Read(output);
            AssertDensity(density, actual.Resolution.Horizontal);
            AssertDensity(density, actual.Resolution.Vertical);
        }
    }

    [Theory]
    [InlineData(0.000000000001)]
    [InlineData(4294967296D)]
    public void UnrepresentableAuthoredRationalDensityFailsBeforeOutput(double density) {
        var metadata = OfficeImageMetadata.Read(MakeTiff(false, 3)); metadata.HorizontalResolution = density;
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageMetadata.Apply(MakeTiff(false, 3), metadata));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageMetadata.Apply(OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1)), metadata));
        using var stream = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTiffCodec.EncodeTo(new OfficeRasterImage(1, 1), stream, new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(density, density) }));
        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public void ExplicitDensityAndUnitEditsWinWhileEqualDensityRetainsNonReducedEncoding() {
        byte[] original = MakeTiff(false, 6, numerator: 2);
        var metadata = OfficeImageMetadata.Read(original);
        metadata.HorizontalResolution = 1D / 3D; metadata.SetExifValue(OfficeExifTag.Artist, "Same value");
        Assert.Equal(RationalBytes(original, 282), RationalBytes(OfficeImageMetadata.Apply(original, metadata), 282));
        metadata.HorizontalResolution = 0.0001; metadata.VerticalResolution = 0.0002;
        metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerMeter;
        var actual = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(original, metadata));
        AssertDensity(0.000001, actual.HorizontalResolution); AssertDensity(0.000002, actual.VerticalResolution);
        Assert.Equal(OfficeImageResolutionUnit.PixelsPerCentimeter, actual.ResolutionUnits);
    }

    [Theory]
    [InlineData(0.001)]
    [InlineData(96.12345)]
    [InlineData(1000000D)]
    public void LegacyTiffDpiRetainsItsSupportedRangeWithAccurateRationals(double density) {
        byte[] output = OfficeTiffCodec.Encode(new OfficeRasterImage(1, 1), new OfficeTiffEncodeOptions { DpiX = density, DpiY = density });
        AssertDensity(density, OfficeImageMetadata.Read(output).PhysicalDpiX!.Value);
    }

    [Theory]
    [InlineData(0.0001)]
    [InlineData(96.12345)]
    [InlineData(1000000D)]
    public void WebpDensityWriterRetainsItsSupportedRangeWithAccurateRationals(double density) {
        var image = new OfficeRasterImage(1, 1, OfficeColor.Blue);
        byte[] output = OfficeWebpCodec.Encode(image, density, density);
        AssertDensity(density, OfficeImageMetadata.Read(output).PhysicalDpiX!.Value);
        Assert.True(OfficeWebpCodec.TryDecode(output, out OfficeRasterImage? decoded));
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void LegacyDpiConstraintsRemainUnchanged() {
        var image = new OfficeRasterImage(1, 1);
        foreach (double value in new[] { 0.00001, 1000001D }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTiffCodec.Encode(image, new OfficeTiffEncodeOptions { DpiX = value }));
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeWebpCodec.Encode(image, value, 96));
        }
    }

    [Fact]
    public void ExifDensityRewriterUsesAccuratePositiveStorage() {
        byte[] exif = OfficeImageMetadata.Read(MakeTiff(true, 3)).EncodeExifProfile()!;
        Assert.True(OfficeExifMetadataEditor.TryRewritePhysicalResolution(exif, 0.000001, 96.12345, out byte[] rewritten));
        var metadata = OfficeImageMetadata.ParseExifProfile(rewritten);
        AssertDensity(0.000001, Assert.IsType<OfficeRational>(metadata.GetExifValue(OfficeExifTag.XResolution)!.Value).ToDouble());
        AssertDensity(96.12345, Assert.IsType<OfficeRational>(metadata.GetExifValue(OfficeExifTag.YResolution)!.Value).ToDouble());
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeExifMetadataEditor.TryRewritePhysicalResolution(exif, 1E-12, 96, out _));
    }

    private static void AssertDensity(double expected, double actual) {
        Assert.True(actual > 0D && Math.Abs(actual - expected) / expected <= 1E-12, $"Expected {expected:R}, actual {actual:R}.");
    }
    private static byte[] MakeTiff(bool bigEndian, uint denominator, uint numerator = 1) {
        // Independent 2x1 uncompressed grayscale TIFF with exact native rationals.
        const int count = 13, x = 8 + 6 + count * 12, y = x + 8, pixels = y + 8;
        var bytes = new byte[pixels + 2]; bool little = !bigEndian; bytes[0] = bytes[1] = (byte)(little ? 73 : 77);
        Put(2, 42, 2); Put(4, 8, 4); Put(8, count, 2); int at = 10;
        Entry(256, 4, 2); Entry(257, 4, 1); Entry(258, 3, 8); Entry(259, 3, 1); Entry(262, 3, 1);
        Entry(273, 4, pixels); Entry(277, 3, 1); Entry(278, 4, 1); Entry(279, 4, 2);
        Entry(282, 5, x); Entry(283, 5, y); Entry(284, 3, 1); Entry(296, 3, 3);
        Put(x, numerator, 4); Put(x + 4, denominator, 4); Put(y, 2, 4); Put(y + 4, 7, 4); bytes[pixels] = 20; bytes[pixels + 1] = 190;
        return bytes;
        void Entry(uint tag, uint type, uint value) { Put(at, tag, 2); Put(at + 2, type, 2); Put(at + 4, 1, 4); Put(at + 8, value, type == 3 ? 2 : 4); at += 12; }
        void Put(int offset, uint value, int length) { for (int i = 0; i < length; i++) bytes[offset + i] = (byte)(value >> (8 * (little ? i : length - i - 1))); }
    }
    private static byte[] RationalBytes(byte[] bytes, uint tag) {
        bool little = bytes[0] == 73; int root = (int)Read(4, 4); int count = (int)Read(root, 2);
        for (int i = 0; i < count; i++) { int entry = root + 2 + i * 12; if (Read(entry, 2) != tag) continue; int value = (int)Read(entry + 8, 4); var result = new byte[8]; Buffer.BlockCopy(bytes, value, result, 0, 8); return result; }
        throw new InvalidDataException("Missing fixture resolution.");
        uint Read(int offset, int length) { uint value = 0; for (int i = 0; i < length; i++) value |= (uint)bytes[offset + i] << (8 * (little ? i : length - i - 1)); return value; }
    }
}
