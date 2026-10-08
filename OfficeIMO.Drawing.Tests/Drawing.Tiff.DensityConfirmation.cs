using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class TiffDensityConfirmationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UIntBoundaryIntermediateFractionsRemainRepresentableAcrossDensityWriters(bool inverse) {
        double value = inverse ? 1D - 1D / 4294967296D : 1D + 1D / 4294967296D;
        var image = new OfficeRasterImage(1, 1, OfficeColor.Blue);
        byte[] tiff = OfficeTiffCodec.Encode(image, new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(value, value) });
        byte[] webp = OfficeWebpCodec.Encode(image, value, value);
        foreach (byte[] original in new[] { tiff, webp }) {
            var metadata = OfficeImageMetadata.Read(original);
            Assert.Equal(value, metadata.Resolution.Horizontal);
            metadata.HorizontalResolution = value; metadata.VerticalResolution = value;
            metadata.SetExifValue(OfficeExifTag.Artist, "Representable boundary");
            Assert.Equal(value, OfficeImageMetadata.Read(OfficeImageMetadata.Apply(original, metadata)).Resolution.Horizontal);
        }
        var fresh = new OfficeImageMetadata { HorizontalResolution = value, VerticalResolution = value };
        Assert.Equal(value, OfficeImageMetadata.Read(OfficeImageMetadata.Apply(OfficeWebpCodec.Encode(image), fresh)).Resolution.Horizontal);
        byte[] exif = OfficeImageMetadata.Read(OfficeTiffCodec.Encode(image)).EncodeExifProfile()!;
        Assert.True(OfficeExifMetadataEditor.TryRewritePhysicalResolution(exif, value, value, out byte[] edited));
        var fields = OfficeImageMetadata.ParseExifProfile(edited);
        Assert.Equal(value, Assert.IsType<OfficeRational>(fields.GetExifValue(OfficeExifTag.XResolution)!.Value).ToDouble());
    }

    [Theory]
    [InlineData(0.000001, OfficeImageResolutionUnit.PixelsPerInch)]
    [InlineData(0.000001, OfficeImageResolutionUnit.PixelsPerCentimeter)]
    [InlineData(0.000001, OfficeImageResolutionUnit.AspectRatio)]
    [InlineData(4294967295D, OfficeImageResolutionUnit.PixelsPerInch)]
    public void EveryGenericTiffRoutePreservesValidNativeDensityAndCallerSettings(double value, OfficeImageResolutionUnit unit) {
        var options = new OfficeRasterEncodingOptions { Tiff = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(value, value, unit) } };
        foreach (byte[] output in EncodeAll(options)) {
            var metadata = OfficeImageMetadata.Read(output);
            Assert.Equal(value, metadata.Resolution.Horizontal); Assert.Equal(value, metadata.Resolution.Vertical);
            Assert.Equal(unit, metadata.ResolutionUnits);
        }
        Assert.Equal(96D, options.Tiff.DpiX); Assert.Equal(96D, options.Tiff.DpiY);
        Assert.Equal(unit, options.Tiff.Resolution!.Unit); Assert.Equal(value, options.Tiff.Resolution.Horizontal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OneAuthoredSharedDpiAxisOverridesNativeUnitsWithoutValidatingTheOtherDerivedAxisAsLegacyDpi(bool vertical) {
        var options = new OfficeRasterEncodingOptions { Tiff = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(0.000001, 0.000001, OfficeImageResolutionUnit.PixelsPerCentimeter) } };
        if (vertical) options.DpiY = 120D; else options.DpiX = 144D;
        foreach (byte[] output in EncodeAll(options)) {
            var metadata = OfficeImageMetadata.Read(output);
            Assert.Equal(OfficeImageResolutionUnit.PixelsPerInch, metadata.ResolutionUnits);
            Assert.Equal(vertical ? 0.00000254 : 144D, metadata.HorizontalResolution, 14);
            Assert.Equal(vertical ? 120D : 0.00000254, metadata.VerticalResolution, 14);
        }
    }

    [Fact]
    public void TwoSharedDpiAssignmentsTakePrecedenceAndRemainBoundedByLegacyDpiRules() {
        var options = new OfficeRasterEncodingOptions { DpiX = 144, DpiY = 120, Tiff = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(0.000001, 0.000001) } };
        foreach (byte[] output in EncodeAll(options)) {
            var metadata = OfficeImageMetadata.Read(output);
            Assert.Equal(144D, metadata.HorizontalResolution); Assert.Equal(120D, metadata.VerticalResolution);
        }
        foreach (double value in new[] { 0.000001, 1000001D }) {
            options.DpiX = value;
            AssertRejectedByAllRoutes(options);
        }
    }

    [Theory]
    [InlineData(0.000000000001)]
    [InlineData(4294967296D)]
    public void UnsupportedNativeDensityIsRejectedBeforeAnyGenericStreamBytes(double value) {
        AssertRejectedByAllRoutes(new OfficeRasterEncodingOptions { Tiff = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(value, value) } });
    }

    private static byte[][] EncodeAll(OfficeRasterEncodingOptions options) {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Blue);
        byte[] ordinary = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options);
        byte[] bounded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options.Clone(), 4096, CancellationToken.None);
        using var stream = new MemoryStream(); OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Tiff, stream, options);
        using var boundedStream = new MemoryStream(); OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Tiff, boundedStream, options.Clone(), 4096, CancellationToken.None);
        Assert.Equal(ordinary, bounded); Assert.Equal(ordinary, stream.ToArray()); Assert.Equal(ordinary, boundedStream.ToArray());
        return new[] { ordinary, bounded, stream.ToArray(), boundedStream.ToArray() };
    }
    private static void AssertRejectedByAllRoutes(OfficeRasterEncodingOptions options) {
        var image = new OfficeRasterImage(1, 1);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options, 4096));
        using var stream = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Tiff, stream, options));
        Assert.Equal(0, stream.Length);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Tiff, stream, options, 4096));
        Assert.Equal(0, stream.Length);
    }
}
