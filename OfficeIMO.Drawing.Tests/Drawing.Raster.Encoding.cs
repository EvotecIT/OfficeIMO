using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterEncodingTests {
    [Theory]
    [InlineData(OfficeTiffCompression.None)]
    [InlineData(OfficeTiffCompression.PackBits)]
    public void OfficeTiffCodec_EncodesIdentifiableRgbaTiff(OfficeTiffCompression compression) {
        OfficeRasterImage image = CreateSampleImage();

        byte[] encoded = OfficeTiffCodec.Encode(image, new OfficeTiffEncodeOptions {
            Compression = compression,
            DpiX = 144D,
            DpiY = 120D
        });

        Assert.True(OfficeTiffCodec.IsTiff(encoded));
        OfficeImageInfo info = OfficeImageReader.Identify(encoded);
        Assert.Equal(OfficeImageFormat.Tiff, info.Format);
        Assert.Equal(3, info.Width);
        Assert.Equal(2, info.Height);
        Assert.Equal(144D, info.DpiX, precision: 3);
        Assert.Equal(120D, info.DpiY, precision: 3);
    }

    [Fact]
    public void OfficeWebpCodec_EncodesIdentifiableLosslessRgbaWebp() {
        OfficeRasterImage image = CreateSampleImage();

        byte[] encoded = OfficeWebpCodec.Encode(image);

        Assert.True(OfficeWebpCodec.IsWebp(encoded));
        Assert.Equal("VP8L", System.Text.Encoding.ASCII.GetString(encoded, 12, 4));
        Assert.Equal(0, encoded.Length % 2);
        OfficeImageInfo info = OfficeImageReader.Identify(encoded);
        Assert.Equal(OfficeImageFormat.Webp, info.Format);
        Assert.Equal(3, info.Width);
        Assert.Equal(2, info.Height);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff, OfficeImageFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
    public void OfficeRasterImageEncoder_RoutesSharedRasterFormats(
        OfficeImageExportFormat format,
        OfficeImageFormat expected) {
        byte[] encoded = OfficeRasterImageEncoder.Encode(CreateSampleImage(), format);

        Assert.Equal(expected, OfficeImageReader.Identify(encoded).Format);
    }

    [Fact]
    public void OfficeRasterImageEncoder_RejectsVectorOutput() {
        Assert.Throws<ArgumentException>(() =>
            OfficeRasterImageEncoder.Encode(CreateSampleImage(), OfficeImageExportFormat.Svg));
    }

    [Fact]
    public void OfficeRasterEncodingOptions_CloneDoesNotShareNestedSettings() {
        var source = new OfficeRasterEncodingOptions();
        OfficeRasterEncodingOptions clone = source.Clone();

        clone.Jpeg.Quality = 42;
        clone.Tiff.Compression = OfficeTiffCompression.None;

        Assert.Equal(85, source.Jpeg.Quality);
        Assert.Equal(OfficeTiffCompression.PackBits, source.Tiff.Compression);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png, ".png", "image/png", true)]
    [InlineData(OfficeImageExportFormat.Svg, ".svg", "image/svg+xml", false)]
    [InlineData(OfficeImageExportFormat.Jpeg, ".jpg", "image/jpeg", true)]
    [InlineData(OfficeImageExportFormat.Tiff, ".tiff", "image/tiff", true)]
    [InlineData(OfficeImageExportFormat.Webp, ".webp", "image/webp", true)]
    public void OfficeImageExportFormat_ProvidesSharedMetadata(
        OfficeImageExportFormat format,
        string extension,
        string mimeType,
        bool raster) {
        Assert.Equal(extension, format.GetFileExtension());
        Assert.Equal(mimeType, format.GetMimeType());
        Assert.Equal(raster, format.IsRaster());
    }

    private static OfficeRasterImage CreateSampleImage() {
        var image = new OfficeRasterImage(3, 2, OfficeColor.Transparent);
        image.SetPixel(0, 0, OfficeColor.FromRgba(255, 0, 0, 255));
        image.SetPixel(1, 0, OfficeColor.FromRgba(0, 255, 0, 128));
        image.SetPixel(2, 0, OfficeColor.FromRgba(0, 0, 255, 0));
        image.SetPixel(0, 1, OfficeColor.FromRgba(12, 34, 56, 255));
        image.SetPixel(1, 1, OfficeColor.FromRgba(78, 90, 123, 200));
        image.SetPixel(2, 1, OfficeColor.FromRgba(210, 220, 230, 255));
        return image;
    }
}
