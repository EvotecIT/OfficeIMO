using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterEncodingTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void SharedRasterEncoderRejectsDensityThatWouldSerializeAsZero(
        OfficeImageExportFormat format) {
        OfficeRasterEncodingOptions options = CreateSubMinimumDensityOptions(format);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            OfficeRasterImageEncoder.Encode(
                CreateSampleImage(),
                format,
                options));
    }

    [Fact]
    public void SharedRasterOptionsCloneAndEncoderPreserveNestedTiffDensity() {
        OfficeRasterImage image = CreateSampleImage();
        var options = new OfficeRasterEncodingOptions {
            Tiff = new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.PackBits,
                DpiX = 144D,
                DpiY = 120D
            }
        };

        OfficeRasterEncodingOptions clone = options.Clone();
        byte[] encoded = OfficeRasterImageEncoder.Encode(
            image,
            OfficeImageExportFormat.Tiff,
            options);

        Assert.Equal(144D, clone.Tiff.DpiX);
        Assert.Equal(120D, clone.Tiff.DpiY);
        OfficeImageInfo info = OfficeImageReader.Identify(encoded);
        Assert.Equal(144D, info.DpiX, precision: 3);
        Assert.Equal(120D, info.DpiY, precision: 3);
    }

    [Fact]
    public void PngReaderPreservesTheUnsignedPhysicalResolutionRange() {
        const double dpi = 60_000_000D;
        byte[] encoded = OfficePngWriter.Encode(
            CreateSampleImage(),
            new OfficePngEncodeOptions {
                DpiX = dpi,
                DpiY = dpi
            });

        OfficeImageInfo info = OfficeImageReader.Identify(encoded);

        Assert.Equal(OfficeImageFormat.Png, info.Format);
        Assert.InRange(info.DpiX, dpi - 0.02D, dpi + 0.02D);
        Assert.InRange(info.DpiY, dpi - 0.02D, dpi + 0.02D);
    }

    private static OfficeRasterEncodingOptions CreateSubMinimumDensityOptions(
        OfficeImageExportFormat format) {
        var options = new OfficeRasterEncodingOptions();
        switch (format) {
            case OfficeImageExportFormat.Png:
                options.Png.DpiX = 0.01D;
                options.Png.DpiY = 0.01D;
                break;
            case OfficeImageExportFormat.Jpeg:
                options.Jpeg.DpiX = 0.49D;
                options.Jpeg.DpiY = 0.49D;
                break;
            case OfficeImageExportFormat.Tiff:
                options.Tiff.DpiX = 0.0009D;
                options.Tiff.DpiY = 0.0009D;
                break;
            case OfficeImageExportFormat.Webp:
                options.DpiX = 0.00009D;
                options.DpiY = 0.00009D;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }

        return options;
    }

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
        Assert.True(OfficeTiffCodec.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
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
        Assert.True(OfficeWebpCodec.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void SharedWebpEncodingPreservesPhysicalResolutionInStandardExifMetadata() {
        OfficeRasterImage image = CreateSampleImage();
        var options = new OfficeRasterEncodingOptions {
            DpiX = 144D,
            DpiY = 120D
        };

        byte[] encoded = OfficeRasterImageEncoder.Encode(
            image,
            OfficeImageExportFormat.Webp,
            options);

        Assert.Equal("VP8X", System.Text.Encoding.ASCII.GetString(encoded, 12, 4));
        Assert.Contains(
            "EXIF",
            System.Text.Encoding.ASCII.GetString(encoded),
            StringComparison.Ordinal);
        OfficeImageInfo info = OfficeImageReader.Identify(encoded);
        Assert.Equal(144D, info.DpiX, precision: 3);
        Assert.Equal(120D, info.DpiY, precision: 3);
        Assert.True(OfficeWebpCodec.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(image.GetPixels(), decoded!.GetPixels());
    }

    [Theory]
    [InlineData(OfficeTiffCompression.None)]
    [InlineData(OfficeTiffCompression.PackBits)]
    public void SharedRasterDecoderRepaintsEncodedTiff(OfficeTiffCompression compression) {
        OfficeRasterImage expected = CreateSampleImage();
        byte[] encoded = OfficeTiffCodec.Encode(expected, new OfficeTiffEncodeOptions { Compression = compression });

        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(expected.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void SharedRasterDecoderRepaintsOfficeImoLiteralLosslessWebp() {
        OfficeRasterImage expected = CreateSampleImage();
        byte[] encoded = OfficeWebpCodec.Encode(expected);

        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(expected.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void NewSourceDecodersRejectTruncatedPayloadsWithoutAllocating() {
        byte[] tiff = OfficeTiffCodec.Encode(CreateSampleImage());
        byte[] webp = OfficeWebpCodec.Encode(CreateSampleImage());

        Assert.False(OfficeTiffCodec.TryDecode(tiff.Take(tiff.Length - 2).ToArray(), out _));
        Assert.False(OfficeWebpCodec.TryDecode(webp.Take(webp.Length - 2).ToArray(), out _));
    }

    [Fact]
    public void OfficeImoWebpDecoderRejectsBytesOutsideItsExactContainer() {
        byte[] webp = OfficeWebpCodec.Encode(CreateSampleImage());

        Assert.False(OfficeWebpCodec.TryDecode(webp.Concat(new byte[] { 0, 0 }).ToArray(), out _));
    }

    [Fact]
    public void OfficeImoWebpDecoderRejectsNonPaddingDataInsideItsDeclaredPayload() {
        byte[] webp = OfficeWebpCodec.Encode(CreateSampleImage());
        int payloadLength = ReadLittleEndian(webp, 16);
        int expandedPayloadLength = payloadLength + 2;
        byte[] expanded = new byte[20 + expandedPayloadLength + (expandedPayloadLength & 1)];
        Buffer.BlockCopy(webp, 0, expanded, 0, 20 + payloadLength);
        expanded[20 + payloadLength] = 1;
        WriteLittleEndian(expanded, 4, expanded.Length - 8);
        WriteLittleEndian(expanded, 16, expandedPayloadLength);

        Assert.False(OfficeWebpCodec.TryDecode(expanded, out _));
    }

    [Fact]
    public void OfficeImoWebpDecoderRejectsInflatedDimensionsBeforeAllocatingPixels() {
        byte[] webp = OfficeWebpCodec.Encode(CreateSampleImage());
        const int bitstreamOffset = 21;
        WriteLsbBits(webp, bitstreamOffset, 0, 14, 4095);
        WriteLsbBits(webp, bitstreamOffset, 14, 14, 4095);

        Assert.False(OfficeWebpCodec.TryDecode(webp, out _));
    }

    [Fact]
    public void OfficeTiffDecoderRejectsExtraUncompressedStripData() {
        byte[] tiff = OfficeTiffCodec.Encode(
            CreateSampleImage(),
            new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.None });
        Array.Resize(ref tiff, tiff.Length + 1);
        const int stripByteCountValueOffset = 126;
        WriteLittleEndian(tiff, stripByteCountValueOffset, 25);

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
    }

    [Fact]
    public void OfficeTiffDecoderRejectsUnexpectedArrayCardinalityBeforeReadingValues() {
        byte[] tiff = OfficeTiffCodec.Encode(CreateSampleImage());
        const int bitsPerSampleCountOffset = 38;
        WriteLittleEndian(tiff, bitsPerSampleCountOffset, 3);

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
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

    private static int ReadLittleEndian(byte[] bytes, int offset) =>
        bytes[offset] |
        bytes[offset + 1] << 8 |
        bytes[offset + 2] << 16 |
        bytes[offset + 3] << 24;

    private static void WriteLittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteLsbBits(byte[] bytes, int byteOffset, int bitOffset, int bitCount, uint value) {
        for (int bit = 0; bit < bitCount; bit++) {
            int absoluteBit = bitOffset + bit;
            int index = byteOffset + absoluteBit / 8;
            int mask = 1 << (absoluteBit % 8);
            if ((value & (1U << bit)) != 0) {
                bytes[index] = (byte)(bytes[index] | mask);
            } else {
                bytes[index] = (byte)(bytes[index] & ~mask);
            }
        }
    }
}
