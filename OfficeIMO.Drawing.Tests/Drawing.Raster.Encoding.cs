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
    [InlineData(OfficeTiffCompression.Lzw)]
    [InlineData(OfficeTiffCompression.PackBits)]
    [InlineData(OfficeTiffCompression.Deflate)]
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
    [InlineData(OfficeTiffCompression.Lzw)]
    [InlineData(OfficeTiffCompression.PackBits)]
    [InlineData(OfficeTiffCompression.Deflate)]
    public void SharedRasterDecoderRepaintsEncodedTiff(OfficeTiffCompression compression) {
        OfficeRasterImage expected = CreateSampleImage();
        byte[] encoded = OfficeTiffCodec.Encode(expected, new OfficeTiffEncodeOptions { Compression = compression });

        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(expected.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void TiffAxisSwappingOrientationAlsoSwapsPhysicalResolution() {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Red);
        byte[] tiff = OfficeTiffCodec.Encode(image, new OfficeTiffEncodeOptions {
            DpiX = 300D,
            DpiY = 150D
        });
        int orientationEntry = FindClassicTiffEntry(tiff, 274);
        WriteLittleEndian(tiff, orientationEntry + 8, 6);

        OfficeImageInfo oriented = OfficeImageReader.Identify(tiff);
        Assert.Equal((1, 2), (oriented.Width, oriented.Height));
        Assert.Equal(150D, oriented.DpiX, 3);
        Assert.Equal(300D, oriented.DpiY, 3);

        Assert.True(OfficeImagePngConverter.TryConvertToPng(tiff, out byte[] png));
        OfficeImageInfo converted = OfficeImageReader.Identify(png);
        Assert.Equal((1, 2), (converted.Width, converted.Height));
        Assert.InRange(converted.DpiX, 149.98D, 150.02D);
        Assert.InRange(converted.DpiY, 299.98D, 300.02D);

        Assert.True(OfficeImageOrientationNormalizer.TryNormalizeToPng(tiff, true, out _, out OfficeImageInfo? normalized));
        Assert.Equal((1, 2), (normalized!.Width, normalized.Height));
        Assert.InRange(normalized.DpiX, 149.98D, 150.02D);
        Assert.InRange(normalized.DpiY, 299.98D, 300.02D);

        Assert.True(OfficeImageOrientationNormalizer.TryNormalizeToPng(tiff, false, out _, out OfficeImageInfo? ignored));
        Assert.Equal((2, 1), (ignored!.Width, ignored.Height));
        Assert.InRange(ignored.DpiX, 299.98D, 300.02D);
        Assert.InRange(ignored.DpiY, 149.98D, 150.02D);
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

    [Fact]
    public void OfficeTiffDecoderValidatesTheCompleteIfdChain() {
        byte[] first = OfficeTiffCodec.Encode(CreateSampleImage());
        byte[] second = OfficeTiffCodec.Encode(CreateSampleImage());
        int firstIfdOffset = ReadLittleEndian(first, 4);
        int firstEntryCount = first[firstIfdOffset] | first[firstIfdOffset + 1] << 8;
        int firstNextIfdPointerOffset = firstIfdOffset + 2 + firstEntryCount * 12;
        int secondIfdOffset = first.Length;
        int secondOffsetAdjustment = secondIfdOffset - 8;
        byte[] chained = new byte[first.Length + second.Length - 8];
        Buffer.BlockCopy(first, 0, chained, 0, first.Length);
        Buffer.BlockCopy(second, 8, chained, secondIfdOffset, second.Length - 8);
        WriteLittleEndian(chained, firstNextIfdPointerOffset, secondIfdOffset);
        int secondEntryCount = chained[secondIfdOffset] | chained[secondIfdOffset + 1] << 8;
        for (int index = 0; index < secondEntryCount; index++) {
            int entryOffset = secondIfdOffset + 2 + index * 12;
            int tag = chained[entryOffset] | chained[entryOffset + 1] << 8;
            int type = chained[entryOffset + 2] | chained[entryOffset + 3] << 8;
            int count = ReadLittleEndian(chained, entryOffset + 4);
            int byteCount = type == 3 ? count * 2 : type == 5 ? count * 8 : count * 4;
            if (byteCount > 4 || tag == 273) {
                WriteLittleEndian(
                    chained,
                    entryOffset + 8,
                    ReadLittleEndian(chained, entryOffset + 8) + secondOffsetAdjustment);
            }
        }

        Assert.True(OfficeTiffCodec.TryDecode(chained, out OfficeRasterImage? firstPage));
        Assert.NotNull(firstPage);
        Assert.True(OfficeTiffCodec.TryGetPageCount(chained, out int pageCount));
        Assert.Equal(2, pageCount);

        Assert.True(OfficeRasterImageDecoder.TryDecode(
            chained,
            options: null,
            out OfficeRasterImage? selectedPage,
            out OfficeRasterDecodeInfo selectedInfo));
        Assert.NotNull(selectedPage);
        Assert.Equal(2, selectedInfo.FrameCount);
        Assert.True(selectedInfo.PagesDiscarded);
        Assert.False(selectedInfo.AnimationDiscarded);

        var rejectFrameLoss = new OfficeRasterDecodeOptions {
            AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated
        };
        Assert.False(OfficeRasterImageDecoder.TryDecode(
            chained,
            rejectFrameLoss,
            out OfficeRasterImage? rejectedPage,
            out OfficeRasterDecodeInfo rejectedInfo));
        Assert.Null(rejectedPage);
        Assert.Equal(2, rejectedInfo.FrameCount);

        OfficeImageOptimizationResult optimization = OfficeImageOptimizer.Optimize(
            chained,
            new OfficeImageOptimizationRequest(1, 1) { KeepOriginalWhenNotSmaller = false });
        Assert.Equal(OfficeImageOptimizationStatus.DecodeFailed, optimization.Status);

        byte[] cyclic = (byte[])first.Clone();
        WriteLittleEndian(cyclic, firstNextIfdPointerOffset, firstIfdOffset);
        Assert.False(OfficeTiffCodec.TryDecode(cyclic, out _));
        Assert.False(OfficeTiffCodec.TryGetPageCount(cyclic, out _));

        byte[] truncated = chained.Take(chained.Length - 1).ToArray();
        Assert.False(OfficeTiffCodec.TryDecode(truncated, out OfficeRasterImage? partial));
        Assert.Null(partial);
        Assert.False(OfficeImageReader.TryValidateContent(truncated, "truncated-chain.tiff", out _));
    }

    [Fact]
    public void OfficeWebpCodecDecodesOrdinaryLosslessVp8lWithTransformsAndBackReferences() {
        byte[] mixed = Convert.FromBase64String(
            "UklGRmIAAABXRUJQVlA4TFYAAAAvB8ABELmM6H/sIqL/ATNt2+wzmPAnuJIQC6a4P44ZBJhmDjnkkEMOd7hD3gKBJNXJVu3BRTCAVjOCQq8GDo+tJjJBWs0Ujb4aXNyc1cTP41vNDIv9MQ==");
        byte[] flat = Convert.FromBase64String(
            "UklGRh4AAABXRUJQVlA4TBEAAAAvH8ADEAdQkTIUp8iBiOh/AAA=");

        Assert.True(OfficeWebpCodec.TryDecode(flat, out OfficeRasterImage? flatImage));
        Assert.NotNull(flatImage);
        Assert.Equal((32, 16), (flatImage!.Width, flatImage.Height));
        Assert.Equal(OfficeColor.FromRgba(12, 34, 56, 200), flatImage.GetPixel(31, 15));

        Assert.True(OfficeWebpCodec.TryDecode(mixed, out OfficeRasterImage? mixedImage));
        Assert.NotNull(mixedImage);
        Assert.Equal((8, 8), (mixedImage!.Width, mixedImage.Height));
        Assert.Equal(OfficeColor.FromRgba(0, 0, 0, 96), mixedImage.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(159, 71, 76, 255), mixedImage.GetPixel(4, 1));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeWebpCodecDecodesShortenedAndDuplicateSymbolHuffmanTrees(
        bool duplicateSimpleTree) {
        byte[] webp = CreateVp8lHuffmanEdgeFixture(duplicateSimpleTree);

        Assert.True(OfficeWebpCodec.TryDecode(webp, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal((1, 1), (image!.Width, image.Height));
        Assert.Equal(OfficeColor.FromRgba(11, 0, 22, 255), image.GetPixel(0, 0));
    }

    [Fact]
    public void OfficeWebpCodecUsesDeterministicPredictionAndLz77WhenTheyReduceThePayload() {
        var source = new OfficeRasterImage(128, 64, OfficeColor.FromRgba(12, 34, 56, 200));

        byte[] first = OfficeWebpCodec.Encode(source);
        byte[] second = OfficeWebpCodec.Encode(source);

        Assert.Equal(first, second);
        Assert.True(first.Length < source.GetPixels().Length / 4);
        Assert.True(OfficeWebpCodec.TryDecode(first, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(source.GetPixels(), decoded!.GetPixels());
    }

    [Theory]
    [InlineData(OfficeTiffCompression.None)]
    [InlineData(OfficeTiffCompression.Lzw)]
    [InlineData(OfficeTiffCompression.PackBits)]
    [InlineData(OfficeTiffCompression.Deflate)]
    public void OfficeTiffCodecWritesAndSelectsMultiplePages(OfficeTiffCompression compression) {
        var first = new OfficeRasterImage(2, 1, OfficeColor.Red);
        var second = new OfficeRasterImage(1, 2, OfficeColor.Blue);
        byte[] encoded = OfficeTiffCodec.EncodePages(
            new[] { first, second },
            new OfficeTiffEncodeOptions { Compression = compression });

        Assert.True(OfficeTiffCodec.TryGetPageCount(encoded, out int pageCount));
        Assert.Equal(2, pageCount);
        Assert.True(OfficeTiffCodec.TryDecodePage(encoded, 1, out OfficeRasterImage? selected));
        Assert.NotNull(selected);
        Assert.Equal((1, 2), (selected!.Width, selected.Height));
        Assert.Equal(OfficeColor.Blue, selected.GetPixel(0, 1));

        var options = new OfficeRasterDecodeOptions { FrameIndex = 1 };
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, options, out selected, out OfficeRasterDecodeInfo info));
        Assert.True(info.PagesDiscarded);
        Assert.Equal(OfficeRasterFrameKind.Page, info.SelectedFrame!.Kind);
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
        clone.Tiff.Predictor = OfficeTiffPredictor.None;

        Assert.Equal(85, source.Jpeg.Quality);
        Assert.Equal(OfficeTiffCompression.PackBits, source.Tiff.Compression);
        Assert.Equal(OfficeTiffPredictor.Horizontal, source.Tiff.Predictor);
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

    private static int FindClassicTiffEntry(byte[] bytes, int expectedTag) {
        int ifdOffset = ReadLittleEndian(bytes, 4);
        int entryCount = bytes[ifdOffset] | bytes[ifdOffset + 1] << 8;
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            int tag = bytes[entryOffset] | bytes[entryOffset + 1] << 8;
            if (tag == expectedTag) return entryOffset;
        }
        throw new InvalidOperationException("TIFF entry was not found.");
    }

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

    [Fact]
    public void OfficeWebpCodecUsesArithmeticShiftForVp8lPredictorMode13() {
        const uint left = 0xFF646464U;
        const uint top = 0xFF000000U;
        const uint topLeft = 0xFF676767U;

        uint predicted = OfficeWebpCodec.PredictVp8l(13, left, top, topLeft, 0U);

        Assert.Equal(0xFF171717U, predicted);
    }

    private static byte[] CreateVp8lHuffmanEdgeFixture(bool duplicateSimpleTree) {
        var writer = new TestLsbBitWriter();
        writer.WriteBits(0, 14); // width - 1
        writer.WriteBits(0, 14); // height - 1
        writer.WriteBits(0, 1);  // no alpha hint
        writer.WriteBits(0, 3);  // version
        writer.WriteBits(0, 1);  // no transforms
        writer.WriteBits(0, 1);  // no color cache
        writer.WriteBits(0, 1);  // one Huffman group
        if (duplicateSimpleTree) {
            WriteDuplicateSimpleTree(writer, 0);
        } else {
            WriteShortenedRleTree(writer);
        }
        WriteSingleSymbolTree(writer, 11);
        WriteSingleSymbolTree(writer, 22);
        WriteSingleSymbolTree(writer, 255);
        WriteSingleSymbolTree(writer, 0);
        if (!duplicateSimpleTree) writer.WriteBits(0, 1); // green symbol zero

        byte[] bits = writer.Finish();
        byte[] payload = new byte[bits.Length + 1];
        payload[0] = 0x2F;
        Buffer.BlockCopy(bits, 0, payload, 1, bits.Length);
        int paddedPayloadLength = payload.Length + (payload.Length & 1);
        byte[] result = new byte[20 + paddedPayloadLength];
        System.Text.Encoding.ASCII.GetBytes("RIFF").CopyTo(result, 0);
        WriteLittleEndian(result, 4, result.Length - 8);
        System.Text.Encoding.ASCII.GetBytes("WEBP").CopyTo(result, 8);
        System.Text.Encoding.ASCII.GetBytes("VP8L").CopyTo(result, 12);
        WriteLittleEndian(result, 16, payload.Length);
        Buffer.BlockCopy(payload, 0, result, 20, payload.Length);
        return result;
    }

    private static void WriteShortenedRleTree(TestLsbBitWriter writer) {
        writer.WriteBits(0, 1); // normal tree
        writer.WriteBits(0, 4); // four code-length code lengths
        writer.WriteBits(1, 3); // symbol 17
        writer.WriteBits(0, 3); // symbol 18
        writer.WriteBits(0, 3); // symbol 0
        writer.WriteBits(1, 3); // symbol 1
        writer.WriteBits(1, 1); // shortened alphabet
        writer.WriteBits(0, 3); // two bits encode max_symbol
        writer.WriteBits(1, 2); // read three encoded code-length symbols
        writer.WriteBits(0, 1); // length 1
        writer.WriteBits(0, 1); // length 1
        writer.WriteBits(1, 1); // repeat zero (symbol 17)
        writer.WriteBits(0, 3); // repeat three zeros
    }

    private static void WriteDuplicateSimpleTree(TestLsbBitWriter writer, int symbol) {
        writer.WriteBits(1, 1); // simple tree
        writer.WriteBits(1, 1); // two encoded symbols
        writer.WriteBits(0, 1); // first symbol uses one bit
        writer.WriteBits((uint)symbol, 1);
        writer.WriteBits((uint)symbol, 8);
    }

    private static void WriteSingleSymbolTree(TestLsbBitWriter writer, int symbol) {
        writer.WriteBits(1, 1);
        writer.WriteBits(0, 1);
        writer.WriteBits(symbol > 1 ? 1U : 0U, 1);
        writer.WriteBits((uint)symbol, symbol > 1 ? 8 : 1);
    }

    private sealed class TestLsbBitWriter {
        private readonly System.Collections.Generic.List<byte> _bytes = new();
        private ulong _buffer;
        private int _bitCount;

        internal void WriteBits(uint value, int count) {
            ulong mask = count == 32 ? uint.MaxValue : (1UL << count) - 1UL;
            _buffer |= ((ulong)value & mask) << _bitCount;
            _bitCount += count;
            while (_bitCount >= 8) {
                _bytes.Add((byte)_buffer);
                _buffer >>= 8;
                _bitCount -= 8;
            }
        }

        internal byte[] Finish() {
            if (_bitCount > 0) _bytes.Add((byte)_buffer);
            return _bytes.ToArray();
        }
    }
}
