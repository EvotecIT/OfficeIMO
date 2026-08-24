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

    [Theory]
    [InlineData((byte)0)]
    [InlineData((byte)0x7F)]
    public void OfficeTiffDecoderRejectsBytesAfterLzwEndCode(byte trailingByte) {
        byte[] tiff = OfficeTiffCodec.Encode(
            CreateSampleImage(),
            new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.Lzw });
        int stripOffsetEntry = FindClassicTiffEntry(tiff, 273);
        int stripByteCountEntry = FindClassicTiffEntry(tiff, 279);
        int stripOffset = ReadLittleEndian(tiff, stripOffsetEntry + 8);
        int stripByteCount = ReadLittleEndian(tiff, stripByteCountEntry + 8);
        Assert.Equal(tiff.Length, stripOffset + stripByteCount);

        Array.Resize(ref tiff, tiff.Length + 1);
        tiff[tiff.Length - 1] = trailingByte;
        WriteLittleEndian(tiff, stripByteCountEntry + 8, stripByteCount + 1);

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
        Assert.True(OfficeTiffCodec.TryDecode(truncated, out OfficeRasterImage? partial));
        Assert.NotNull(partial);
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

    [Fact]
    public void Vp8lInverseColorTransformReadsMultipliersFromSpecifiedChannels() {
        const uint transformedColor = 0xFF402010U;
        const uint transformData = 0xFF030201U;

        uint restored = OfficeWebpCodec.ApplyVp8lInverseColorTransform(transformedColor, transformData);

        Assert.Equal(0xFF412018U, restored);
    }

    [Fact]
    public void Vp8lDecodeUsesTheConfiguredPixelLimitBeyondSixteenMillionPixels() {
        const int width = 4001;
        const int height = 4000;
        byte[] webp = CreateUniformVp8lFixture(width, height);

        Assert.True(OfficeRasterImageDecoder.TryDecode(
            webp,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = (long)width * height },
            out OfficeRasterImage? image,
            out OfficeRasterDecodeInfo info));
        Assert.Equal((width, height), (image!.Width, image.Height));
        Assert.Equal(OfficeColor.FromRgba(11, 0, 22, 255), image.GetPixel(width - 1, height - 1));
        Assert.Null(info.Diagnostic);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void PngUnfilterObservesCancellationWithinWideScanlines(int filter) {
        var current = new byte[8192];
        var previous = new byte[current.Length];
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficePngReader.Unfilter(current, previous, 4, filter, cancellation.Token));
    }

    [Fact]
    public void PngWideRowCopyAndClearObserveCancellation() {
        var source = new byte[128 * 1024];
        var destination = new byte[source.Length];
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => OfficePngReader.CopyBytes(
            source, 0, destination, 0, source.Length, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() =>
            OfficePngReader.ClearBytes(destination, cancellation.Token));
    }

    [Fact]
    public void Vp8lMetaGroupScanObservesCancellation() {
        var prefixImage = new uint[8192];
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeWebpCodec.FindMaximumVp8lGroup(
                prefixImage,
                cancellation.Token));
    }

    [Fact]
    public void Vp8lAllocationBudgetRejectsAggregateBuffersBeyondTheManagedBoundary() {
        var budget = new OfficeWebpCodec.Vp8lAllocationBudget();

        Assert.True(budget.TryReserveBytes(OfficeRasterGuards.MaximumDecodedBytes - 128L));
        Assert.True(budget.TryReserveArray(16, sizeof(uint)));
        Assert.False(budget.TryReserveArray(16, sizeof(uint)));
    }

    [Fact]
    public void LiteralWebpEncodeAndDecodePreflightRetainedBuffersTogether() {
        const long oneHundredTwentyEightMiB = 128L * 1024L * 1024L;

        Assert.False(OfficeWebpCodec.IsEncodingWorkingSetWithinLimit(
            rgbaBytes: oneHundredTwentyEightMiB,
            outputBytes: oneHundredTwentyEightMiB + 2L,
            compressedCandidateBytes: 0L,
            metadataBytes: 0L));
        Assert.False(OfficeWebpCodec.IsEncodingWorkingSetWithinLimit(
            rgbaBytes: oneHundredTwentyEightMiB,
            outputBytes: oneHundredTwentyEightMiB,
            compressedCandidateBytes: 0L,
            metadataBytes: 0L));
        Assert.True(OfficeWebpCodec.IsEncodingWorkingSetWithinLimit(
            rgbaBytes: 4L * 1024L * 1024L,
            outputBytes: 4L * 1024L * 1024L,
            compressedCandidateBytes: 2L * 1024L * 1024L,
            metadataBytes: 86L));

        Assert.False(OfficeWebpCodec.IsLiteralDecodeWorkingSetWithinLimit(
            encodedBytes: oneHundredTwentyEightMiB,
            pixels: oneHundredTwentyEightMiB / 4L));
        Assert.True(OfficeWebpCodec.IsLiteralDecodeWorkingSetWithinLimit(
            encodedBytes: 64L * 1024L,
            pixels: 1024L * 1024L));
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
    public void OfficeWebpCodecDecodesSparseMetaHuffmanGroupIdentifiers() {
        byte[] webp = CreateVp8lSparseHuffmanGroupFixture();

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

    [Fact]
    public void OfficeWebpCodecFallsBackToLiteralsForUnrepresentableDirectDistances() {
        const int width = 1024;
        const int height = 1025;
        int pixelCount = checked(width * height);
        var residuals = new uint[pixelCount];
        for (int index = 0; index < residuals.Length; index++) residuals[index] = 0x02040608U;
        residuals[0] = residuals[pixelCount - 3] = 0x01030507U;
        residuals[1] = residuals[pixelCount - 2] = 0x11131517U;
        residuals[2] = residuals[pixelCount - 1] = 0x21232527U;
        OfficeRasterImage source = CreateImageFromVp8lResiduals(width, height, residuals);

        byte[] encoded = OfficeWebpCodec.Encode(source);

        Assert.True(OfficeWebpCodec.TryDecode(encoded, out OfficeRasterImage? decoded));
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

    [Fact]
    public void OfficeTiffPackBitsPacketsRestartAtEveryRowAcrossEncodingSurfaces() {
        var image = new OfficeRasterImage(1, 2, OfficeColor.FromRgba(1, 2, 3, 4));
        var options = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits };
        byte[] expectedStrip = { 3, 1, 2, 3, 4, 3, 1, 2, 3, 4 };

        byte[] single = OfficeTiffCodec.Encode(image, options);
        byte[] multiPage = OfficeTiffCodec.EncodePages(new[] { image }, options);
        using var stream = new System.IO.MemoryStream();
        OfficeTiffCodec.EncodeTo(image, stream, options);

        Assert.Equal(expectedStrip, ReadTiffStrip(single));
        Assert.Equal(expectedStrip, ReadTiffStrip(multiPage));
        Assert.Equal(expectedStrip, ReadTiffStrip(stream.ToArray()));
    }

    [Fact]
    public void CompleteTiffValidationBoundsAggregatePagePixels() {
        byte[] encoded = OfficeTiffCodec.EncodePages(new[] {
            new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeRasterImage(2, 1, OfficeColor.Blue)
        });

        Assert.True(OfficeTiffCodec.TryValidateAllPages(
            encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 4 }));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(
            encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 3 }));
    }

    [Fact]
    public void TiffContainerInspectionEnforcesPixelLimitForEveryPage() {
        byte[] encoded = OfficeTiffCodec.EncodePages(new[] {
            new OfficeRasterImage(1, 1, OfficeColor.Red),
            new OfficeRasterImage(2, 2, OfficeColor.Blue)
        });

        Assert.True(OfficeRasterContainerInspector.TryInspect(
            encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 4 },
            out OfficeRasterContainerInfo? container));
        Assert.Equal(2, container!.Count);
        Assert.False(OfficeRasterContainerInspector.TryInspect(
            encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 3 },
            out _));
    }

    [Fact]
    public void CompleteTiffValidationChargesEveryVisitToAliasedPackBitsData() {
        byte[] encoded = OfficeTiffCodec.EncodePages(
            new[] {
                new OfficeRasterImage(1, 1, OfficeColor.Red),
                new OfficeRasterImage(1, 1, OfficeColor.Red)
            },
            new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits });
        int firstIfd = ReadLittleEndian(encoded, 4);
        int secondIfd = ReadLittleEndian(
            encoded,
            firstIfd + 2 + ReadUInt16LittleEndian(encoded, firstIfd) * 12);
        int originalStripOffset = ReadTiffLongTag(encoded, firstIfd, 273);
        int originalStripLength = ReadTiffLongTag(encoded, firstIfd, 279);
        const int paddingLength = 8192;
        int aliasedStripOffset = encoded.Length;
        Array.Resize(ref encoded, encoded.Length + originalStripLength + paddingLength);
        Buffer.BlockCopy(encoded, originalStripOffset, encoded, aliasedStripOffset, originalStripLength);
        for (int index = aliasedStripOffset + originalStripLength; index < encoded.Length; index++) {
            encoded[index] = 0x80;
        }
        int aliasedStripLength = originalStripLength + paddingLength;
        SetTiffLongTag(encoded, firstIfd, 273, aliasedStripOffset);
        SetTiffLongTag(encoded, firstIfd, 279, aliasedStripLength);
        SetTiffLongTag(encoded, secondIfd, 273, aliasedStripOffset);
        SetTiffLongTag(encoded, secondIfd, 279, aliasedStripLength);
        long oneVisitWork = aliasedStripLength + 4L;
        var options = new OfficeRasterDecodeOptions { MaximumDecodedPixels = 2 };

        Assert.True(OfficeTiffCodec.TryValidateAllPages(encoded, options, oneVisitWork * 2));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(encoded, options, oneVisitWork * 2 - 1));
        Assert.True(OfficeTiffCodec.TryDecodePage(encoded, 1, out OfficeRasterImage? selected));
        Assert.Equal(OfficeColor.Red, selected!.GetPixel(0, 0));
    }

    [Fact]
    public void CompleteTiffValidationChargesCompressedAndDecodedTileBytes() {
        byte[] encoded = OfficeTiffCodec.Encode(
            new OfficeRasterImage(1, 1, OfficeColor.Blue),
            new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.None });
        int ifd = ReadLittleEndian(encoded, 4);
        int stripOffset = ReadTiffLongTag(encoded, ifd, 273);
        SetTiffTag(encoded, ifd, 273, 322);
        SetTiffLongTag(encoded, ifd, 322, 1);
        SetTiffTag(encoded, ifd, 274, 323);
        SetTiffShortTag(encoded, ifd, 323, 1);
        SetTiffTag(encoded, ifd, 278, 324);
        SetTiffLongTag(encoded, ifd, 324, stripOffset);
        SetTiffTag(encoded, ifd, 279, 325);

        Assert.True(OfficeTiffCodec.TryValidateAllPages(
            encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 },
            maximumValidationWorkBytes: 8));
        Assert.False(OfficeTiffCodec.TryValidateAllPages(
            encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 },
            maximumValidationWorkBytes: 7));
    }

    [Fact]
    public void PackBitsPaddingValidationObservesCancellation() {
        byte[] padding = Enumerable.Repeat((byte)0x80, 8192).ToArray();
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => {
            OfficeTiffCodec.TryValidatePackBitsPadding(
                padding,
                0,
                padding.Length,
                cancellation.Token);
        });
    }

    [Theory]
    [InlineData(259, 6)]
    [InlineData(274, 9)]
    [InlineData(277, 9)]
    public void OfficeTiffCodecDecodesSelectedPageWhenAnotherPageUsesUnsupportedTags(
        int tag,
        int value) {
        byte[] encoded = OfficeTiffCodec.EncodePages(new[] {
            new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeRasterImage(1, 2, OfficeColor.Blue)
        });
        int firstIfd = ReadLittleEndian(encoded, 4);
        int secondIfd = ReadLittleEndian(encoded, firstIfd + 2 + ReadUInt16LittleEndian(encoded, firstIfd) * 12);
        SetTiffShortTag(encoded, secondIfd, tag, value);

        Assert.True(OfficeRasterContainerInspector.TryInspect(encoded, out OfficeRasterContainerInfo? container));
        Assert.Equal(2, container!.Count);
        Assert.True(OfficeTiffCodec.TryDecodePage(encoded, 0, out OfficeRasterImage? selected));
        Assert.Equal(OfficeColor.Red, selected!.GetPixel(0, 0));
        Assert.False(OfficeTiffCodec.TryDecodePage(encoded, 1, out _));
    }

    [Fact]
    public void TiffContainerReportsTheOrientationNormalizedFirstPageCanvas() {
        byte[] encoded = OfficeTiffCodec.Encode(
            new OfficeRasterImage(3, 2, OfficeColor.Red));
        int firstIfd = ReadLittleEndian(encoded, 4);
        SetTiffShortTag(encoded, firstIfd, 274, 6);

        Assert.True(OfficeRasterContainerInspector.TryInspect(
            encoded,
            out OfficeRasterContainerInfo? container));
        OfficeRasterFrameInfo frame = Assert.Single(container!.Frames);
        Assert.Equal((2, 3), (container.CanvasWidth, container.CanvasHeight));
        Assert.Equal((3, 2), (frame.Width, frame.Height));
        Assert.True(OfficeTiffCodec.TryDecodePage(encoded, 0, out OfficeRasterImage? decoded));
        Assert.Equal((2, 3), (decoded!.Width, decoded.Height));
    }

    [Fact]
    public void JpegContainerRequiresACompletePayloadAndReportsTheOrientationNormalizedCanvas() {
        byte[] jpeg = OfficeJpegCodec.Encode(
            new OfficeRasterImage(3, 2, OfficeColor.Red),
            new OfficeJpegEncodeOptions {
                Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(6))
            });
        int scanOffset = FindJpegMarker(jpeg, 0xDA);
        var incomplete = new byte[scanOffset];
        Buffer.BlockCopy(jpeg, 0, incomplete, 0, incomplete.Length);

        Assert.True(OfficeImageReader.TryIdentifyByContent(incomplete, null, out OfficeImageInfo identified));
        Assert.Equal(OfficeImageFormat.Jpeg, identified.Format);
        Assert.False(OfficeRasterContainerInspector.TryInspect(incomplete, out _));
        Assert.True(OfficeRasterContainerInspector.TryInspect(
            jpeg, out OfficeRasterContainerInfo? container));
        OfficeRasterFrameInfo frame = Assert.Single(container!.Frames);
        Assert.Equal((2, 3), (container.CanvasWidth, container.CanvasHeight));
        Assert.Equal((3, 2), (frame.Width, frame.Height));
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out OfficeRasterImage? decoded));
        Assert.Equal((2, 3), (decoded!.Width, decoded.Height));
    }

    [Theory]
    [InlineData(6, 1, 3, 2)]
    [InlineData(1, 6, 2, 3)]
    public void JpegContainerUsesTheLastValidExifOrientation(
        ushort firstOrientation,
        ushort lastOrientation,
        int expectedWidth,
        int expectedHeight) {
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red));
        jpeg = InsertJpegExifSegmentsAfterStart(jpeg, firstOrientation, lastOrientation);

        Assert.True(OfficeRasterContainerInspector.TryInspect(
            jpeg, out OfficeRasterContainerInfo? container));
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out OfficeRasterImage? decoded));
        Assert.Equal((expectedWidth, expectedHeight), (container!.CanvasWidth, container.CanvasHeight));
        Assert.Equal((expectedWidth, expectedHeight), (decoded!.Width, decoded.Height));
    }

    [Fact]
    public void JpegContainerRejectsAFrameOutsideTheManagedRgbaSubset() {
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red));
        byte[] twoComponent = RemoveLastJpegFrameComponent(jpeg);

        Assert.True(OfficeImageReader.TryIdentifyByContent(twoComponent, null, out OfficeImageInfo identified));
        Assert.Equal(OfficeImageFormat.Jpeg, identified.Format);
        Assert.False(OfficeRasterContainerInspector.TryInspect(twoComponent, out _));
        Assert.False(OfficeJpegCodec.TryDecode(twoComponent, out _));
    }

    [Fact]
    public void StaticLossyWebpRemainsOutsideTheManagedContainerSubset() {
        byte[] webp = CreateStaticLossyWebpHeader();

        Assert.True(OfficeImageReader.TryIdentifyByContent(webp, null, out OfficeImageInfo identified));
        Assert.Equal(OfficeImageFormat.Webp, identified.Format);
        Assert.False(OfficeRasterContainerInspector.TryInspect(webp, out _));
        Assert.False(OfficeWebpCodec.TryDecode(webp, out _));
    }

    [Fact]
    public void RasterContainerInspectionRejectsIdentifiableButUnsupportedFormats() {
        byte[] svg = System.Text.Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"2\" height=\"3\"><rect width=\"2\" height=\"3\"/></svg>");
        byte[] icon = CreateSingleEntryIcon(
            OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue)));

        Assert.True(OfficeImageReader.TryIdentifyByContent(svg, null, out OfficeImageInfo svgInfo));
        Assert.Equal(OfficeImageFormat.Svg, svgInfo.Format);
        Assert.False(OfficeRasterContainerInspector.TryInspect(svg, out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(icon, null, out OfficeImageInfo iconInfo));
        Assert.Equal(OfficeImageFormat.Icon, iconInfo.Format);
        Assert.False(OfficeRasterContainerInspector.TryInspect(icon, out _));
    }

    [Fact]
    public void PngDecodeAccountsForEncodedScanlineAndRgbaBuffersTogether() {
        const long oneHundredTwentyMiB = 120L * 1024L * 1024L;
        Assert.False(OfficePngReader.IsDecodeWorkingSetWithinLimit(
            oneHundredTwentyMiB,
            compressedBufferBytes: 0L,
            compressedCopyBytes: 0L,
            width: 6000,
            height: 5000,
            stride: 24_000,
            scanlineBytes: oneHundredTwentyMiB,
            paletteBytes: 0L,
            transparencyBytes: 0L,
            includeRgbaOutput: true));
        Assert.True(OfficePngReader.IsDecodeWorkingSetWithinLimit(
            encodedBytes: 64 * 1024,
            compressedBufferBytes: 64 * 1024,
            compressedCopyBytes: 0L,
            width: 1024,
            height: 1024,
            stride: 4096,
            scanlineBytes: 4L * 1024L * 1024L,
            paletteBytes: 0L,
            transparencyBytes: 0L,
            includeRgbaOutput: true));
    }

    [Fact]
    public void JpegDecodeAccountsForTheRetainedEncodedPayload() {
        Assert.False(OfficeJpegReader.TryInitializeDecodeWorkingSet(
            retainedEncodedBytes: OfficeRasterGuards.MaximumEncodedBytes,
            width: 8192,
            height: 4096,
            orientation: 1,
            out _));
        Assert.True(OfficeJpegReader.TryInitializeDecodeWorkingSet(
            retainedEncodedBytes: 64 * 1024,
            width: 1024,
            height: 1024,
            orientation: 1,
            out long reservedBytes));
        Assert.True(reservedBytes > 4L * 1024L * 1024L);
    }

    [Fact]
    public void BmpDecodeAccountsForTheRetainedEncodedPayloadAndRgbaOutput() {
        Assert.False(OfficeBmpReader.IsDecodeWorkingSetWithinLimit(
            encodedBytes: OfficeRasterGuards.MaximumEncodedBytes,
            width: 8192,
            height: 4096));
        Assert.True(OfficeBmpReader.IsDecodeWorkingSetWithinLimit(
            encodedBytes: 64 * 1024,
            width: 1024,
            height: 1024));
    }

    [Fact]
    public void JpegDecodeChargesALateExifOrientationCanvasOnce() {
        const int width = 1024;
        const int height = 1024;
        Assert.True(OfficeJpegReader.TryInitializeDecodeWorkingSet(
            retainedEncodedBytes: 64 * 1024,
            width,
            height,
            orientation: 1,
            out long reservedBytes));
        long initialReservation = reservedBytes;
        bool orientationCanvasReserved = false;

        Assert.True(OfficeJpegReader.TryReserveOrientationCanvas(
            width, height, ref reservedBytes, ref orientationCanvasReserved));
        Assert.Equal(initialReservation + (long)width * height * 4L, reservedBytes);
        Assert.True(OfficeJpegReader.TryReserveOrientationCanvas(
            width, height, ref reservedBytes, ref orientationCanvasReserved));
        Assert.Equal(initialReservation + (long)width * height * 4L, reservedBytes);

        reservedBytes = OfficeRasterGuards.MaximumDecodedBytes - (long)width * height * 4L + 1L;
        orientationCanvasReserved = false;
        Assert.False(OfficeJpegReader.TryReserveOrientationCanvas(
            width, height, ref reservedBytes, ref orientationCanvasReserved));
        Assert.False(orientationCanvasReserved);
    }

    [Fact]
    public void JpegDecodeRejectsASecondFrameSegment() {
        byte[] jpeg = OfficeJpegCodec.Encode(CreateSampleImage(), new OfficeJpegEncodeOptions());
        byte[] duplicateFrame = DuplicateFirstJpegFrameSegment(jpeg);

        Assert.False(OfficeJpegCodec.TryDecode(duplicateFrame, out _));
    }

    [Fact]
    public void PngSuggestedPaletteNamesHaveABoundedRetainedMetadataBudget() {
        long metadataBytes = 0L;
        Assert.True(OfficePngContainerValidator.TryReserveSuggestedPaletteName(
            encodedBytes: 64 * 1024,
            nameLength: 79,
            ref metadataBytes));

        metadataBytes = OfficePngContainerValidator.MaximumSuggestedPaletteMetadataBytes - 1L;
        Assert.False(OfficePngContainerValidator.TryReserveSuggestedPaletteName(
            encodedBytes: 64 * 1024,
            nameLength: 1,
            ref metadataBytes));
    }

    [Fact]
    public void PngCompressedMetadataAccountsForSourceCopyAndExpansionBuffers() {
        Assert.True(OfficePngContainerValidator.TryGetCompressedMetadataOutputLimit(
            encodedBytes: 64 * 1024,
            compressedBytes: 64 * 1024,
            requestedMaximumOutputBytes: 1024 * 1024,
            out int maximumOutputBytes));
        Assert.Equal(1024 * 1024, maximumOutputBytes);

        Assert.False(OfficePngContainerValidator.TryGetCompressedMetadataOutputLimit(
            encodedBytes: 128L * 1024L * 1024L,
            compressedBytes: 128L * 1024L * 1024L,
            requestedMaximumOutputBytes: 1024 * 1024,
            out _));
    }

    [Fact]
    public void ApngSecondaryFrameValidationAccountsForTheRetainedDefaultPayload() {
        const long oneHundredTwentyMiB = 120L * 1024L * 1024L;
        const long eightyMiB = 80L * 1024L * 1024L;
        const long twentyMiB = 20L * 1024L * 1024L;
        const long fortyMiB = 40L * 1024L * 1024L;

        Assert.True(OfficePngAnimationValidator.IsFrameValidationWorkingSetWithinLimit(
            encodedBytes: oneHundredTwentyMiB,
            retainedPayloadBytes: 0L,
            compressedBytes: twentyMiB,
            validationWorkingSetBytes: eightyMiB,
            segmentCount: 1,
            paletteBytes: 0L));
        Assert.False(OfficePngAnimationValidator.IsFrameValidationWorkingSetWithinLimit(
            encodedBytes: oneHundredTwentyMiB,
            retainedPayloadBytes: fortyMiB,
            compressedBytes: twentyMiB,
            validationWorkingSetBytes: eightyMiB,
            segmentCount: 1,
            paletteBytes: 0L));
    }

    [Fact]
    public void SharedTiffDecoderAppliesPixelLimitOnlyToSelectedPage() {
        byte[] encoded = OfficeTiffCodec.EncodePages(new[] {
            new OfficeRasterImage(2, 1, OfficeColor.Red),
            new OfficeRasterImage(1, 2, OfficeColor.Blue)
        });
        int firstIfd = ReadLittleEndian(encoded, 4);
        SetTiffLongTag(encoded, firstIfd, 256, 1000);
        SetTiffLongTag(encoded, firstIfd, 257, 1000);
        var options = new OfficeRasterDecodeOptions {
            FrameIndex = 1,
            MaximumDecodedPixels = 4
        };

        Assert.True(OfficeRasterImageDecoder.TryDecode(
            encoded,
            options,
            out OfficeRasterImage? selected,
            out OfficeRasterDecodeInfo info));
        Assert.Equal((1, 2), (selected!.Width, selected.Height));
        Assert.Equal(1, info.SelectedFrameIndex);
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
        clone.WriteResolutionMetadata = false;
        clone.Png.WritePhysicalResolution = false;
        clone.Tiff.Compression = OfficeTiffCompression.None;
        clone.Tiff.Predictor = OfficeTiffPredictor.None;
        clone.Tiff.WriteResolution = false;

        Assert.Equal(85, source.Jpeg.Quality);
        Assert.True(source.WriteResolutionMetadata);
        Assert.True(source.Png.WritePhysicalResolution);
        Assert.Equal(OfficeTiffCompression.PackBits, source.Tiff.Compression);
        Assert.Equal(OfficeTiffPredictor.Horizontal, source.Tiff.Predictor);
        Assert.True(source.Tiff.WriteResolution);
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

    private static int ReadUInt16LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8;

    private static void SetTiffShortTag(byte[] bytes, int ifdOffset, int expectedTag, int value) {
        int entryCount = ReadUInt16LittleEndian(bytes, ifdOffset);
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            if (ReadUInt16LittleEndian(bytes, entryOffset) != expectedTag) continue;
            bytes[entryOffset + 8] = (byte)value;
            bytes[entryOffset + 9] = (byte)(value >> 8);
            return;
        }
        throw new InvalidOperationException("TIFF entry was not found.");
    }

    private static byte[] CreateSingleEntryIcon(byte[] payload) {
        var icon = new byte[22 + payload.Length];
        icon[2] = 1;
        icon[4] = 1;
        icon[6] = 1;
        icon[7] = 1;
        icon[10] = 1;
        icon[12] = 32;
        WriteLittleEndian(icon, 14, payload.Length);
        WriteLittleEndian(icon, 18, 22);
        Buffer.BlockCopy(payload, 0, icon, 22, payload.Length);
        return icon;
    }

    private static byte[] DuplicateFirstJpegFrameSegment(byte[] jpeg) {
        for (int offset = 2; offset + 3 < jpeg.Length; offset++) {
            if (jpeg[offset] != 0xFF || (jpeg[offset + 1] != 0xC0 && jpeg[offset + 1] != 0xC2)) continue;
            int segmentLength = jpeg[offset + 2] << 8 | jpeg[offset + 3];
            int totalLength = segmentLength + 2;
            if (segmentLength < 2 || offset + totalLength > jpeg.Length) break;
            var result = new byte[jpeg.Length + totalLength];
            Buffer.BlockCopy(jpeg, 0, result, 0, offset + totalLength);
            Buffer.BlockCopy(jpeg, offset, result, offset + totalLength, totalLength);
            Buffer.BlockCopy(
                jpeg,
                offset + totalLength,
                result,
                offset + totalLength * 2,
                jpeg.Length - offset - totalLength);
            return result;
        }
        throw new InvalidOperationException("JPEG frame segment was not found.");
    }

    private static byte[] RemoveLastJpegFrameComponent(byte[] jpeg) {
        int markerOffset = FindJpegMarker(jpeg, 0xC0);
        int segmentLength = jpeg[markerOffset + 2] << 8 | jpeg[markerOffset + 3];
        int componentCountOffset = markerOffset + 9;
        Assert.Equal(3, jpeg[componentCountOffset]);
        var result = new byte[jpeg.Length - 3];
        int removedComponentOffset = markerOffset + 2 + segmentLength - 3;
        Buffer.BlockCopy(jpeg, 0, result, 0, removedComponentOffset);
        Buffer.BlockCopy(
            jpeg,
            removedComponentOffset + 3,
            result,
            removedComponentOffset,
            jpeg.Length - removedComponentOffset - 3);
        result[markerOffset + 2] = (byte)((segmentLength - 3) >> 8);
        result[markerOffset + 3] = (byte)(segmentLength - 3);
        result[componentCountOffset] = 2;
        return result;
    }

    private static byte[] InsertJpegExifSegmentsAfterStart(
        byte[] jpeg,
        ushort firstOrientation,
        ushort lastOrientation) {
        byte[] first = CreateJpegExifSegment(firstOrientation);
        byte[] last = CreateJpegExifSegment(lastOrientation);
        var result = new byte[jpeg.Length + first.Length + last.Length];
        Buffer.BlockCopy(jpeg, 0, result, 0, 2);
        Buffer.BlockCopy(first, 0, result, 2, first.Length);
        Buffer.BlockCopy(last, 0, result, 2 + first.Length, last.Length);
        Buffer.BlockCopy(jpeg, 2, result, 2 + first.Length + last.Length, jpeg.Length - 2);
        return result;
    }

    private static byte[] CreateJpegExifSegment(ushort orientation) {
        byte[] tiff = CreateExifOrientation(orientation);
        int segmentLength = checked(tiff.Length + 8);
        var segment = new byte[segmentLength + 2];
        segment[0] = 0xFF;
        segment[1] = 0xE1;
        segment[2] = (byte)(segmentLength >> 8);
        segment[3] = (byte)segmentLength;
        System.Text.Encoding.ASCII.GetBytes("Exif\0\0").CopyTo(segment, 4);
        Buffer.BlockCopy(tiff, 0, segment, 10, tiff.Length);
        return segment;
    }

    private static byte[] CreateStaticLossyWebpHeader() {
        var webp = new byte[30];
        System.Text.Encoding.ASCII.GetBytes("RIFF").CopyTo(webp, 0);
        WriteLittleEndian(webp, 4, webp.Length - 8);
        System.Text.Encoding.ASCII.GetBytes("WEBPVP8 ").CopyTo(webp, 8);
        WriteLittleEndian(webp, 16, 10);
        webp[23] = 0x9D;
        webp[24] = 0x01;
        webp[25] = 0x2A;
        webp[26] = 1;
        webp[28] = 1;
        return webp;
    }

    private static int FindJpegMarker(byte[] jpeg, byte marker) {
        for (int offset = 2; offset + 1 < jpeg.Length; offset++) {
            if (jpeg[offset] == 0xFF && jpeg[offset + 1] == marker) return offset;
        }
        throw new InvalidOperationException($"JPEG marker 0x{marker:X2} was not found.");
    }

    private static byte[] CreateExifOrientation(ushort orientation) => new byte[] {
        (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
        0x01, 0x00,
        0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00,
        (byte)orientation, (byte)(orientation >> 8), 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00
    };

    private static void SetTiffLongTag(byte[] bytes, int ifdOffset, int expectedTag, int value) {
        int entryCount = ReadUInt16LittleEndian(bytes, ifdOffset);
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            if (ReadUInt16LittleEndian(bytes, entryOffset) != expectedTag) continue;
            WriteLittleEndian(bytes, entryOffset + 8, value);
            return;
        }
        throw new InvalidOperationException("TIFF entry was not found.");
    }

    private static int ReadTiffLongTag(byte[] bytes, int ifdOffset, int expectedTag) {
        int entryCount = ReadUInt16LittleEndian(bytes, ifdOffset);
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            if (ReadUInt16LittleEndian(bytes, entryOffset) == expectedTag) {
                return ReadLittleEndian(bytes, entryOffset + 8);
            }
        }
        throw new InvalidOperationException("TIFF entry was not found.");
    }

    private static void SetTiffTag(byte[] bytes, int ifdOffset, int expectedTag, int replacementTag) {
        int entryCount = ReadUInt16LittleEndian(bytes, ifdOffset);
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            if (ReadUInt16LittleEndian(bytes, entryOffset) != expectedTag) continue;
            bytes[entryOffset] = (byte)replacementTag;
            bytes[entryOffset + 1] = (byte)(replacementTag >> 8);
            return;
        }
        throw new InvalidOperationException("TIFF entry was not found.");
    }

    private static OfficeRasterImage CreateImageFromVp8lResiduals(int width, int height, uint[] residuals) {
        var image = new OfficeRasterImage(width, height);
        var pixels = new uint[residuals.Length];
        for (int position = 0; position < residuals.Length; position++) {
            int x = position % width;
            int y = position / width;
            uint predictor = position == 0 ? 0xFF000000U : x == 0 ? pixels[position - width] : pixels[position - 1];
            uint transformed = residuals[position];
            int green = (int)(transformed >> 8) & 255;
            int red = (((int)(transformed >> 16) & 255) + green) & 255;
            int blue = (((int)transformed & 255) + green) & 255;
            int alpha = (int)(transformed >> 24) & 255;
            uint color = AddArgb(predictor, alpha, red, green, blue);
            pixels[position] = color;
            image.SetPixel(x, y, OfficeColor.FromRgba(
                (byte)(color >> 16),
                (byte)(color >> 8),
                (byte)color,
                (byte)(color >> 24)));
        }
        return image;
    }

    private static uint AddArgb(uint predictor, int alpha, int red, int green, int blue) =>
        (uint)((((int)(predictor >> 24) + alpha) & 255) << 24 |
               (((int)(predictor >> 16) + red) & 255) << 16 |
               (((int)(predictor >> 8) + green) & 255) << 8 |
               (((int)predictor + blue) & 255));

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

    private static byte[] ReadTiffStrip(byte[] bytes) {
        int stripOffset = ReadLittleEndian(bytes, FindClassicTiffEntry(bytes, 273) + 8);
        int stripLength = ReadLittleEndian(bytes, FindClassicTiffEntry(bytes, 279) + 8);
        var strip = new byte[stripLength];
        Buffer.BlockCopy(bytes, stripOffset, strip, 0, stripLength);
        return strip;
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

    private static byte[] CreateUniformVp8lFixture(int width, int height) {
        var writer = new TestLsbBitWriter();
        writer.WriteBits((uint)(width - 1), 14);
        writer.WriteBits((uint)(height - 1), 14);
        writer.WriteBits(0, 1);  // no alpha hint
        writer.WriteBits(0, 3);  // version
        writer.WriteBits(0, 1);  // no transforms
        writer.WriteBits(0, 1);  // no color cache
        writer.WriteBits(0, 1);  // one Huffman group
        WriteSingleSymbolTree(writer, 0);
        WriteSingleSymbolTree(writer, 11);
        WriteSingleSymbolTree(writer, 22);
        WriteSingleSymbolTree(writer, 255);
        WriteSingleSymbolTree(writer, 0);

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

    private static byte[] CreateVp8lSparseHuffmanGroupFixture() {
        var writer = new TestLsbBitWriter();
        writer.WriteBits(0, 14); // width - 1
        writer.WriteBits(0, 14); // height - 1
        writer.WriteBits(0, 1);  // no alpha hint
        writer.WriteBits(0, 3);  // version
        writer.WriteBits(0, 1);  // no transforms
        writer.WriteBits(0, 1);  // no color cache
        writer.WriteBits(1, 1);  // meta Huffman codes
        writer.WriteBits(0, 3);  // prefix bits = 2

        writer.WriteBits(0, 1);  // prefix image has no color cache
        WriteSingleSymbolTree(writer, 1);   // group identifier 1 in green
        WriteSingleSymbolTree(writer, 0);
        WriteSingleSymbolTree(writer, 0);
        WriteSingleSymbolTree(writer, 255);
        WriteSingleSymbolTree(writer, 0);

        for (int group = 0; group < 2; group++) {
            WriteSingleSymbolTree(writer, 0);
            WriteSingleSymbolTree(writer, 11);
            WriteSingleSymbolTree(writer, 22);
            WriteSingleSymbolTree(writer, 255);
            WriteSingleSymbolTree(writer, 0);
        }

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
        writer.WriteBits(1, 2); // three encoded instructions; repeat expands to five code lengths
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
