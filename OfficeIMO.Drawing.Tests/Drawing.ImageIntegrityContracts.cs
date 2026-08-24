using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void PngReaderAndExportResultRejectInvalidChunkCrc() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        png[29] ^= 0x01;

        Assert.False(OfficePngReader.TryGetFrameCount(png, out _));
        Assert.False(OfficePngReader.TryDecode(png, out _));
        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png));
    }

    [Theory]
    [InlineData("<SVG xmlns='http://www.w3.org/2000/svg' width='1' height='1'/>")]
    [InlineData("<svg xmlns='urn:not-svg' width='1' height='1'/>")]
    public void CompleteContentValidationRejectsIncorrectlyQualifiedSvgRoots(string markup) {
        byte[] svg = System.Text.Encoding.UTF8.GetBytes(markup);

        Assert.True(OfficeImageReader.TryIdentify(svg, "invalid.svg", out _));
        Assert.False(OfficeImageReader.TryIdentifyByContent(svg, "invalid.svg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(svg, "invalid.svg", out _));
        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(OfficeImageExportFormat.Svg, 1, 1, svg));
    }

    [Fact]
    public void CompleteContentValidationParsesSvgBeyondTheMetadataPrefix() {
        string markup = "<!--" + new string('x', 5000) + "-->" +
                        "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='1'/>";
        byte[] svg = System.Text.Encoding.UTF8.GetBytes(markup);

        Assert.True(OfficeImageReader.TryIdentifyByContent(svg, "misleading.png", out OfficeImageInfo info));
        Assert.Equal(OfficeImageFormat.Svg, info.Format);
        Assert.True(OfficeImageReader.TryValidateContent(svg, "misleading.png", out _));
    }

    [Fact]
    public void PngReaderAndExportResultRejectInvalidZlibChecksumWithValidChunkCrc() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int idatOffset = FindPngChunk(png, "IDAT");
        int length = ReadBigEndianInt32(png, idatOffset);
        png[idatOffset + 8 + length - 1] ^= 0x01;
        WritePngChunkCrc(png, idatOffset, length);

        Assert.True(OfficePngReader.TryGetFrameCount(png, out _));
        Assert.False(OfficePngReader.TryDecode(png, out _));
        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png));
    }

    [Fact]
    public void CompleteContentValidationRejectsTrailingBytesInsidePngAndApngZlibStreams() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] pngWithTrailingDeflateByte = InsertByteBeforePngChunkChecksum(png, "IDAT");
        byte[] apng = CreateTwoFrameApng(png);
        byte[] apngWithTrailingFrameDeflateByte = InsertByteBeforePngChunkChecksum(apng, "fdAT");

        Assert.True(OfficePngReader.TryGetFrameCount(pngWithTrailingDeflateByte, out _));
        Assert.True(OfficePngReader.TryGetFrameCount(apngWithTrailingFrameDeflateByte, out _));
        Assert.False(OfficeImageReader.TryValidateContent(pngWithTrailingDeflateByte, "trailing.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(apngWithTrailingFrameDeflateByte, "trailing-apng.png", out _));
    }

    [Fact]
    public void ManagedPngReaderDecodesStructurallyValidAdam7Png() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        png[28] = 1;
        WritePngChunkCrc(png, 8, 13);

        Assert.True(OfficePngReader.TryGetFrameCount(png, out int frameCount));
        Assert.Equal(1, frameCount);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.White, image!.GetPixel(0, 0));
        var result = new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png);
        Assert.Equal(png, result.Bytes);
    }

    [Fact]
    public void ManagedPngReaderMapsEveryAdam7PassToTheLogicalCanvas() {
        var source = new OfficeRasterImage(9, 9);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgba(
                    (byte)(x * 23), (byte)(y * 23), (byte)(x * 11 + y), (byte)(255 - x - y)));
            }
        }
        byte[] png = CreateAdam7RgbaPng(source);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? decoded));
        Assert.Equal(source.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void PngContainerRejectsUnknownCriticalAndMisplacedPaletteChunks() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] unknownCritical = InsertPngChunkBefore(png, "IDAT", "ABCD", Array.Empty<byte>());
        byte[] misplacedPalette = InsertPngChunkBefore(png, "IEND", "PLTE", new byte[] { 0, 0, 0 });

        Assert.False(OfficePngReader.TryGetFrameCount(unknownCritical, out _));
        Assert.False(OfficePngReader.TryGetFrameCount(misplacedPalette, out _));
    }

    [Fact]
    public void PngContainerRejectsNonContiguousImageDataChunks() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] split = SplitPngImageDataWithAncillaryChunk(png);

        Assert.False(OfficePngReader.TryGetFrameCount(split, out _));
        Assert.False(OfficeImageReader.TryValidateContent(split, "split.png", out _));
    }

    [Fact]
    public void PngContainerRejectsTransparencySamplesOutsideTheBitDepthRange() {
        byte[] grayscale = OfficePngWriter.EncodeScanlines(
            1, 1, 1, 0, new byte[] { 0, 0 });
        byte[] validGrayscale = InsertPngChunkBefore(grayscale, "IDAT", "tRNS", new byte[] { 0, 1 });
        byte[] invalidGrayscale = InsertPngChunkBefore(grayscale, "IDAT", "tRNS", new byte[] { 0, 2 });
        byte[] truecolor = OfficePngWriter.EncodeScanlines(
            1, 1, 8, 2, new byte[] { 0, 255, 0, 0 });
        byte[] invalidTruecolor = InsertPngChunkBefore(
            truecolor, "IDAT", "tRNS", new byte[] { 1, 0, 0, 0, 0, 0 });

        Assert.True(OfficeImageReader.TryValidateContent(validGrayscale, "valid-trns.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidGrayscale, "invalid-gray-trns.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidTruecolor, "invalid-rgb-trns.png", out _));
    }

    [Fact]
    public void PngContainerRequiresOneWellFormedPhysicalDimensionsChunkBeforeImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] physicalDimensions = { 0, 0, 0x0E, 0xC4, 0, 0, 0x0E, 0xC4, 1 };
        byte[] withPhysicalDimensions = InsertPngChunkBefore(png, "IDAT", "pHYs", physicalDimensions);
        byte[] wrongLength = InsertPngChunkBefore(png, "IDAT", "pHYs", new byte[8]);
        byte[] duplicate = InsertPngChunkBefore(withPhysicalDimensions, "IDAT", "pHYs", physicalDimensions);
        byte[] misplaced = InsertPngChunkBefore(png, "IEND", "pHYs", physicalDimensions);
        byte[] invalidUnit = InsertPngChunkBefore(
            png, "IDAT", "pHYs", new byte[] { 0, 0, 0x0E, 0xC4, 0, 0, 0x0E, 0xC4, 2 });

        Assert.True(OfficeImageReader.TryValidateContent(png, "valid-phys.png", out _));
        Assert.True(OfficeImageReader.TryValidateContent(withPhysicalDimensions, "valid-phys.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(wrongLength, "wrong-length-phys.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-phys.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(misplaced, "misplaced-phys.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidUnit, "invalid-unit-phys.png", out _));
    }

    [Fact]
    public void PngContainerRequiresOneWellFormedStandardRgbChunkBeforePaletteAndImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] withStandardRgb = InsertPngChunkBefore(png, "IDAT", "sRGB", new byte[] { 0 });
        byte[] wrongLength = InsertPngChunkBefore(png, "IDAT", "sRGB", Array.Empty<byte>());
        byte[] duplicate = InsertPngChunkBefore(withStandardRgb, "IDAT", "sRGB", new byte[] { 1 });
        byte[] misplaced = InsertPngChunkBefore(png, "IEND", "sRGB", new byte[] { 0 });
        byte[] invalidIntent = InsertPngChunkBefore(png, "IDAT", "sRGB", new byte[] { 4 });

        Assert.True(OfficeImageReader.TryValidateContent(withStandardRgb, "valid-srgb.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(wrongLength, "wrong-length-srgb.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-srgb.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(misplaced, "misplaced-srgb.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidIntent, "invalid-intent-srgb.png", out _));
    }

    [Fact]
    public void CompleteContentValidationHonorsBmpDeclaredFileSizeAndReservedFields() {
        byte[] bmp = new byte[58];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteInt32LittleEndian(bmp, 2, bmp.Length);
        WriteInt32LittleEndian(bmp, 10, 54);
        WriteInt32LittleEndian(bmp, 14, 40);
        WriteInt32LittleEndian(bmp, 18, 1);
        WriteInt32LittleEndian(bmp, 22, 1);
        WriteUInt16LittleEndian(bmp, 26, 1);
        WriteUInt16LittleEndian(bmp, 28, 24);
        bmp[56] = 255;
        byte[] oversizedDeclaration = (byte[])bmp.Clone();
        WriteInt32LittleEndian(oversizedDeclaration, 2, bmp.Length + 1);
        byte[] undersizedDeclaration = (byte[])bmp.Clone();
        WriteInt32LittleEndian(undersizedDeclaration, 2, bmp.Length - 1);
        byte[] reservedField = (byte[])bmp.Clone();
        reservedField[6] = 1;
        byte[] inventedDibHeader = new byte[59];
        Buffer.BlockCopy(bmp, 0, inventedDibHeader, 0, 54);
        WriteInt32LittleEndian(inventedDibHeader, 2, inventedDibHeader.Length);
        WriteInt32LittleEndian(inventedDibHeader, 10, 55);
        WriteInt32LittleEndian(inventedDibHeader, 14, 41);
        inventedDibHeader[57] = 255;

        Assert.True(OfficeImageReader.TryValidateContent(bmp, "valid.bmp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(oversizedDeclaration, "truncated.bmp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(undersizedDeclaration, "trailing.bmp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(reservedField, "reserved.bmp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(inventedDibHeader, "invented-header.bmp", out _));
    }

    [Fact]
    public void CompleteContentValidationRejectsTruncatedGifCorruptPngAndMarkerOnlyJpeg() {
        byte[] truncatedGif = { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 1, 0, 1, 0, 0, 0, 0 };
        byte[] markerOnlyJpeg = {
            0xFF, 0xD8,
            0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0xFF, 0xD9
        };
        byte[] corruptPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int idatOffset = FindPngChunk(corruptPng, "IDAT");
        int length = ReadBigEndianInt32(corruptPng, idatOffset);
        corruptPng[idatOffset + 8 + length - 1] ^= 0x01;
        WritePngChunkCrc(corruptPng, idatOffset, length);

        Assert.True(OfficeImageReader.TryIdentifyByContent(truncatedGif, "truncated.gif", out _));
        Assert.False(OfficeImageReader.TryValidateContent(truncatedGif, "truncated.gif", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(corruptPng, "corrupt.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(corruptPng, "corrupt.png", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(markerOnlyJpeg, "marker-only.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(markerOnlyJpeg, "marker-only.jpg", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksEveryApngFrameAndJpegScan() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);
        Assert.True(OfficeImageReader.TryValidateContent(apng, "animated.png", out _));
        Assert.True(OfficePngReader.TryGetFrameCount(apng, out int frameCount));
        Assert.Equal(2, frameCount);

        byte[] corruptApng = (byte[])apng.Clone();
        int frameDataOffset = FindPngChunk(corruptApng, "fdAT");
        int frameDataLength = ReadBigEndianInt32(corruptApng, frameDataOffset);
        corruptApng[frameDataOffset + 8 + frameDataLength - 1] ^= 0x01;
        WritePngChunkCrc(corruptApng, frameDataOffset, frameDataLength);
        Assert.True(OfficeRasterContainerInspector.TryInspect(corruptApng, out OfficeRasterContainerInfo? corruptInventory));
        Assert.Equal(2, corruptInventory!.Count);
        Assert.False(OfficeImageReader.TryValidateContent(corruptApng, "animated.png", out _));

        byte[] jpegWithEmptyFinalScan = {
            0xFF, 0xD8,
            0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0x01,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0xFF, 0xD9
        };
        Assert.True(OfficeImageReader.TryIdentifyByContent(jpegWithEmptyFinalScan, "progressive.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(jpegWithEmptyFinalScan, "progressive.jpg", out _));
    }

    [Fact]
    public void ApngSecondaryFrameValidationObservesCancellation() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(png);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficePngAnimationValidator.TryValidateAdditionalFrames(apng, cancellation.Token));
    }

    [Fact]
    public void WideSingleRowApngCompositionObservesCancellationInsideTheRow() {
        const int width = 8 * 1024 * 1024;
        var canvas = new OfficeRasterImage(width, 1, OfficeColor.Transparent);
        var frameImage = new OfficeRasterImage(width, 1, OfficeColor.FromRgba(32, 96, 224, 128));
        var frame = new OfficeRasterFrameInfo(
            0,
            OfficeRasterFrameKind.AnimationFrame,
            width,
            1,
            0,
            0,
            TimeSpan.Zero,
            OfficeRasterFrameDisposal.None,
            OfficeRasterFrameBlend.Over,
            isDefaultImage: false);
        using var cancellation = new CancellationTokenSource();
        Exception? workerException = null;
        var worker = new Thread(() => {
            try {
                OfficeApngDecoder.Composite(canvas, frameImage, frame, cancellation.Token);
            } catch (Exception exception) {
                Volatile.Write(ref workerException, exception);
            }
        }) { IsBackground = true };
        worker.Start();

        try {
            Assert.True(SpinWait.SpinUntil(
                () => Volatile.Read(ref canvas.PixelBuffer[(32 * 4) + 3]) != 0,
                TimeSpan.FromSeconds(5)),
                "APNG composition did not begin within the bounded wait.");
            cancellation.Cancel();
        } finally {
            Assert.True(worker.Join(TimeSpan.FromSeconds(5)));
        }

        Assert.IsType<OperationCanceledException>(Volatile.Read(ref workerException));
    }

    [Fact]
    public void RasterDecoderComposesExplicitlySelectedApngFrames() {
        OfficeColor expected = OfficeColor.FromRgba(32, 96, 224, 128);
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, expected));
        byte[] apng = CreateTwoFrameApng(staticPng);
        var options = new OfficeRasterDecodeOptions { FrameIndex = 1 };

        Assert.True(OfficeRasterImageDecoder.TryDecode(apng, options, out OfficeRasterImage? image, out OfficeRasterDecodeInfo info));
        Assert.NotNull(image);
        Assert.Equal(expected, image!.GetPixel(0, 0));
        Assert.Equal(2, info.FrameCount);
        Assert.Equal(1, info.SelectedFrameIndex);
        Assert.True(info.IsAnimated);
        Assert.True(info.FramesOrPagesDiscarded);
        Assert.Equal(OfficeRasterFrameBlend.Source, info.SelectedFrame!.Blend);
    }

    [Fact]
    public void CompleteContentValidationRejectsUndecodableJpegBmpTiffAndWebpBodies() {
        byte[] jpegWithoutTables = {
            0xFF, 0xD8,
            0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0x01, 0xFF, 0xD9
        };
        byte[] bmpWithoutPixels = CreateBmpInfoHeader(24, 0, height: 2);
        byte[] tiffWithoutCompleteStrip = OfficeTiffCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        Array.Resize(ref tiffWithoutCompleteStrip, tiffWithoutCompleteStrip.Length - 1);
        byte[] webpHeaderOnly = {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 18, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P',
            (byte)'V', (byte)'P', (byte)'8', (byte)'L', 5, 0, 0, 0,
            0x2F, 0, 0, 0, 0, 0
        };

        Assert.True(OfficeImageReader.TryIdentifyByContent(jpegWithoutTables, "missing-tables.jpg", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(bmpWithoutPixels, "missing-pixels.bmp", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(tiffWithoutCompleteStrip, "missing-strip.tiff", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(webpHeaderOnly, "header-only.webp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(jpegWithoutTables, "missing-tables.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(bmpWithoutPixels, "missing-pixels.bmp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(tiffWithoutCompleteStrip, "missing-strip.tiff", out _));
        Assert.False(OfficeImageReader.TryValidateContent(webpHeaderOnly, "header-only.webp", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksEveryIconEntryBody() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] dib = CreateOnePixelIconDib();
        byte[] pngIcon = CreateIcon(png);
        byte[] dibIcon = CreateIcon(dib);
        byte[] iconWithInvalidOnlyEntry = CreateIcon(new byte[] { 0x01 });
        byte[] iconWithInvalidSecondEntry = CreateIcon(png, new byte[] { 0x01 });

        Assert.True(OfficeImageReader.TryValidateContent(pngIcon, "valid-png.ico", out _));
        Assert.True(OfficeImageReader.TryValidateContent(dibIcon, "valid-dib.ico", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(iconWithInvalidOnlyEntry, "invalid.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(iconWithInvalidOnlyEntry, "invalid.ico", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(iconWithInvalidSecondEntry, "invalid.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(iconWithInvalidSecondEntry, "invalid.ico", out _));
    }

    [Fact]
    public void CompleteContentValidationCachesRepeatedIconPayloads() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] paddedPng = InsertPngChunkBefore(png, "IDAT", "vpAg", new byte[128 * 1024]);
        byte[] icon = CreateIconWithSharedPayload(paddedPng, ushort.MaxValue);

        Assert.True(OfficeImageReader.TryValidateContent(icon, "shared-payload.ico", out _));
    }

    [Fact]
    public async Task AsyncBatchPassesTheOperationDeadlineToConsumerCallbacks() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var options = new OfficeImageExportOptions { RenderTimeout = TimeSpan.FromMilliseconds(250) };

        await Assert.ThrowsAsync<OfficeImageExportTimeoutException>(() =>
            OfficeImageExportBatchProcessor.RunAsync(
                options,
                (accept, _) => accept(CreateResult("timeout", png), CancellationToken.None),
                async (_, callbackToken) => await Task.Delay(Timeout.InfiniteTimeSpan, callbackToken)));
    }

    [Fact]
    public void CompleteContentValidationRejectsInvalidApngSequenceCountAndFrameBounds() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);

        byte[] invalidSequence = (byte[])apng.Clone();
        int frameDataOffset = FindPngChunk(invalidSequence, "fdAT");
        WriteBigEndianInt32(invalidSequence, frameDataOffset + 8, 7);
        WritePngChunkCrc(invalidSequence, frameDataOffset, ReadBigEndianInt32(invalidSequence, frameDataOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidSequence, "sequence.png", out _));

        byte[] invalidCount = (byte[])apng.Clone();
        int animationControlOffset = FindPngChunk(invalidCount, "acTL");
        WriteBigEndianInt32(invalidCount, animationControlOffset + 8, 3);
        WritePngChunkCrc(invalidCount, animationControlOffset, ReadBigEndianInt32(invalidCount, animationControlOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidCount, "count.png", out _));

        byte[] invalidBounds = (byte[])apng.Clone();
        int firstFrameControlOffset = FindPngChunk(invalidBounds, "fcTL");
        int secondFrameControlOffset = FindPngChunk(
            invalidBounds,
            "fcTL",
            firstFrameControlOffset + 12 + ReadBigEndianInt32(invalidBounds, firstFrameControlOffset));
        WriteBigEndianInt32(invalidBounds, secondFrameControlOffset + 12, 2);
        WritePngChunkCrc(invalidBounds, secondFrameControlOffset, ReadBigEndianInt32(invalidBounds, secondFrameControlOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidBounds, "bounds.png", out _));

        byte[] invalidFirstDisposal = (byte[])apng.Clone();
        int disposalFrameControlOffset = FindPngChunk(invalidFirstDisposal, "fcTL");
        invalidFirstDisposal[disposalFrameControlOffset + 8 + 24] = 3;
        WritePngChunkCrc(
            invalidFirstDisposal,
            disposalFrameControlOffset,
            ReadBigEndianInt32(invalidFirstDisposal, disposalFrameControlOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidFirstDisposal, "first-disposal.png", out _));
    }

    [Theory]
    [InlineData(0x20)]
    [InlineData(0x08)]
    [InlineData(0x04)]
    public void AnimatedWebpRequiresEveryDeclaredMetadataChunk(int featureFlag) {
        byte[] animated = Convert.FromBase64String(
            "UklGRoQAAABXRUJQVlA4WAoAAAACAAAAAQAAAQAAQU5JTQYAAAAAAAAAAABBTk1GKAAAAAAAAAAAAAEAAAEAAGQAAAJWUDhMDwAAAC8BQAAABxD9j/4HIqL/AQBBTk1GKAAAAAAAAAAAAAEAAAEAAGQAAABWUDhMDwAAAC8BQAAABxDR//4HIqL/AQA=");
        byte[] inconsistent = (byte[])animated.Clone();
        inconsistent[20] |= (byte)featureFlag;

        Assert.True(OfficeImageReader.TryIdentifyByContent(animated, "animated.webp", out _));
        Assert.False(OfficeImageReader.TryIdentifyByContent(inconsistent, "inconsistent.webp", out _));
    }

    [Theory]
    [InlineData(0x01)]
    [InlineData(0x20)]
    [InlineData(0x40)]
    public void WebpAlphaChunksRejectReservedHeaderValues(int invalidControl) {
        byte[] valid = CreateAlphaWebp(control: 0);
        byte[] invalid = CreateAlphaWebp((byte)invalidControl);

        Assert.True(OfficeImageReader.TryIdentifyByContent(valid, "valid-alpha.webp", out _));
        Assert.False(OfficeImageReader.TryIdentifyByContent(invalid, "invalid-alpha.webp", out _));
    }

    [Fact]
    public void WebpValidationRequiresStructurallyValidExifMetadata() {
        byte[] valid = OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(1, 1, OfficeColor.White),
            OfficeImageExportFormat.Webp,
            new OfficeRasterEncodingOptions { DpiX = 144D, DpiY = 120D });
        int exifOffset = FindWebpChunk(valid, "EXIF");
        byte[] malformed = (byte[])valid.Clone();
        malformed[exifOffset + 8] = (byte)'X';

        Assert.True(OfficeImageReader.TryValidateContent(valid, "valid-exif.webp", out _));
        Assert.False(OfficeImageReader.TryIdentifyByContent(malformed, "malformed-exif.webp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "malformed-exif.webp", out _));
    }

    [Fact]
    public void WebpValidationAcceptsStructurallyValidExifWithoutImageDimensionTags() {
        byte[] simple = OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] exif = {
            (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
            0, 0,
            0, 0, 0, 0
        };
        var bytes = new List<byte> {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P'
        };
        bytes.AddRange(CreateWebpChunk("VP8X", new byte[] { 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0 }));
        bytes.AddRange(simple.Skip(12));
        bytes.AddRange(CreateWebpChunk("EXIF", exif));
        byte[] extended = bytes.ToArray();
        WriteInt32LittleEndian(extended, 4, extended.Length - 8);

        Assert.True(OfficeImageReader.TryIdentifyByContent(extended, "metadata-only-exif.webp", out _));
        Assert.True(OfficeImageReader.TryValidateContent(extended, "metadata-only-exif.webp", out _));
    }

    [Fact]
    public void ApngValidationBoundsAggregateDecodedFramePixels() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);
        const int frameWidth = 30000000;

        int ihdrOffset = FindPngChunk(apng, "IHDR");
        WriteBigEndianInt32(apng, ihdrOffset + 8, frameWidth);
        int firstFrameControlOffset = FindPngChunk(apng, "fcTL");
        int secondFrameControlOffset = FindPngChunk(
            apng,
            "fcTL",
            firstFrameControlOffset + 12 + ReadBigEndianInt32(apng, firstFrameControlOffset));
        WriteBigEndianInt32(apng, firstFrameControlOffset + 12, frameWidth);
        WriteBigEndianInt32(apng, secondFrameControlOffset + 12, frameWidth);

        Assert.True(OfficePngAnimationValidator.TryValidateStructure(apng));
        Assert.False(OfficePngAnimationValidator.TryValidateAdditionalFrames(apng));
    }

    [Fact]
    public void ApngEarlyFrameSelectionDoesNotChargeUnselectedFramePixels() {
        byte[] staticPng = OfficePngWriter.Encode(
            new OfficeRasterImage(1000, 1000, OfficeColor.White));
        byte[] apng = CreateRepeatedFrameApng(staticPng, frameCount: 51);

        Assert.True(OfficeRasterContainerInspector.TryInspect(
            apng, out OfficeRasterContainerInfo? container));
        Assert.Equal(51, container!.Count);
        Assert.True(OfficeRasterImageDecoder.TryDecode(
            apng,
            new OfficeRasterDecodeOptions { FrameIndex = 0 },
            out OfficeRasterImage? selected,
            out OfficeRasterDecodeInfo info));

        Assert.Equal((1000, 1000), (selected!.Width, selected.Height));
        Assert.Equal(0, info.SelectedFrameIndex);
        Assert.Equal(51, info.FrameCount);
        Assert.False(OfficeRasterImageDecoder.TryDecode(
            apng,
            new OfficeRasterDecodeOptions { FrameIndex = 50 },
            out _,
            out _));
    }

    [Fact]
    public void PngContainerTreatsApngFrameControlAsEndingTheIdatRun() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);
        byte[] nonContiguous = InsertPngChunkBefore(apng, "fdAT", "IDAT", Array.Empty<byte>());

        Assert.False(OfficePngReader.TryGetFrameCount(nonContiguous, out _));
        Assert.False(OfficeImageReader.TryValidateContent(nonContiguous, "non-contiguous-apng.png", out _));
    }

    [Fact]
    public void CompleteContentValidationAllowsEmptyApngFrameDataSegments() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);
        byte[] emptyFrameData = new byte[4];
        WriteBigEndianInt32(emptyFrameData, 0, 3);
        byte[] split = InsertPngChunkBefore(apng, "IEND", "fdAT", emptyFrameData);

        Assert.True(OfficeImageReader.TryValidateContent(split, "empty-frame-segment.png", out _));
    }

    [Fact]
    public async Task GuardedAsyncConsumerSerializesConcurrentAdmissionAndSequenceAssignment() {
        const int maximum = 300;
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var results = new List<OfficeImageExportResult>();
        var options = new OfficeImageExportOptions { MaximumOutputCount = maximum };

        await Assert.ThrowsAsync<OfficeImageExportBatchLimitException>(() =>
            OfficeImageExportBatchProcessor.RunAsync(
                options,
                async (accept, token) => await Task.WhenAll(
                    Enumerable.Range(0, 500).Select(index => accept(
                        new OfficeImageExportResult(
                            OfficeImageExportFormat.Png,
                            1,
                            1,
                            png,
                            name: index.ToString()),
                        token))),
                (result, _) => {
                    results.Add(result);
                    return Task.CompletedTask;
                }));

        Assert.Equal(maximum, results.Count);
        Assert.Equal(Enumerable.Range(0, maximum), results.Select(result => result.SequenceIndex!.Value));
    }

    [Fact]
    public async Task GuardedAsyncConsumerAllowsConsumerReentry() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var observed = new ConcurrentQueue<int>();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 2 },
            async (result, token) => {
                observed.Enqueue(result.SequenceIndex!.Value);
                if (result.Name == "outer") {
                    await accept!(new OfficeImageExportResult(
                        OfficeImageExportFormat.Png,
                        1,
                        1,
                        png,
                        name: "inner"), token);
                }
            },
            cancellation.Token);

        await accept(new OfficeImageExportResult(
            OfficeImageExportFormat.Png,
            1,
            1,
            png,
            name: "outer"), cancellation.Token);

        Assert.Equal(new[] { 0, 1 }, observed);
    }

    [Fact]
    public async Task GuardedAsyncConsumerKeepsDiscardedSynchronousReentryInsideTheGate() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var innerStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseInner = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var laterStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 3 },
            async (result, token) => {
                if (result.Name == "outer") {
                    _ = accept!(CreateResult("inner", png), token);
                } else if (result.Name == "inner") {
                    innerStarted.TrySetResult(true);
                    await releaseInner.Task.ConfigureAwait(false);
                } else {
                    laterStarted.TrySetResult(true);
                }
            });

        Task outer = accept(CreateResult("outer", png), default);
        await innerStarted.Task;
        Assert.False(outer.IsCompleted);
        Task later = accept(CreateResult("later", png), default);
        Assert.False(laterStarted.Task.IsCompleted);

        releaseInner.TrySetResult(true);
        await outer;
        await later;
        Assert.Equal(TaskStatus.RanToCompletion, laterStarted.Task.Status);
    }

    [Fact]
    public async Task GuardedAsyncConsumerDrainsReentryAfterSynchronousCallbackFailure() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var innerStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseInner = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var laterStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 3 },
            (result, token) => {
                if (result.Name == "outer") {
                    _ = accept!(CreateResult("inner", png), token);
                    throw new InvalidOperationException("outer failed");
                }
                if (result.Name == "inner") {
                    innerStarted.TrySetResult(true);
                    return releaseInner.Task;
                }
                laterStarted.TrySetResult(true);
                return Task.CompletedTask;
            });

        Task outer = accept(CreateResult("outer", png), default);
        await innerStarted.Task;
        Assert.False(outer.IsCompleted);
        Task later = accept(CreateResult("later", png), default);
        Assert.False(laterStarted.Task.IsCompleted);

        releaseInner.TrySetResult(true);
        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() => outer);
        Assert.Equal("outer failed", failure.Message);
        await later;
        Assert.Equal(TaskStatus.RanToCompletion, laterStarted.Task.Status);
    }

    [Fact]
    public async Task GuardedAsyncConsumerPreservesAsyncCallbackFailureAfterNestedFailure() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 3 },
            async (result, token) => {
                if (result.Name == "outer") {
                    _ = accept!(CreateResult("inner", png), token);
                    await Task.Yield();
                    throw new InvalidOperationException("outer async failed");
                }
                if (result.Name == "inner") throw new ArgumentException("inner failed");
            });

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            accept(CreateResult("outer", png), default));
        Assert.Equal("outer async failed", failure.Message);

        await accept(CreateResult("later", png), default);
    }

    [Fact]
    public async Task GuardedAsyncConsumerRejectsForkedReentryWithoutConcurrentAdmission() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var observed = new ConcurrentQueue<int>();
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 3 },
            async (result, token) => {
                observed.Enqueue(result.SequenceIndex!.Value);
                if (result.Name == "outer") {
                    await Task.WhenAll(
                        Task.Run(() => accept!(CreateResult("fork-1", png), token), token),
                        Task.Run(() => accept!(CreateResult("fork-2", png), token), token));
                }
            });

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            accept(CreateResult("outer", png), default));
        Assert.Equal(new[] { 0 }, observed);
    }

    [Fact]
    public async Task GuardedAsyncConsumerRejectsDeferredReentryAfterCallbackCompletion() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        Task<Exception>? deferred = null;
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 2 },
            (result, token) => {
                if (result.Name == "outer") {
                    deferred = Task.Run(async () => {
                        await release.Task.ConfigureAwait(false);
                        return await Record.ExceptionAsync(() => accept!(CreateResult("late", png), token));
                    });
                }
                return Task.CompletedTask;
            });

        await accept(CreateResult("outer", png), default);
        release.TrySetResult(true);
        Exception error = Assert.IsType<InvalidOperationException>(await deferred!);
        Assert.Contains("deferred reentry", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GuardedAsyncConsumerSerializesSiblingSynchronousReentries() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var firstStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirst = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var secondStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var observed = new ConcurrentQueue<string>();
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 3 },
            async (result, token) => {
                observed.Enqueue(result.Name!);
                if (result.Name == "outer") {
                    Task first = accept!(CreateResult("first", png), token);
                    Task second = accept!(CreateResult("second", png), token);
                    await Task.WhenAll(first, second).ConfigureAwait(false);
                } else if (result.Name == "first") {
                    firstStarted.TrySetResult(true);
                    await releaseFirst.Task.ConfigureAwait(false);
                } else if (result.Name == "second") {
                    secondStarted.TrySetResult(true);
                }
            });

        Task outer = accept(CreateResult("outer", png), default);
        await firstStarted.Task;
        Assert.False(secondStarted.Task.IsCompleted);
        releaseFirst.TrySetResult(true);
        await outer;

        Assert.Equal(new[] { "outer", "first", "second" }, observed);
    }

    [Fact]
    public void GuardedConsumerSerializesConcurrentAdmissionAndSequenceAssignment() {
        const int maximum = 300;
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var results = new List<OfficeImageExportResult>();
        var options = new OfficeImageExportOptions { MaximumOutputCount = maximum };
        OfficeImageExportConsumer accept = OfficeImageExportBatchProcessor.CreateGuardedConsumer(
            options,
            result => results.Add(result));
        int rejected = 0;

        System.Threading.Tasks.Parallel.For(0, 500, index => {
            try {
                accept(new OfficeImageExportResult(
                    OfficeImageExportFormat.Png,
                    1,
                    1,
                    png,
                    name: index.ToString()));
            } catch (OfficeImageExportBatchLimitException) {
                Interlocked.Increment(ref rejected);
            }
        });

        Assert.Equal(maximum, results.Count);
        Assert.Equal(200, rejected);
        Assert.Equal(Enumerable.Range(0, maximum), results.Select(result => result.SequenceIndex!.Value));
    }

    [Fact]
    public void EffectiveScaleUsesTargetDpiWithoutRequiringValidationSideEffects() {
        var options = new OfficeImageExportOptions {
            Scale = 1D,
            TargetDpi = 192D
        };

        Assert.Equal(2D, options.GetEffectiveScale(100D, 100D));
        Assert.Equal(1D, options.Scale);
    }

    [Fact]
    public void ImageResultRejectsUndefinedFileConflictPolicyBeforeWriting() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var result = new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png);
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N") + ".png");

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            result.Save(path, (OfficeImageExportFileConflictPolicy)999));
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void StreamIdentificationRejectsOversizedSeekablePayloadBeforeReading() {
        using var stream = new DeclaredLengthStream(128L * 1024L * 1024L + 1L);

        Assert.False(OfficeImageReader.TryIdentifyByContent(stream, "oversized.png", out _));
        Assert.Equal(0, stream.ReadCount);
        Assert.Equal(0L, stream.Position);
    }

    [Fact]
    public void BoundedRasterDecodeSupportsSeekableAndForwardOnlyStreams() {
        OfficeColor expected = OfficeColor.FromRgba(12, 34, 56, 178);
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(3, 2, expected));
        using var seekable = new MemoryStream(png, writable: false);
        seekable.Position = 0;

        Assert.True(OfficeRasterImageDecoder.TryDecode(seekable, out OfficeRasterImage? seekableImage));
        Assert.Equal(0L, seekable.Position);
        Assert.Equal(expected, seekableImage!.GetPixel(2, 1));

        using var forwardOnly = new ForwardOnlyReadStream(png);
        Assert.True(OfficeRasterImageDecoder.TryDecode(forwardOnly, out OfficeRasterImage? forwardImage));
        Assert.Equal(expected, forwardImage!.GetPixel(2, 1));
    }

    [Fact]
    public void RasterStreamLimitsAndCancellationFailBeforeDecodeWork() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(64, 64, OfficeColor.SteelBlue));
        var bounded = new OfficeRasterDecodeOptions { MaximumEncodedBytes = png.Length - 1 };
        using var oversized = new MemoryStream(png, writable: false);
        Assert.False(OfficeRasterImageDecoder.TryDecode(oversized, bounded, out _, out OfficeRasterDecodeInfo boundedInfo));
        Assert.False(boundedInfo.Succeeded);
        Assert.Equal(0L, oversized.Position);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var cancelled = new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token };
        using var cancelledStream = new ForwardOnlyReadStream(png);
        Assert.Throws<OperationCanceledException>(() =>
            OfficeRasterImageDecoder.TryDecode(cancelledStream, cancelled, out _, out _));
        Assert.Equal(0, cancelledStream.ReadCount);
    }

    [Fact]
    public void BoundedForwardOnlyReadsConsumeAtMostOneByteBeyondTheLimit() {
        using var stream = new ForwardOnlyReadStream(new byte[128]);

        Assert.False(OfficeBoundedStreamReader.TryRead(
            stream, maximumBytes: 7, CancellationToken.None, out _));

        Assert.Equal(8L, stream.Position);
    }

    [Fact]
    public void ManagedJpegBmpAndGifCodecsPropagateCancellation() {
        OfficeRasterImage source = new OfficeRasterImage(2, 2, OfficeColor.SteelBlue);
        byte[] jpeg = OfficeJpegCodec.Encode(source);
        byte[] bmp = new byte[58];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteInt32LittleEndian(bmp, 2, bmp.Length);
        WriteInt32LittleEndian(bmp, 10, 54);
        WriteInt32LittleEndian(bmp, 14, 40);
        WriteInt32LittleEndian(bmp, 18, 1);
        WriteInt32LittleEndian(bmp, 22, 1);
        WriteUInt16LittleEndian(bmp, 26, 1);
        WriteUInt16LittleEndian(bmp, 28, 24);
        byte[] gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeJpegCodec.TryDecode(jpeg, cancellation.Token, out _));
        Assert.Throws<OperationCanceledException>(() =>
            OfficeBmpReader.TryDecode(bmp, cancellation.Token, out _));
        Assert.Throws<OperationCanceledException>(() =>
            OfficeGifReader.TryDecodeFrame(gif, 0, cancellation.Token, out _, out _));
    }

    [Fact]
    public void ContainerInspectorReportsApngTimingAndSelectionSemantics() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(png);

        Assert.True(OfficeRasterContainerInspector.TryInspect(apng, out OfficeRasterContainerInfo? container));
        Assert.NotNull(container);
        Assert.Equal(2, container!.Count);
        Assert.True(container.IsAnimated);
        Assert.False(container.IsMultiPage);
        Assert.Equal(TimeSpan.FromMilliseconds(10), container.Frames[0].Duration);
        Assert.True(container.Frames[0].IsDefaultImage);
        Assert.False(container.Frames[1].IsDefaultImage);
    }

    [Fact]
    public void ContainerInspectorRejectsApngLoopCountsOutsideItsPublicContract() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(png);
        int animationControl = FindPngChunk(apng, "acTL");
        WriteBigEndianInt32(apng, animationControl + 12, int.MinValue);
        WritePngChunkCrc(apng, animationControl, 8);

        Assert.True(OfficePngReader.TryGetFrameCount(apng, out int frameCount));
        Assert.Equal(2, frameCount);
        Assert.False(OfficeRasterContainerInspector.TryInspect(apng, out _));
    }

    [Fact]
    public void ContainerInspectorRejectsInvalidApngFrameSequencing() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(png);
        int frameData = FindPngChunk(apng, "fdAT");
        WriteBigEndianInt32(apng, frameData + 8, 7);
        WritePngChunkCrc(apng, frameData, ReadBigEndianInt32(apng, frameData));

        Assert.True(OfficePngReader.TryGetFrameCount(apng, out int frameCount));
        Assert.Equal(2, frameCount);
        Assert.False(OfficeRasterContainerInspector.TryInspect(apng, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(
            apng,
            new OfficeRasterDecodeOptions {
                AnimationPolicy = OfficeRasterAnimationPolicy.UseSelectedFrame,
                FrameIndex = 1
            },
            out _,
            out _));
    }

    [Fact]
    public void ApngAcceptsFrameControlBeforeAnimationControl() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(png);
        int animationControl = FindPngChunk(apng, "acTL");
        int animationControlLength = ReadBigEndianInt32(apng, animationControl) + 12;
        int firstFrameControl = FindPngChunk(apng, "fcTL");
        int firstFrameControlLength = ReadBigEndianInt32(apng, firstFrameControl) + 12;
        Assert.Equal(animationControl + animationControlLength, firstFrameControl);

        byte[] reordered = (byte[])apng.Clone();
        Buffer.BlockCopy(apng, firstFrameControl, reordered, animationControl, firstFrameControlLength);
        Buffer.BlockCopy(apng, animationControl, reordered, animationControl + firstFrameControlLength,
            animationControlLength);

        Assert.True(OfficeImageReader.TryValidateContent(reordered, "reordered-animation-control.png", out _));
        Assert.True(OfficeRasterContainerInspector.TryInspect(reordered, out OfficeRasterContainerInfo? container));
        Assert.Equal(2, container!.Count);
        Assert.True(OfficeRasterImageDecoder.TryDecode(
            reordered,
            new OfficeRasterDecodeOptions { FrameIndex = 1 },
            out OfficeRasterImage? selected,
            out _));
        Assert.Equal(OfficeColor.Lime, selected!.GetPixel(0, 0));
    }

    [Fact]
    public void ApngTreatsFirstFramePreviousDisposalAsBackground() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(png);
        int firstFrameControl = FindPngChunk(apng, "fcTL");
        apng[firstFrameControl + 8 + 24] = 2;
        WritePngChunkCrc(apng, firstFrameControl, ReadBigEndianInt32(apng, firstFrameControl));

        Assert.True(OfficeImageReader.TryValidateContent(apng, "first-previous-disposal.png", out _));
        Assert.True(OfficeRasterContainerInspector.TryInspect(apng, out OfficeRasterContainerInfo? container));
        Assert.Equal(OfficeRasterFrameDisposal.Background, container!.Frames[0].Disposal);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void ApngValidationCopiesEachSecondaryFramePayloadAtMostOnce() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(png);
        int frameDataOffset = FindPngChunk(apng, "fdAT");
        int frameDataLength = ReadBigEndianInt32(apng, frameDataOffset);
        var largeFrameData = new byte[8 * 1024 * 1024 + 4];
        Buffer.BlockCopy(apng, frameDataOffset + 8, largeFrameData, 0, 4);
        byte[] replacement = CreatePngChunk("fdAT", largeFrameData);
        byte[] expanded = new byte[apng.Length - frameDataLength - 12 + replacement.Length];
        Buffer.BlockCopy(apng, 0, expanded, 0, frameDataOffset);
        Buffer.BlockCopy(replacement, 0, expanded, frameDataOffset, replacement.Length);
        Buffer.BlockCopy(apng, frameDataOffset + frameDataLength + 12, expanded,
            frameDataOffset + replacement.Length,
            apng.Length - frameDataOffset - frameDataLength - 12);

        long before = GC.GetAllocatedBytesForCurrentThread();
        Assert.False(OfficePngAnimationValidator.TryValidateAdditionalFrames(expanded));
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.True(allocated < 12L * 1024L * 1024L,
            $"APNG validation allocated {allocated:N0} bytes for an 8 MB secondary payload.");
    }
#endif

    [Fact]
    public void SelectedApngFrameDecodesAdam7Payload() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        png[28] = 1;
        WritePngChunkCrc(png, 8, 13);
        byte[] apng = CreateTwoFrameApng(png);
        var options = new OfficeRasterDecodeOptions {
            AnimationPolicy = OfficeRasterAnimationPolicy.UseSelectedFrame,
            FrameIndex = 1
        };

        Assert.True(OfficeRasterImageDecoder.TryDecode(
            apng, options, out OfficeRasterImage? image, out OfficeRasterDecodeInfo info));
        Assert.Equal(OfficeColor.Lime, image!.GetPixel(0, 0));
        Assert.True(info.IsAnimated);
        Assert.True(info.AnimationDiscarded);
    }

    [Fact]
    public void AnimatedWebpFramesDoNotClaimAStaticDefaultImage() {
        byte[] animated = Convert.FromBase64String(
            "UklGRoQAAABXRUJQVlA4WAoAAAACAAAAAQAAAQAAQU5JTQYAAAAAAAAAAABBTk1GKAAAAAAAAAAAAAEAAAEAAGQAAAJWUDhMDwAAAC8BQAAABxD9j/4HIqL/AQBBTk1GKAAAAAAAAAAAAAEAAAEAAGQAAABWUDhMDwAAAC8BQAAABxDR//4HIqL/AQA=");

        Assert.True(OfficeRasterContainerInspector.TryInspect(
            animated, out OfficeRasterContainerInfo? container));
        Assert.True(container!.IsAnimated);
        Assert.All(container.Frames, frame => Assert.False(frame.IsDefaultImage));
    }

    [Fact]
    public void ContainerInspectorPreservesSingleFrameApngAfterFallbackImageData() {
        byte[] fallback = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        byte[] animatedFrame = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateSingleFrameApngAfterFallback(fallback, animatedFrame);

        Assert.True(OfficeRasterContainerInspector.TryInspect(apng, out OfficeRasterContainerInfo? container));
        Assert.True(container!.IsAnimated);
        Assert.Single(container.Frames);
        Assert.False(container.Frames[0].IsDefaultImage);

        var options = new OfficeRasterDecodeOptions {
            AnimationPolicy = OfficeRasterAnimationPolicy.UseSelectedFrame
        };
        Assert.True(OfficeRasterImageDecoder.TryDecode(apng, options, out OfficeRasterImage? image, out OfficeRasterDecodeInfo info));
        Assert.Equal(OfficeColor.Lime, image!.GetPixel(0, 0));
        Assert.True(info.IsAnimated);
        Assert.True(info.AnimationDiscarded);
        Assert.False(info.FramesOrPagesDiscarded);
        Assert.NotNull(info.Diagnostic);

        var reject = new OfficeRasterDecodeOptions {
            AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated
        };
        Assert.False(OfficeRasterImageDecoder.TryDecode(apng, reject, out _, out _));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void SelectedApngDecodeDoesNotCopyUnselectedFramePayloads() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateTwoFrameApng(staticPng);
        int frameDataOffset = FindPngChunk(apng, "fdAT");
        int frameDataLength = ReadBigEndianInt32(apng, frameDataOffset);
        var largeFrameData = new byte[8 * 1024 * 1024 + 4];
        Buffer.BlockCopy(apng, frameDataOffset + 8, largeFrameData, 0, 4);
        byte[] replacement = CreatePngChunk("fdAT", largeFrameData);
        byte[] expanded = new byte[apng.Length - frameDataLength - 12 + replacement.Length];
        Buffer.BlockCopy(apng, 0, expanded, 0, frameDataOffset);
        Buffer.BlockCopy(replacement, 0, expanded, frameDataOffset, replacement.Length);
        Buffer.BlockCopy(apng, frameDataOffset + frameDataLength + 12, expanded,
            frameDataOffset + replacement.Length,
            apng.Length - frameDataOffset - frameDataLength - 12);

        long before = GC.GetAllocatedBytesForCurrentThread();
        Assert.True(OfficeRasterImageDecoder.TryDecode(
            expanded,
            new OfficeRasterDecodeOptions { FrameIndex = 0 },
            out OfficeRasterImage? selected,
            out _));
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Equal(OfficeColor.Lime, selected!.GetPixel(0, 0));
        Assert.True(allocated < 4L * 1024L * 1024L, $"Selected-frame decode allocated {allocated:N0} bytes.");
    }
#endif

    [Fact]
    public void ContentValidationUsesTheCanonicalEncodedPayloadLimit() {
        Assert.True(OfficeRasterGuards.IsEncodedPayloadWithinLimits(OfficeRasterGuards.MaximumEncodedBytes));
        Assert.False(OfficeRasterGuards.IsEncodedPayloadWithinLimits(OfficeRasterGuards.MaximumEncodedBytes + 1));
        Assert.False(OfficeRasterGuards.IsEncodedPayloadWithinLimits(0));
    }

    private static int FindPngChunk(byte[] bytes, string expectedType, int offset = 8) {
        while (offset + 12 <= bytes.Length) {
            int length = ReadBigEndianInt32(bytes, offset);
            string type = System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == expectedType) return offset;
            offset += 12 + length;
        }
        throw new InvalidDataException("PNG chunk was not found.");
    }

    private static OfficeImageExportResult CreateResult(string name, byte[] png) =>
        new(OfficeImageExportFormat.Png, 1, 1, png, name: name);

    private static byte[] CreateAlphaWebp(byte control) {
        byte[] lossy = Convert.FromBase64String(
            "UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoCAAIAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=");
        var bytes = new List<byte>(lossy.Length + 32) {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P'
        };
        bytes.AddRange(CreateWebpChunk("VP8X", new byte[] { 0x10, 0, 0, 0, 1, 0, 0, 1, 0, 0 }));
        bytes.AddRange(CreateWebpChunk("ALPH", new byte[] { control, 0xFF }));
        bytes.AddRange(lossy.Skip(12));
        byte[] result = bytes.ToArray();
        WriteInt32LittleEndian(result, 4, result.Length - 8);
        return result;
    }

    private static byte[] CreateWebpChunk(string type, byte[] data) {
        byte[] chunk = new byte[8 + data.Length + (data.Length & 1)];
        System.Text.Encoding.ASCII.GetBytes(type, 0, 4, chunk, 0);
        WriteInt32LittleEndian(chunk, 4, data.Length);
        Buffer.BlockCopy(data, 0, chunk, 8, data.Length);
        return chunk;
    }

    private static int FindWebpChunk(byte[] bytes, string expectedType) {
        int offset = 12;
        while (offset <= bytes.Length - 8) {
            string type = System.Text.Encoding.ASCII.GetString(bytes, offset, 4);
            int length = bytes[offset + 4] |
                         bytes[offset + 5] << 8 |
                         bytes[offset + 6] << 16 |
                         bytes[offset + 7] << 24;
            if (type == expectedType) return offset;
            offset += 8 + length + (length & 1);
        }
        throw new InvalidDataException("WebP chunk was not found.");
    }

    private static byte[] CreateTwoFrameApng(byte[] png) {
        int idatOffset = FindPngChunk(png, "IDAT");
        int idatLength = ReadBigEndianInt32(png, idatOffset);
        int idatEnd = idatOffset + 12 + idatLength;
        byte[] animationControl = new byte[8];
        WriteBigEndianInt32(animationControl, 0, 2);
        byte[] firstFrameControl = CreateFrameControl(sequence: 0);
        byte[] secondFrameControl = CreateFrameControl(sequence: 1);
        byte[] secondFrameData = new byte[idatLength + 4];
        WriteBigEndianInt32(secondFrameData, 0, 2);
        Buffer.BlockCopy(png, idatOffset + 8, secondFrameData, 4, idatLength);
        byte[] prefix = CreatePngChunk("acTL", animationControl)
            .Concat(CreatePngChunk("fcTL", firstFrameControl))
            .ToArray();
        byte[] suffix = CreatePngChunk("fcTL", secondFrameControl)
            .Concat(CreatePngChunk("fdAT", secondFrameData))
            .ToArray();
        byte[] result = new byte[png.Length + prefix.Length + suffix.Length];
        Buffer.BlockCopy(png, 0, result, 0, idatOffset);
        Buffer.BlockCopy(prefix, 0, result, idatOffset, prefix.Length);
        Buffer.BlockCopy(png, idatOffset, result, idatOffset + prefix.Length, idatEnd - idatOffset);
        Buffer.BlockCopy(suffix, 0, result, idatEnd + prefix.Length, suffix.Length);
        Buffer.BlockCopy(png, idatEnd, result, idatEnd + prefix.Length + suffix.Length, png.Length - idatEnd);
        return result;
    }

    private static byte[] CreateRepeatedFrameApng(byte[] png, int frameCount) {
        if (frameCount < 1) throw new ArgumentOutOfRangeException(nameof(frameCount));
        int ihdrOffset = FindPngChunk(png, "IHDR");
        int width = ReadBigEndianInt32(png, ihdrOffset + 8);
        int height = ReadBigEndianInt32(png, ihdrOffset + 12);
        int idatOffset = FindPngChunk(png, "IDAT");
        int idatLength = ReadBigEndianInt32(png, idatOffset);
        int idatEnd = idatOffset + 12 + idatLength;
        byte[] animationControl = new byte[8];
        WriteBigEndianInt32(animationControl, 0, frameCount);
        byte[] prefix = CreatePngChunk("acTL", animationControl)
            .Concat(CreatePngChunk("fcTL", CreateFrameControl(0, width, height)))
            .ToArray();
        var suffix = new List<byte>();
        for (int frame = 1; frame < frameCount; frame++) {
            suffix.AddRange(CreatePngChunk(
                "fcTL",
                CreateFrameControl(checked(frame * 2 - 1), width, height)));
            byte[] frameData = new byte[idatLength + 4];
            WriteBigEndianInt32(frameData, 0, checked(frame * 2));
            Buffer.BlockCopy(png, idatOffset + 8, frameData, 4, idatLength);
            suffix.AddRange(CreatePngChunk("fdAT", frameData));
        }

        byte[] result = new byte[png.Length + prefix.Length + suffix.Count];
        Buffer.BlockCopy(png, 0, result, 0, idatOffset);
        Buffer.BlockCopy(prefix, 0, result, idatOffset, prefix.Length);
        Buffer.BlockCopy(png, idatOffset, result, idatOffset + prefix.Length, idatEnd - idatOffset);
        suffix.CopyTo(result, idatEnd + prefix.Length);
        Buffer.BlockCopy(
            png,
            idatEnd,
            result,
            idatEnd + prefix.Length + suffix.Count,
            png.Length - idatEnd);
        return result;
    }

    private static byte[] CreateAdam7RgbaPng(OfficeRasterImage image) {
        int[] startX = { 0, 4, 0, 2, 0, 1, 0 };
        int[] startY = { 0, 0, 4, 0, 2, 0, 1 };
        int[] stepX = { 8, 8, 4, 4, 2, 2, 1 };
        int[] stepY = { 8, 8, 8, 4, 4, 2, 2 };
        var scanlines = new List<byte>();
        for (int pass = 0; pass < 7; pass++) {
            for (int y = startY[pass]; y < image.Height; y += stepY[pass]) {
                if (startX[pass] >= image.Width) continue;
                scanlines.Add(0);
                for (int x = startX[pass]; x < image.Width; x += stepX[pass]) {
                    OfficeColor color = image.GetPixel(x, y);
                    scanlines.Add(color.R);
                    scanlines.Add(color.G);
                    scanlines.Add(color.B);
                    scanlines.Add(color.A);
                }
            }
        }

        byte[] png = OfficePngWriter.Encode(image);
        png[28] = 1;
        WritePngChunkCrc(png, 8, 13);
        int idatOffset = FindPngChunk(png, "IDAT");
        int idatLength = ReadBigEndianInt32(png, idatOffset);
        int idatEnd = idatOffset + idatLength + 12;
        byte[] replacement = CreatePngChunk("IDAT", OfficeZlibCodec.Compress(scanlines.ToArray()));
        byte[] result = new byte[png.Length - idatLength - 12 + replacement.Length];
        Buffer.BlockCopy(png, 0, result, 0, idatOffset);
        Buffer.BlockCopy(replacement, 0, result, idatOffset, replacement.Length);
        Buffer.BlockCopy(png, idatEnd, result, idatOffset + replacement.Length, png.Length - idatEnd);
        return result;
    }

    private static byte[] CreateSingleFrameApngAfterFallback(byte[] fallback, byte[] animatedFrame) {
        int fallbackIdat = FindPngChunk(fallback, "IDAT");
        int fallbackIdatLength = ReadBigEndianInt32(fallback, fallbackIdat);
        int fallbackIdatEnd = fallbackIdat + 12 + fallbackIdatLength;
        int animatedIdat = FindPngChunk(animatedFrame, "IDAT");
        int animatedIdatLength = ReadBigEndianInt32(animatedFrame, animatedIdat);

        byte[] animationControl = new byte[8];
        WriteBigEndianInt32(animationControl, 0, 1);
        byte[] frameData = new byte[animatedIdatLength + 4];
        WriteBigEndianInt32(frameData, 0, 1);
        Buffer.BlockCopy(animatedFrame, animatedIdat + 8, frameData, 4, animatedIdatLength);

        byte[] prefix = CreatePngChunk("acTL", animationControl);
        byte[] suffix = CreatePngChunk("fcTL", CreateFrameControl(sequence: 0))
            .Concat(CreatePngChunk("fdAT", frameData))
            .ToArray();
        byte[] result = new byte[fallback.Length + prefix.Length + suffix.Length];
        Buffer.BlockCopy(fallback, 0, result, 0, fallbackIdat);
        Buffer.BlockCopy(prefix, 0, result, fallbackIdat, prefix.Length);
        Buffer.BlockCopy(fallback, fallbackIdat, result, fallbackIdat + prefix.Length,
            fallbackIdatEnd - fallbackIdat);
        Buffer.BlockCopy(suffix, 0, result, fallbackIdatEnd + prefix.Length, suffix.Length);
        Buffer.BlockCopy(fallback, fallbackIdatEnd, result,
            fallbackIdatEnd + prefix.Length + suffix.Length, fallback.Length - fallbackIdatEnd);
        return result;
    }

    private static byte[] CreateFrameControl(int sequence, int width = 1, int height = 1) {
        byte[] data = new byte[26];
        WriteBigEndianInt32(data, 0, sequence);
        WriteBigEndianInt32(data, 4, width);
        WriteBigEndianInt32(data, 8, height);
        data[21] = 1;
        return data;
    }

    private static void WritePngChunkCrc(byte[] bytes, int chunkOffset, int length) {
        uint crc = 0xFFFFFFFFU;
        for (int index = chunkOffset + 4; index < chunkOffset + 8 + length; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
            }
        }
        crc ^= 0xFFFFFFFFU;
        int offset = chunkOffset + 8 + length;
        bytes[offset] = (byte)(crc >> 24);
        bytes[offset + 1] = (byte)(crc >> 16);
        bytes[offset + 2] = (byte)(crc >> 8);
        bytes[offset + 3] = (byte)crc;
    }

    private static byte[] InsertPngChunkBefore(byte[] png, string beforeType, string newType, byte[] data) {
        int beforeOffset = FindPngChunk(png, beforeType);
        byte[] chunk = CreatePngChunk(newType, data);
        byte[] result = new byte[png.Length + chunk.Length];
        Buffer.BlockCopy(png, 0, result, 0, beforeOffset);
        Buffer.BlockCopy(chunk, 0, result, beforeOffset, chunk.Length);
        Buffer.BlockCopy(png, beforeOffset, result, beforeOffset + chunk.Length, png.Length - beforeOffset);
        return result;
    }

    private static byte[] SplitPngImageDataWithAncillaryChunk(byte[] png) {
        int idatOffset = FindPngChunk(png, "IDAT");
        int length = ReadBigEndianInt32(png, idatOffset);
        int firstLength = length / 2;
        int secondLength = length - firstLength;
        byte[] first = new byte[firstLength];
        byte[] second = new byte[secondLength];
        Buffer.BlockCopy(png, idatOffset + 8, first, 0, firstLength);
        Buffer.BlockCopy(png, idatOffset + 8 + firstLength, second, 0, secondLength);
        byte[] firstChunk = CreatePngChunk("IDAT", first);
        byte[] ancillary = CreatePngChunk("vpAg", Array.Empty<byte>());
        byte[] secondChunk = CreatePngChunk("IDAT", second);
        int originalChunkEnd = idatOffset + 12 + length;
        byte[] result = new byte[png.Length - (12 + length) + firstChunk.Length + ancillary.Length + secondChunk.Length];
        int offset = 0;
        Buffer.BlockCopy(png, 0, result, offset, idatOffset);
        offset += idatOffset;
        Buffer.BlockCopy(firstChunk, 0, result, offset, firstChunk.Length);
        offset += firstChunk.Length;
        Buffer.BlockCopy(ancillary, 0, result, offset, ancillary.Length);
        offset += ancillary.Length;
        Buffer.BlockCopy(secondChunk, 0, result, offset, secondChunk.Length);
        offset += secondChunk.Length;
        Buffer.BlockCopy(png, originalChunkEnd, result, offset, png.Length - originalChunkEnd);
        return result;
    }

    private static byte[] CreatePngChunk(string type, byte[] data) {
        byte[] chunk = new byte[12 + data.Length];
        WriteBigEndianInt32(chunk, 0, data.Length);
        System.Text.Encoding.ASCII.GetBytes(type, 0, 4, chunk, 4);
        Buffer.BlockCopy(data, 0, chunk, 8, data.Length);
        WritePngChunkCrc(chunk, 0, data.Length);
        return chunk;
    }

    private static byte[] InsertByteBeforePngChunkChecksum(byte[] png, string chunkType) {
        int chunkOffset = FindPngChunk(png, chunkType);
        int length = ReadBigEndianInt32(png, chunkOffset);
        int insertionOffset = chunkOffset + 8 + length - 4;
        byte[] result = new byte[png.Length + 1];
        Buffer.BlockCopy(png, 0, result, 0, insertionOffset);
        result[insertionOffset] = 0xA5;
        Buffer.BlockCopy(png, insertionOffset, result, insertionOffset + 1, png.Length - insertionOffset);
        WriteBigEndianInt32(result, chunkOffset, length + 1);
        WritePngChunkCrc(result, chunkOffset, length + 1);
        return result;
    }

    private static byte[] CreateIcon(params byte[][] payloads) {
        int directoryLength = 6 + payloads.Length * 16;
        byte[] icon = new byte[directoryLength + payloads.Sum(payload => payload.Length)];
        WriteUInt16LittleEndian(icon, 2, 1);
        WriteUInt16LittleEndian(icon, 4, checked((ushort)payloads.Length));
        int payloadOffset = directoryLength;
        for (int index = 0; index < payloads.Length; index++) {
            int entryOffset = 6 + index * 16;
            icon[entryOffset] = 1;
            icon[entryOffset + 1] = 1;
            WriteUInt16LittleEndian(icon, entryOffset + 4, 1);
            WriteUInt16LittleEndian(icon, entryOffset + 6, 32);
            WriteInt32LittleEndian(icon, entryOffset + 8, payloads[index].Length);
            WriteInt32LittleEndian(icon, entryOffset + 12, payloadOffset);
            Buffer.BlockCopy(payloads[index], 0, icon, payloadOffset, payloads[index].Length);
            payloadOffset += payloads[index].Length;
        }
        return icon;
    }

    private static byte[] CreateIconWithSharedPayload(byte[] payload, ushort count) {
        int directoryLength = 6 + count * 16;
        byte[] icon = new byte[directoryLength + payload.Length];
        WriteUInt16LittleEndian(icon, 2, 1);
        WriteUInt16LittleEndian(icon, 4, count);
        for (int index = 0; index < count; index++) {
            int entryOffset = 6 + index * 16;
            icon[entryOffset] = 1;
            icon[entryOffset + 1] = 1;
            WriteUInt16LittleEndian(icon, entryOffset + 4, 1);
            WriteUInt16LittleEndian(icon, entryOffset + 6, 32);
            WriteInt32LittleEndian(icon, entryOffset + 8, payload.Length);
            WriteInt32LittleEndian(icon, entryOffset + 12, directoryLength);
        }
        Buffer.BlockCopy(payload, 0, icon, directoryLength, payload.Length);
        return icon;
    }

    private static byte[] CreateOnePixelIconDib() {
        var dib = new byte[48];
        WriteInt32LittleEndian(dib, 0, 40);
        WriteInt32LittleEndian(dib, 4, 1);
        WriteInt32LittleEndian(dib, 8, 2);
        WriteUInt16LittleEndian(dib, 12, 1);
        WriteUInt16LittleEndian(dib, 14, 32);
        WriteInt32LittleEndian(dib, 20, 4);
        dib[40] = 0xFF;
        dib[43] = 0xFF;
        return dib;
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private sealed class DeclaredLengthStream : Stream {
        private long _position;

        internal DeclaredLengthStream(long length) => Length = length;

        internal int ReadCount { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length { get; }
        public override long Position { get => _position; set => _position = value; }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { ReadCount++; return 0; }
        public override long Seek(long offset, SeekOrigin origin) => _position = offset;
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class ForwardOnlyReadStream : Stream {
        private readonly byte[] _bytes;
        private int _offset;

        internal ForwardOnlyReadStream(byte[] bytes) => _bytes = bytes;

        internal int ReadCount { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => _offset; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) {
            ReadCount++;
            int take = Math.Min(count, _bytes.Length - _offset);
            if (take <= 0) return 0;
            Buffer.BlockCopy(_bytes, _offset, buffer, offset, take);
            _offset += take;
            return take;
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
