using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPremiumTiffDecodingTests {
    [Fact]
    public void DeflateValidationStopsWhenExpansionCrossesTheDeclaredLimit() {
        byte[] rawDeflate = CompressRawDeflate(new byte[1024 * 1024]);

        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(
            rawDeflate,
            0,
            rawDeflate.Length,
            maximumOutputBytes: 32));
    }

    [Fact]
    public void ZlibRejectsExpectedOutputBeyondConfiguredLimitBeforePayloadValidation() {
        byte[] zlib = OfficeZlibCodec.Compress(new byte[] { 1 });
        zlib[2] ^= 0xFF;

        Assert.Throws<OfficeDecompressionSizeLimitException>(() =>
            OfficeZlibCodec.Decompress(zlib, maximumOutputBytes: 32, expectedOutputBytes: 64));
    }

    [Fact]
    public void ZlibTreatsExpansionBeyondSmallerExpectedSizeAsMalformedData() {
        byte[] zlib = OfficeZlibCodec.Compress(new byte[64]);

        Assert.Throws<InvalidDataException>(() =>
            OfficeZlibCodec.Decompress(zlib, maximumOutputBytes: 128, expectedOutputBytes: 32));
    }

    [Fact]
    public void TiffDecoder_DecodesLibTiffDeflatePredictorFixture() {
        byte[] tiff = Convert.FromBase64String(
            "SUkqABYAAAB4nPvPwMD4nwEACAECAA8AAAEDAAEAAAACAAAAAQEDAAEAAAABAAAAAgEDAAMAAADQAAAAAwEDAAEAAAAIAAAABgEDAAEAAAACAAAACgEDAAEAAAABAAAAEQEEAAEAAAAIAAAAEgEDAAEAAAABAAAAFQEDAAEAAAADAAAAFgEDAAEAAAABAAAAFwEEAAEAAAAOAAAAHAEDAAEAAAABAAAAKAEDAAEAAAACAAAAPQEDAAEAAAACAAAAUwEDAAMAAADWAAAAAAAAAAgACAAIAAEAAQABAA==");

        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(
            new byte[] {
                255, 0, 0, 255,
                0, 255, 0, 255
            },
            image!.GetPixels());
    }

    [Fact]
    public void TiffDecoder_DecodesDeflateGrayscaleWithHorizontalPrediction() {
        byte[] tiff = CreateTiff(
            width: 3,
            height: 1,
            photometric: 1,
            samples: 1,
            pixels: new byte[] { 0, 127, 255 },
            predictor: 2,
            includeSamplesPerPixel: false);

        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(
            new byte[] {
                0, 0, 0, 255,
                127, 127, 127, 255,
                255, 255, 255, 255
            },
            image!.GetPixels());
    }

    [Fact]
    public void TiffDecoder_DecodesDeflatePalettePixels() {
        var colorMap = new int[768];
        colorMap[0] = 65535;
        colorMap[256 + 1] = 65535;
        byte[] tiff = CreateTiff(
            width: 2,
            height: 1,
            photometric: 3,
            samples: 1,
            pixels: new byte[] { 0, 1 },
            colorMap: colorMap);

        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(
            new byte[] {
                255, 0, 0, 255,
                0, 255, 0, 255
            },
            image!.GetPixels());
    }

    [Theory]
    [InlineData(5, true)]
    [InlineData(6, true)]
    [InlineData(7, false)]
    [InlineData(8, false)]
    public void TiffDecoder_AppliesAxisSwappingOrientations(int orientation, bool redFirst) {
        byte[] tiff = CreateTiff(
            width: 2,
            height: 1,
            photometric: 2,
            samples: 3,
            pixels: new byte[] {
                255, 0, 0,
                0, 255, 0
            },
            orientation: orientation);

        Assert.True(OfficeImageOrientationNormalizer.TryRead(tiff, out OfficeImageOrientation identified));
        Assert.Equal((OfficeImageOrientation)orientation, identified);
        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal((1, 2), (image!.Width, image.Height));
        Assert.Equal(redFirst ? OfficeColor.Red : OfficeColor.Lime, image.GetPixel(0, 0));
        Assert.Equal(redFirst ? OfficeColor.Lime : OfficeColor.Red, image.GetPixel(0, 1));
    }

    [Fact]
    public void TiffDecoder_DecodesDeflateDeviceCmykPixels() {
        byte[] tiff = CreateTiff(
            width: 2,
            height: 1,
            photometric: 5,
            samples: 4,
            pixels: new byte[] {
                255, 0, 0, 0,
                0, 255, 0, 0
            });

        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(
            new byte[] {
                0, 255, 255, 255,
                255, 0, 255, 255
            },
            image!.GetPixels());
    }

    [Fact]
    public void TiffDecoder_RejectsSeparatedDataThatIsNotDeviceCmyk() {
        byte[] tiff = CreateTiff(
            width: 1,
            height: 1,
            photometric: 5,
            samples: 4,
            pixels: new byte[] { 0, 0, 0, 0 },
            inkSet: 2);

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.Null(image);
    }

    [Fact]
    public void TiffDecoder_RequiresExactRawDeflateStripConsumption() {
        byte[] tiff = CreateTiff(
            width: 2,
            height: 1,
            photometric: 1,
            samples: 1,
            pixels: new byte[] { 0, 255 },
            rawDeflate: true);
        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);

        int stripByteCountOffset = FindEntryValueOffset(tiff, 279);
        int stripByteCount = ReadUInt32(tiff, stripByteCountOffset);
        Array.Resize(ref tiff, tiff.Length + 1);
        tiff[tiff.Length - 1] = 0x5A;
        WriteUInt32(tiff, stripByteCountOffset, stripByteCount + 1);

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
        Assert.False(OfficeImageReader.TryValidateContent(tiff, "trailing-raw-deflate.tiff", out _));
    }

    [Fact]
    public void TiffDecoder_RejectsRawDeflateStoredUnderAdobeDeflateTag() {
        byte[] tiff = CreateTiff(
            width: 2,
            height: 1,
            photometric: 1,
            samples: 1,
            pixels: new byte[] { 0, 255 },
            rawDeflate: true);
        WriteUInt32(tiff, FindEntryValueOffset(tiff, 259), (int)OfficeTiffCompression.Deflate);

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
        Assert.False(OfficeImageReader.TryValidateContent(tiff, "raw-adobe-deflate.tiff", out _));
    }

    [Fact]
    public void TiffDecoderRejectsDeflateOutputBeyondTheDeclaredStripWithoutThrowing() {
        byte[] tiff = CreateTiff(
            width: 1,
            height: 1,
            photometric: 1,
            samples: 1,
            pixels: new byte[] { 0 });
        byte[] oversizedStrip = OfficeZlibCodec.Compress(new byte[4096]);
        int stripOffset = ReadUInt32(tiff, FindEntryValueOffset(tiff, 273));
        Array.Resize(ref tiff, stripOffset + oversizedStrip.Length);
        Buffer.BlockCopy(oversizedStrip, 0, tiff, stripOffset, oversizedStrip.Length);
        WriteUInt32(tiff, FindEntryValueOffset(tiff, 279), oversizedStrip.Length);

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
        Assert.False(OfficeImageReader.TryValidateContent(tiff, "oversized-deflate-strip.tiff", out _));
    }

    private static byte[] CreateTiff(
        int width,
        int height,
        int photometric,
        int samples,
        byte[] pixels,
        int predictor = 1,
        int[]? colorMap = null,
        bool includeSamplesPerPixel = true,
        int? inkSet = null,
        bool rawDeflate = false,
        int? orientation = null) {
        byte[] predicted = pixels.ToArray();
        if (predictor == 2) {
            int rowBytes = width * samples;
            for (int row = 0; row < height; row++) {
                int rowOffset = row * rowBytes;
                for (int index = rowBytes - 1; index >= samples; index--) {
                    predicted[rowOffset + index] = unchecked(
                        (byte)(predicted[rowOffset + index] - predicted[rowOffset + index - samples]));
                }
            }
        }
        byte[] strip = rawDeflate ? CompressRawDeflate(predicted) : OfficeZlibCodec.Compress(predicted);

        int entryCount =
            (colorMap == null ? 10 : 11) -
            (includeSamplesPerPixel ? 0 : 1) +
            (inkSet.HasValue ? 1 : 0) +
            (orientation.HasValue ? 1 : 0);
        const int ifdOffset = 8;
        int dataOffset = ifdOffset + 2 + (entryCount * 12) + 4;
        int bitsOffset = samples > 2 ? dataOffset : 0;
        if (samples > 2) dataOffset += samples * 2;
        int colorMapOffset = colorMap == null ? 0 : dataOffset;
        if (colorMap != null) dataOffset += colorMap.Length * 2;
        int stripOffset = dataOffset;
        var output = new byte[stripOffset + strip.Length];

        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, ifdOffset);
        WriteUInt16(output, ifdOffset, entryCount);

        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, width);
        WriteEntry(output, ref entry, 257, 4, 1, height);
        WriteEntry(
            output,
            ref entry,
            258,
            3,
            samples,
            samples > 2 ? bitsOffset : samples == 2 ? 8 | 8 << 16 : 8);
        WriteEntry(output, ref entry, 259, 3, 1, rawDeflate ? 32946 : (int)OfficeTiffCompression.Deflate);
        WriteEntry(output, ref entry, 262, 3, 1, photometric);
        WriteEntry(output, ref entry, 273, 4, 1, stripOffset);
        if (orientation.HasValue) {
            WriteEntry(output, ref entry, 274, 3, 1, orientation.Value);
        }
        if (includeSamplesPerPixel) {
            WriteEntry(output, ref entry, 277, 3, 1, samples);
        }
        WriteEntry(output, ref entry, 278, 4, 1, height);
        WriteEntry(output, ref entry, 279, 4, 1, strip.Length);
        WriteEntry(output, ref entry, 317, 3, 1, predictor);
        if (colorMap != null) {
            WriteEntry(output, ref entry, 320, 3, colorMap.Length, colorMapOffset);
        }
        if (inkSet.HasValue) {
            WriteEntry(output, ref entry, 332, 3, 1, inkSet.Value);
        }
        WriteUInt32(output, entry, 0);

        if (samples > 2) {
            for (int index = 0; index < samples; index++) {
                WriteUInt16(output, bitsOffset + index * 2, 8);
            }
        }
        if (colorMap != null) {
            for (int index = 0; index < colorMap.Length; index++) {
                WriteUInt16(output, colorMapOffset + index * 2, colorMap[index]);
            }
        }
        Buffer.BlockCopy(strip, 0, output, stripOffset, strip.Length);
        return output;
    }

    private static byte[] CompressRawDeflate(byte[] input) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(input, 0, input.Length);
        }
        return output.ToArray();
    }

    private static int FindEntryValueOffset(byte[] tiff, int expectedTag) {
        int ifdOffset = ReadUInt32(tiff, 4);
        int entryCount = tiff[ifdOffset] | tiff[ifdOffset + 1] << 8;
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            int tag = tiff[entryOffset] | tiff[entryOffset + 1] << 8;
            if (tag == expectedTag) return entryOffset + 8;
        }
        throw new InvalidOperationException("TIFF entry was not found.");
    }

    private static int ReadUInt32(byte[] input, int offset) =>
        input[offset] |
        input[offset + 1] << 8 |
        input[offset + 2] << 16 |
        input[offset + 3] << 24;

    private static void WriteEntry(
        byte[] output,
        ref int offset,
        int tag,
        int type,
        int count,
        int value) {
        WriteUInt16(output, offset, tag);
        WriteUInt16(output, offset + 2, type);
        WriteUInt32(output, offset + 4, count);
        WriteUInt32(output, offset + 8, value);
        offset += 12;
    }

    private static void WriteUInt16(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
        output[offset + 2] = (byte)(value >> 16);
        output[offset + 3] = (byte)(value >> 24);
    }
}
