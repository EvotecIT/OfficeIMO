using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void CompleteContentValidationAcceptsOnlySupportedIconDibHeaders() {
        foreach (int headerSize in new[] { 40, 52, 56, 108, 124 }) {
            Assert.True(OfficeImageReader.TryValidateContent(
                CreateIcon(CreateIconDibWithHeaderSize(headerSize)),
                "supported-header.ico",
                out _));
        }
        Assert.False(OfficeImageReader.TryValidateContent(
            CreateIcon(CreateIconDibWithHeaderSize(41)),
            "invented-header.ico",
            out _));
    }

    [Fact]
    public void CompleteContentValidationChecksIconBitfieldMasks() {
        byte[] valid = CreateBitfieldIconDib(3, 0xF800, 0x07E0, 0x001F, 0);
        byte[] zeroMask = CreateBitfieldIconDib(3, 0, 0x07E0, 0x001F, 0);
        byte[] overlapping = CreateBitfieldIconDib(3, 0xF800, 0xF800, 0x001F, 0);
        byte[] nonContiguous = CreateBitfieldIconDib(3, 0xA000, 0x07E0, 0x001F, 0);
        byte[] outsideBitDepth = CreateBitfieldIconDib(3, 0xF8000000, 0x07E0, 0x001F, 0);
        byte[] missingAlpha = CreateBitfieldIconDib(6, 0x7C00, 0x03E0, 0x001F, 0);

        Assert.True(OfficeImageReader.TryValidateContent(CreateIcon(valid), "bitfields.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(CreateIcon(zeroMask), "zero-mask.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(CreateIcon(overlapping), "overlap.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(CreateIcon(nonContiguous), "split-mask.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(CreateIcon(outsideBitDepth), "wide-mask.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(CreateIcon(missingAlpha), "missing-alpha.ico", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksRecognizableJfifHeaders() {
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int jfifOffset = FindAsciiSequence(jpeg, "JFIF\0");
        Assert.True(jfifOffset >= 2);

        byte[] invalidUnits = (byte[])jpeg.Clone();
        invalidUnits[jfifOffset + 7] = 3;
        byte[] missingThumbnail = (byte[])jpeg.Clone();
        missingThumbnail[jfifOffset + 12] = 1;
        missingThumbnail[jfifOffset + 13] = 1;
        byte[] truncatedHeader = (byte[])jpeg.Clone();
        truncatedHeader[jfifOffset - 2] = 0;
        truncatedHeader[jfifOffset - 1] = 15;

        Assert.True(OfficeImageReader.TryValidateContent(jpeg, "valid-jfif.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidUnits, "invalid-units.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(missingThumbnail, "missing-thumbnail.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(truncatedHeader, "truncated-jfif.jpg", out _));
    }

    [Fact]
    public void CompleteContentValidationBoundsAggregateExifIfdScheduling() {
        const int pointerCount = 1025;
        int tableLength = 2 + pointerCount * 12 + 4;
        var exif = new byte[8 + tableLength + pointerCount * 6];
        exif[0] = (byte)'I'; exif[1] = (byte)'I'; exif[2] = 42; exif[4] = 8;
        WriteUInt16LittleEndian(exif, 8, (ushort)pointerCount);
        for (int index = 0; index < pointerCount; index++) {
            int entry = 10 + index * 12;
            WriteUInt16LittleEndian(exif, entry, 1);
            WriteUInt16LittleEndian(exif, entry + 2, 13);
            WriteInt32LittleEndian(exif, entry + 4, 1);
            WriteInt32LittleEndian(exif, entry + 8, 8 + tableLength + index * 6);
        }
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "eXIf", exif), "ifd-budget.png", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksSuggestedPngPalettes() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] validPalette = { (byte)'p', 0, 8, 1, 2, 3, 4, 5, 0 };
        byte[] valid = InsertPngChunkBefore(png, "IDAT", "sPLT", validPalette);
        byte[] duplicate = InsertPngChunkBefore(valid, "IDAT", "sPLT", validPalette);
        Assert.True(OfficeImageReader.TryValidateContent(valid, "palette.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-palette.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "sPLT", Array.Empty<byte>()), "empty-palette.png", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksWebpXmpPackets() {
        byte[] simple = OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] valid = CreateExtendedWebpWithXmp(simple, Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>"));
        byte[] malformed = CreateExtendedWebpWithXmp(simple, Encoding.UTF8.GetBytes("<x:xmpmeta>"));
        Assert.True(OfficeImageReader.TryValidateContent(valid, "metadata.webp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "malformed-metadata.webp", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksRecognizableJpegXmpPackets() {
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        byte[] valid = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(
                xmp: Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>"))
        });
        byte[] malformed = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(xmp: Encoding.UTF8.GetBytes("<x:xmpmeta>"))
        });
        byte[] invalidUtf8 = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(xmp: new byte[] { 0xC3, 0x28 })
        });

        Assert.True(OfficeImageReader.TryValidateContent(valid, "metadata.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "malformed-metadata.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidUtf8, "invalid-utf8.jpg", out _));

        int xmpOffset = FindAsciiSequence(valid, "http://ns.adobe.com/xap/1.0/\0");
        int segmentOffset = xmpOffset - 4;
        int segmentLength = valid[xmpOffset - 2] << 8 | valid[xmpOffset - 1];
        byte[] duplicate = valid.Take(segmentOffset)
            .Concat(valid.Skip(segmentOffset).Take(segmentLength + 2))
            .Concat(valid.Skip(segmentOffset))
            .ToArray();
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-xmp.jpg", out _));
    }

    [Fact]
    public void CompleteContentValidationReconcilesOmittedIconMasksWithImageSize() {
        byte[] omittedMask = CreateOnePixelIconDib().Take(44).ToArray();
        byte[] valid = CreateIcon(omittedMask);
        byte[] mismatched = (byte[])omittedMask.Clone();
        WriteInt32LittleEndian(mismatched, 20, 8);
        Assert.True(OfficeImageReader.TryValidateContent(valid, "omitted-mask.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(CreateIcon(mismatched), "truncated-mask.ico", out _));
    }

    [Theory]
    [InlineData(1, 0x80)]
    [InlineData(4, 0xF0)]
    [InlineData(8, 0x01)]
    public void CompleteContentValidationRejectsIndexedIconPixelsOutsideTheDeclaredPalette(
        int bitsPerPixel,
        int invalidPackedPixel) {
        byte[] valid = CreateIcon(CreateIndexedIconDib(bitsPerPixel, packedPixel: 0));
        byte[] invalid = CreateIcon(CreateIndexedIconDib(bitsPerPixel, (byte)invalidPackedPixel));

        Assert.True(OfficeImageReader.TryValidateContent(valid, "indexed.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalid, "invalid-indexed.ico", out _));
    }

    [Fact]
    public void ExtendedVp8lUsesDecodedAlphaInsteadOfTheCodecHint() {
        byte[] simple = OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var bytes = new List<byte> {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P'
        };
        bytes.AddRange(CreateWebpChunk("VP8X", new byte[10]));
        bytes.AddRange(simple.Skip(12));
        byte[] conservativeHint = bytes.ToArray();
        WriteInt32LittleEndian(conservativeHint, 4, conservativeHint.Length - 8);
        int vp8lOffset = FindWebpChunk(conservativeHint, "VP8L");
        conservativeHint[vp8lOffset + 12] |= 0x10;

        Assert.True(OfficeImageReader.TryIdentifyByContent(conservativeHint, "hint.webp", out _));
        Assert.True(OfficeImageReader.TryValidateContent(conservativeHint, "hint.webp", out _));

        byte[] incorrectContainerFlag = (byte[])conservativeHint.Clone();
        int vp8xOffset = FindWebpChunk(incorrectContainerFlag, "VP8X");
        incorrectContainerFlag[vp8xOffset + 8] |= 0x10;
        Assert.False(OfficeImageReader.TryIdentifyByContent(incorrectContainerFlag, "wrong-alpha-flag.webp", out _));
    }

    [Fact]
    public void CompleteContentValidationRequiresStructurallyValidWebpIccProfiles() {
        byte[] simple = OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] profile = CreateMinimalIccProfile();
        byte[] valid = CreateExtendedWebpWithIcc(simple, profile);
        byte[] malformedProfile = (byte[])profile.Clone();
        WriteBigEndianInt32(malformedProfile, 0, malformedProfile.Length + 4);
        byte[] malformed = CreateExtendedWebpWithIcc(simple, malformedProfile);

        Assert.True(OfficeImageReader.TryValidateContent(valid, "profile.webp", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "malformed-profile.webp", out _));
    }

    [Fact]
    public void CompleteContentValidationRejectsMalformedJpegExifAndIccProfiles() {
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        byte[] exif = {
            (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
            0, 0,
            0, 0, 0, 0
        };
        byte[] validExif = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(exif: exif)
        });
        byte[] malformedExif = (byte[])exif.Clone();
        malformedExif[4] = 0x40;
        byte[] invalidExif = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(exif: malformedExif)
        });
        byte[] validIcc = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(icc: CreateMinimalIccProfile())
        });
        byte[] malformedIcc = CreateMinimalIccProfile();
        malformedIcc[36] = (byte)'X';
        byte[] invalidIcc = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Metadata = new OfficeJpegMetadata(icc: malformedIcc)
        });

        Assert.True(OfficeImageReader.TryValidateContent(validExif, "exif.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidExif, "malformed-exif.jpg", out _));
        Assert.True(OfficeImageReader.TryValidateContent(validIcc, "profile.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidIcc, "malformed-profile.jpg", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksGifPlainTextRenderingHeaders() {
        var strictSource = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==").ToList();
        strictSource[30] = 2;
        strictSource.Insert(32, 0x01);
        byte[] source = strictSource.ToArray();
        byte[] valid = InsertGifPlainTextExtension(source, hasGlobalColorTable: true);
        byte[] missingGlobalPalette = InsertGifPlainTextExtension(
            RemoveGifGlobalColorTable(source),
            hasGlobalColorTable: false);
        byte[] invalidGrid = InsertGifPlainTextExtension(source, hasGlobalColorTable: true, gridWidth: 2);
        byte[] invalidCell = InsertGifPlainTextExtension(source, hasGlobalColorTable: true, cellWidth: 0);
        byte[] invalidColor = InsertGifPlainTextExtension(source, hasGlobalColorTable: true, foregroundIndex: 2);

        Assert.True(OfficeImageReader.TryValidateContent(valid, "plain-text.gif", out _));
        Assert.False(OfficeImageReader.TryValidateContent(missingGlobalPalette, "no-global-palette.gif", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidGrid, "invalid-grid.gif", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidCell, "invalid-cell.gif", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidColor, "invalid-color.gif", out _));
    }

    private static byte[] CreateIndexedIconDib(int bitsPerPixel, byte packedPixel) {
        const int headerSize = 40;
        const int paletteBytes = 4;
        const int xorBytes = 4;
        const int maskBytes = 4;
        var dib = new byte[headerSize + paletteBytes + xorBytes + maskBytes];
        WriteInt32LittleEndian(dib, 0, headerSize);
        WriteInt32LittleEndian(dib, 4, 1);
        WriteInt32LittleEndian(dib, 8, 2);
        WriteUInt16LittleEndian(dib, 12, 1);
        WriteUInt16LittleEndian(dib, 14, (ushort)bitsPerPixel);
        WriteInt32LittleEndian(dib, 20, xorBytes);
        WriteInt32LittleEndian(dib, 32, 1);
        dib[headerSize + paletteBytes] = packedPixel;
        return dib;
    }

    private static byte[] CreateBitfieldIconDib(
        int compression,
        uint red,
        uint green,
        uint blue,
        uint alpha) {
        const int headerSize = 40;
        int maskCount = compression == 6 ? 4 : 3;
        const int xorBytes = 4;
        const int andMaskBytes = 4;
        var dib = new byte[headerSize + maskCount * 4 + xorBytes + andMaskBytes];
        WriteInt32LittleEndian(dib, 0, headerSize);
        WriteInt32LittleEndian(dib, 4, 1);
        WriteInt32LittleEndian(dib, 8, 2);
        WriteUInt16LittleEndian(dib, 12, 1);
        WriteUInt16LittleEndian(dib, 14, 16);
        WriteInt32LittleEndian(dib, 16, compression);
        WriteInt32LittleEndian(dib, 20, xorBytes);
        WriteInt32LittleEndian(dib, 40, unchecked((int)red));
        WriteInt32LittleEndian(dib, 44, unchecked((int)green));
        WriteInt32LittleEndian(dib, 48, unchecked((int)blue));
        if (maskCount == 4) WriteInt32LittleEndian(dib, 52, unchecked((int)alpha));
        return dib;
    }

    private static byte[] CreateIconDibWithHeaderSize(int headerSize) {
        var dib = new byte[headerSize + 4];
        WriteInt32LittleEndian(dib, 0, headerSize);
        WriteInt32LittleEndian(dib, 4, 1);
        WriteInt32LittleEndian(dib, 8, 2);
        WriteUInt16LittleEndian(dib, 12, 1);
        WriteUInt16LittleEndian(dib, 14, 32);
        WriteInt32LittleEndian(dib, 20, 4);
        dib[headerSize] = 0xFF;
        dib[headerSize + 3] = 0xFF;
        return dib;
    }

    private static int FindAsciiSequence(byte[] bytes, string value) {
        byte[] expected = Encoding.ASCII.GetBytes(value);
        for (int offset = 0; offset <= bytes.Length - expected.Length; offset++) {
            bool match = true;
            for (int index = 0; index < expected.Length; index++) {
                if (bytes[offset + index] == expected[index]) continue;
                match = false;
                break;
            }
            if (match) return offset;
        }
        return -1;
    }

    private static byte[] CreateExtendedWebpWithIcc(byte[] simple, byte[] profile) {
        var bytes = new List<byte> {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P'
        };
        bytes.AddRange(CreateWebpChunk("VP8X", new byte[] { 0x20, 0, 0, 0, 0, 0, 0, 0, 0, 0 }));
        bytes.AddRange(CreateWebpChunk("ICCP", profile));
        bytes.AddRange(simple.Skip(12));
        byte[] result = bytes.ToArray();
        WriteInt32LittleEndian(result, 4, result.Length - 8);
        return result;
    }

    private static byte[] CreateExtendedWebpWithXmp(byte[] simple, byte[] xmp) {
        var bytes = new List<byte> {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P'
        };
        bytes.AddRange(CreateWebpChunk("VP8X", new byte[] { 0x04, 0, 0, 0, 0, 0, 0, 0, 0, 0 }));
        bytes.AddRange(simple.Skip(12));
        bytes.AddRange(CreateWebpChunk("XMP ", xmp));
        byte[] result = bytes.ToArray();
        WriteInt32LittleEndian(result, 4, result.Length - 8);
        return result;
    }

    private static byte[] InsertGifPlainTextExtension(
        byte[] gif,
        bool hasGlobalColorTable,
        int gridWidth = 1,
        int cellWidth = 1,
        int foregroundIndex = 0) {
        int descriptorOffset = hasGlobalColorTable ? 19 : 13;
        byte[] extension = {
            0x21, 0x01, 0x0C,
            0, 0, 0, 0,
            (byte)gridWidth, 0, 1, 0,
            (byte)cellWidth, 1,
            (byte)foregroundIndex, 1,
            1, (byte)'A', 0
        };
        return gif.Take(descriptorOffset).Concat(extension).Concat(gif.Skip(descriptorOffset)).ToArray();
    }

    private static byte[] RemoveGifGlobalColorTable(byte[] gif) {
        var result = gif.Take(13).Concat(gif.Skip(19)).ToArray();
        result[10] &= 0x7F;
        result[11] = 0;
        result[22] |= 0x80;
        return result.Take(23)
            .Concat(gif.Skip(13).Take(6))
            .Concat(result.Skip(23))
            .ToArray();
    }
}
