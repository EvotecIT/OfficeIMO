using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
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
