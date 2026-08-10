using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void PngContainerRequiresOneWellFormedIccProfileBeforePaletteAndImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] compressedProfile = {
            0x78, 0x01,
            0x01, 0x03, 0x00, 0xFC, 0xFF, (byte)'i', (byte)'c', (byte)'c',
            0x02, 0x67, 0x01, 0x30
        };
        byte[] profile = new byte[] { (byte)'P', (byte)'r', (byte)'o', (byte)'f', (byte)'i', (byte)'l', (byte)'e', 0, 0 }
            .Concat(compressedProfile)
            .ToArray();
        byte[] withProfile = InsertPngChunkBefore(png, "IDAT", "iCCP", profile);
        byte[] duplicate = InsertPngChunkBefore(withProfile, "IDAT", "iCCP", profile);
        byte[] misplaced = InsertPngChunkBefore(png, "IEND", "iCCP", profile);
        byte[] withStandardRgb = InsertPngChunkBefore(png, "IDAT", "sRGB", new byte[] { 0 });
        byte[] conflictingColorProfiles = InsertPngChunkBefore(withStandardRgb, "IDAT", "iCCP", profile);
        byte[] invalidName = InsertPngChunkBefore(
            png,
            "IDAT",
            "iCCP",
            new byte[] { (byte)' ', 0, 0 }.Concat(compressedProfile).ToArray());
        byte[] invalidMethod = (byte[])profile.Clone();
        invalidMethod[8] = 1;
        invalidMethod = InsertPngChunkBefore(png, "IDAT", "iCCP", invalidMethod);
        byte[] invalidStream = (byte[])profile.Clone();
        invalidStream[invalidStream.Length - 1] ^= 0x01;
        invalidStream = InsertPngChunkBefore(png, "IDAT", "iCCP", invalidStream);

        Assert.True(OfficeImageReader.TryValidateContent(withProfile, "profile.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-profile.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(misplaced, "misplaced-profile.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(conflictingColorProfiles, "conflicting-profiles.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidName, "invalid-name.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidMethod, "invalid-method.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidStream, "invalid-stream.png", out _));
    }

    [Fact]
    public void CompleteContentValidationRejectsOutOfRangeOptionalTiffIfdValues() {
        byte[] tiff = OfficeTiffCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int resolutionEntry = FindClassicTiffEntry(tiff, 282);
        WriteInt32LittleEndian(tiff, resolutionEntry + 8, tiff.Length + 8);

        Assert.True(OfficeImageReader.TryIdentifyByContent(tiff, "truncated-resolution.tiff", out _));
        Assert.False(OfficeImageReader.TryValidateContent(tiff, "truncated-resolution.tiff", out _));
        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(OfficeImageExportFormat.Tiff, 1, 1, tiff));
    }

    private static int FindClassicTiffEntry(byte[] bytes, int expectedTag) {
        int ifdOffset = bytes[4] | bytes[5] << 8 | bytes[6] << 16 | bytes[7] << 24;
        int entryCount = bytes[ifdOffset] | bytes[ifdOffset + 1] << 8;
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + index * 12;
            int tag = bytes[entryOffset] | bytes[entryOffset + 1] << 8;
            if (tag == expectedTag) return entryOffset;
        }
        throw new InvalidOperationException("The expected TIFF entry was not found.");
    }
}
