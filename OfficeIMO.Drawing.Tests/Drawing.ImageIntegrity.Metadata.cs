using System;
using System.Linq;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void PngContainerRequiresOnePositiveGammaChunkBeforePaletteAndImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] gamma = { 0, 0, 0xB1, 0x8F };
        byte[] withGamma = InsertPngChunkBefore(png, "IDAT", "gAMA", gamma);
        byte[] duplicate = InsertPngChunkBefore(withGamma, "IDAT", "gAMA", gamma);
        byte[] misplaced = InsertPngChunkBefore(png, "IEND", "gAMA", gamma);
        byte[] zero = InsertPngChunkBefore(png, "IDAT", "gAMA", new byte[4]);
        byte[] wrongLength = InsertPngChunkBefore(png, "IDAT", "gAMA", new byte[3]);
        byte[] maximum = InsertPngChunkBefore(png, "IDAT", "gAMA", new byte[] { 0x7F, 0xFF, 0xFF, 0xFF });
        byte[] outOfRange = InsertPngChunkBefore(png, "IDAT", "gAMA", new byte[] { 0x80, 0, 0, 0 });

        Assert.True(OfficeImageReader.TryValidateContent(withGamma, "gamma.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-gamma.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(misplaced, "misplaced-gamma.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(zero, "zero-gamma.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(wrongLength, "short-gamma.png", out _));
        Assert.True(OfficeImageReader.TryValidateContent(maximum, "maximum-gamma.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(outOfRange, "out-of-range-gamma.png", out _));
    }

    [Fact]
    public void PngContainerRequiresOneCompleteChromaticitiesChunkBeforePaletteAndImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] chromaticities = new byte[32];
        int[] coordinates = { 31270, 32900, 64000, 33000, 30000, 60000, 15000, 6000 };
        for (int index = 0; index < coordinates.Length; index++) {
            WriteBigEndianInt32(chromaticities, index * 4, coordinates[index]);
        }
        byte[] withChromaticities = InsertPngChunkBefore(png, "IDAT", "cHRM", chromaticities);
        byte[] duplicate = InsertPngChunkBefore(withChromaticities, "IDAT", "cHRM", chromaticities);
        byte[] misplaced = InsertPngChunkBefore(png, "IEND", "cHRM", chromaticities);
        byte[] wrongLength = InsertPngChunkBefore(png, "IDAT", "cHRM", new byte[31]);
        byte[] outOfIntegerRange = (byte[])chromaticities.Clone();
        WriteBigEndianInt32(outOfIntegerRange, 0, unchecked((int)0xFFFFFFFF));
        byte[] impossiblePair = (byte[])chromaticities.Clone();
        WriteBigEndianInt32(impossiblePair, 0, 80000);
        WriteBigEndianInt32(impossiblePair, 4, 30000);
        byte[] zeroWhiteY = (byte[])chromaticities.Clone();
        WriteBigEndianInt32(zeroWhiteY, 4, 0);

        Assert.True(OfficeImageReader.TryValidateContent(withChromaticities, "chromaticities.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-chromaticities.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(misplaced, "misplaced-chromaticities.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(wrongLength, "short-chromaticities.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "cHRM", outOfIntegerRange),
            "large-chromaticities.png",
            out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "cHRM", impossiblePair),
            "impossible-chromaticities.png",
            out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "cHRM", zeroWhiteY),
            "zero-white-y.png",
            out _));
    }

    [Fact]
    public void PngContainerValidatesSignificantBitsBackgroundAndPaletteHistogramMetadata() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] withSignificantBits = InsertPngChunkBefore(
            png, "IDAT", "sBIT", new byte[] { 8, 8, 8, 8 });
        byte[] invalidSignificantBits = InsertPngChunkBefore(
            png, "IDAT", "sBIT", new byte[] { 9, 8, 8, 8 });
        byte[] withBackground = InsertPngChunkBefore(
            png, "IDAT", "bKGD", new byte[] { 0, 0xFF, 0, 0xFF, 0, 0xFF });
        byte[] invalidBackground = InsertPngChunkBefore(
            png, "IDAT", "bKGD", new byte[] { 0, 0xFF });
        byte[] histogramWithoutPalette = InsertPngChunkBefore(
            png, "IDAT", "hIST", new byte[] { 0, 1 });

        Assert.True(OfficeImageReader.TryValidateContent(withSignificantBits, "significant-bits.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidSignificantBits, "invalid-significant-bits.png", out _));
        Assert.True(OfficeImageReader.TryValidateContent(withBackground, "background.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidBackground, "invalid-background.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(histogramWithoutPalette, "histogram.png", out _));
    }

    [Fact]
    public void PngContainerRequiresOneWellFormedIccProfileBeforePaletteAndImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] compressedProfile = OfficeZlibCodec.Compress(CreateMinimalIccProfile());
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
        byte[] malformedProfile = CreateMinimalIccProfile();
        malformedProfile[36] = (byte)'X';
        byte[] malformedPayload = new byte[] { (byte)'P', (byte)'r', (byte)'o', (byte)'f', (byte)'i', (byte)'l', (byte)'e', 0, 0 }
            .Concat(OfficeZlibCodec.Compress(malformedProfile))
            .ToArray();

        Assert.True(OfficeImageReader.TryValidateContent(withProfile, "profile.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-profile.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(misplaced, "misplaced-profile.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(conflictingColorProfiles, "conflicting-profiles.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidName, "invalid-name.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidMethod, "invalid-method.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalidStream, "invalid-stream.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "iCCP", malformedPayload),
            "malformed-profile.png",
            out _));
    }

    [Fact]
    public void PngContainerRequiresOneStructurallyValidExifTiffPayload() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] exif = {
            (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
            0, 0,
            0, 0, 0, 0
        };
        byte[] withExif = InsertPngChunkBefore(png, "IDAT", "eXIf", exif);
        byte[] duplicate = InsertPngChunkBefore(withExif, "IEND", "eXIf", exif);
        byte[] badOffset = (byte[])exif.Clone();
        badOffset[4] = 0x40;
        byte[] malformed = InsertPngChunkBefore(png, "IDAT", "eXIf", badOffset);

        Assert.True(OfficeImageReader.TryValidateContent(withExif, "exif.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-exif.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(malformed, "malformed-exif.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "eXIf", Array.Empty<byte>()),
            "empty-exif.png",
            out _));
    }

    [Fact]
    public void PngContainerValidatesTextKeywordsAndCompressedTextStreams() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] compressed = OfficeZlibCodec.Compress(Encoding.UTF8.GetBytes("OfficeIMO"));
        byte[] zText = Encoding.ASCII.GetBytes("Comment")
            .Concat(new byte[] { 0, 0 })
            .Concat(compressed)
            .ToArray();
        byte[] internationalText = Encoding.ASCII.GetBytes("Description")
            .Concat(new byte[] { 0, 1, 0, 0, 0 })
            .Concat(compressed)
            .ToArray();
        byte[] invalidMethod = (byte[])zText.Clone();
        invalidMethod[8] = 1;
        byte[] invalidStream = (byte[])zText.Clone();
        invalidStream[invalidStream.Length - 1] ^= 0x01;
        byte[] invalidInternationalStream = (byte[])internationalText.Clone();
        invalidInternationalStream[invalidInternationalStream.Length - 1] ^= 0x01;

        Assert.True(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "zTXt", zText), "compressed-text.png", out _));
        Assert.True(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "iTXt", internationalText), "international-text.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "zTXt", Array.Empty<byte>()), "empty-text.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "zTXt", invalidMethod), "text-method.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "zTXt", invalidStream), "text-stream.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "iTXt", invalidInternationalStream), "international-stream.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "tEXt", new byte[] { (byte)' ', 0 }), "text-keyword.png", out _));
    }

    [Fact]
    public void PngContainerRequiresOneValidModificationTime() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] leapSecond = { 0x07, 0xE8, 2, 29, 23, 59, 60 };
        byte[] beforeImageData = InsertPngChunkBefore(png, "IDAT", "tIME", leapSecond);
        byte[] afterImageData = InsertPngChunkBefore(png, "IEND", "tIME", leapSecond);
        byte[] duplicate = InsertPngChunkBefore(beforeImageData, "IEND", "tIME", leapSecond);
        byte[] invalidCalendarDate = { 0x07, 0xE7, 2, 29, 12, 0, 0 };
        byte[] invalidClock = { 0x07, 0xE8, 1, 1, 24, 0, 0 };

        Assert.True(OfficeImageReader.TryValidateContent(beforeImageData, "modified.png", out _));
        Assert.True(OfficeImageReader.TryValidateContent(afterImageData, "modified-after-data.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(duplicate, "duplicate-modified.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "tIME", Array.Empty<byte>()), "empty-modified.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "tIME", invalidCalendarDate), "invalid-date.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(
            InsertPngChunkBefore(png, "IDAT", "tIME", invalidClock), "invalid-clock.png", out _));
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

    private static byte[] CreateMinimalIccProfile() {
        var profile = new byte[132];
        WriteBigEndianInt32(profile, 0, profile.Length);
        Encoding.ASCII.GetBytes("acsp", 0, 4, profile, 36);
        return profile;
    }
}
