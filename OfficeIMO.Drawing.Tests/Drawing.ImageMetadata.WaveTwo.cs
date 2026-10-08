using System;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class ImageMetadataWaveTwoTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg, 0.5, 1)]
    [InlineData(OfficeImageExportFormat.Png, 0.5, 1)]
    [InlineData(OfficeImageExportFormat.Jpeg, 0.25, 0.75)]
    [InlineData(OfficeImageExportFormat.Png, 0.25, 0.75)]
    [InlineData(OfficeImageExportFormat.Jpeg, 1.5, 2.5)]
    [InlineData(OfficeImageExportFormat.Png, 1.5, 2.5)]
    [InlineData(OfficeImageExportFormat.Jpeg, 7, 3)]
    [InlineData(OfficeImageExportFormat.Png, 7, 3)]
    [InlineData(OfficeImageExportFormat.Jpeg, 1000000, 2000000)]
    [InlineData(OfficeImageExportFormat.Png, 1000000, 2000000)]
    [InlineData(OfficeImageExportFormat.Jpeg, 1D / 65535D, 1)]
    [InlineData(OfficeImageExportFormat.Png, 1D / 4294967295D, 1)]
    public void IntegerDensityCarriersPreserveRepresentableFractionalAspectRatios(OfficeImageExportFormat format, double x, double y) {
        byte[] original = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Blue), format);
        var metadata = new OfficeImageMetadata { ResolutionUnits = OfficeImageResolutionUnit.AspectRatio, HorizontalResolution = x, VerticalResolution = y };
        metadata.SetExifValue(OfficeExifTag.Artist, "Ratio author");
        byte[] result = OfficeImageMetadata.Apply(original, metadata);
        var actual = OfficeImageMetadata.Read(result);
        Assert.Equal(OfficeImageResolutionUnit.AspectRatio, actual.ResolutionUnits);
        Assert.Null(actual.PhysicalDpiX); Assert.Null(actual.PhysicalDpiY);
        Assert.True(Math.Abs(actual.HorizontalResolution / actual.VerticalResolution - x / y) <= x / y * 1E-12);
        if (x == 7D && y == 3D) { Assert.Equal(x, actual.HorizontalResolution); Assert.Equal(y, actual.VerticalResolution); }
        AssertPixels(original, result);
        actual.SetExifValue(OfficeExifTag.Software, "Unrelated edit");
        var repeated = OfficeImageMetadata.Read(OfficeImageMetadata.Apply(result, actual));
        Assert.Equal(actual.HorizontalResolution, repeated.HorizontalResolution); Assert.Equal(actual.VerticalResolution, repeated.VerticalResolution);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg, 1E-20, 1)]
    [InlineData(OfficeImageExportFormat.Png, 1E-20, 1)]
    [InlineData(OfficeImageExportFormat.Jpeg, 1E300, 1E-300)]
    [InlineData(OfficeImageExportFormat.Png, 1E300, 1E-300)]
    public void UnrepresentableAspectRatiosFailWithoutChangingCallerInput(OfficeImageExportFormat format, double x, double y) {
        byte[] original = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1), format); byte[] expected = (byte[])original.Clone();
        var metadata = new OfficeImageMetadata { ResolutionUnits = OfficeImageResolutionUnit.AspectRatio, HorizontalResolution = x, VerticalResolution = y };
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageMetadata.Apply(original, metadata));
        Assert.Equal(expected, original); Assert.Equal(x, metadata.HorizontalResolution); Assert.Equal(y, metadata.VerticalResolution);
    }

    [Theory]
    [InlineData(true, 1)]
    [InlineData(false, 1)]
    [InlineData(true, 2)]
    [InlineData(false, 2)]
    [InlineData(true, 4)]
    [InlineData(false, 4)]
    [InlineData(true, 8)]
    [InlineData(false, 8)]
    public void JfifIsFirstAfterSoiForEachReplacementProfileAndSynthesizedHeader(bool existing, int family) {
        byte[] original = OfficeJpegCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Blue), new OfficeJpegEncodeOptions { WriteJfifHeader = existing });
        if (existing) original = AddJfifThumbnail(original);
        var metadata = Profiles(family);
        byte[] result = OfficeImageMetadata.Apply(original, metadata);
        AssertFirstJfif(result);
        if (existing) Assert.Equal(Jfif(original), Jfif(result));
        Assert.Equal(Scan(original), Scan(result)); AssertPixels(original, result);
        var read = OfficeImageMetadata.Read(result);
        if ((family & 1) != 0) Assert.Equal("JFIF creator", read.GetExifValue(OfficeExifTag.Artist)!.Value);
        if ((family & 2) != 0) Assert.Equal(metadata.IptcProfile, read.IptcProfile);
        if ((family & 4) != 0) Assert.Equal(metadata.XmpProfile, read.XmpProfile);
        if ((family & 8) != 0) Assert.Equal(metadata.IccProfile, read.IccProfile);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void JfifPlacementAndEntropySurviveAllProfilesProgressiveScansAndRemoval(bool existing, bool progressive) {
        byte[] original = OfficeJpegCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red), new OfficeJpegEncodeOptions { WriteJfifHeader = existing, Progressive = progressive });
        byte[] result = OfficeImageMetadata.Apply(original, Profiles(15)); AssertFirstJfif(result);
        Assert.Equal(Scan(original), Scan(result)); AssertPixels(original, result);
        byte[] removed = OfficeImageMetadata.Remove(result, OfficeImageMetadataProfileKinds.All).EncodedBytes;
        AssertFirstJfif(removed); Assert.Equal(Scan(original), Scan(removed)); AssertPixels(original, removed);
        Assert.False(OfficeImageMetadata.Read(removed).HasExifProfile);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void OptionalZeroTiffDensityDoesNotPoisonMetadataResolutionOrUnrelatedEdits(int invalidAxis) {
        byte[] original = OfficeTiffCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Green), new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(60, 50, OfficeImageResolutionUnit.PixelsPerCentimeter) });
        byte[] expectedPixels = (byte[])original.Clone();
        if (invalidAxis == 0 || invalidAxis == 2) ZeroRationalWord(original, 282, false);
        if (invalidAxis == 1 || invalidAxis == 2) ZeroRationalWord(original, 283, false);
        if (invalidAxis == 3) ZeroRationalWord(original, 282, true);
        Assert.True(OfficeTiffStructureValidator.TryValidate(original, 0, original.Length));
        AssertPixels(expectedPixels, original);
        Assert.True(OfficeRasterContainerInspector.TryInspect(original, out OfficeRasterContainerInfo? inventory));
        Assert.Null(inventory!.Frames[0].DpiX); Assert.Null(inventory.Frames[0].DpiY);
        var metadata = OfficeImageMetadata.Read(original);
        Assert.True(metadata.Resolution.Horizontal > 0 && metadata.Resolution.Vertical > 0);
        Assert.True(metadata.PhysicalDpiX > 0); Assert.True(metadata.PhysicalDpiY > 0);
        metadata.SetExifValue(OfficeExifTag.Artist, "Optional density");
        byte[] result = OfficeImageMetadata.Apply(original, metadata); AssertPixels(expectedPixels, result);
        var actual = OfficeImageMetadata.Read(result);
        Assert.Equal("Optional density", actual.GetExifValue(OfficeExifTag.Artist)!.Value);
        Assert.True(actual.Resolution.Horizontal > 0 && actual.PhysicalDpiY > 0);
    }

    private static OfficeImageMetadata Profiles(int families) {
        var metadata = new OfficeImageMetadata();
        if ((families & 1) != 0) metadata.SetExifValue(OfficeExifTag.Artist, "JFIF creator");
        if ((families & 2) != 0) metadata.IptcProfile = new byte[] { 0x1C, 2, 120, 0, 1, 0 };
        if ((families & 4) != 0) metadata.XmpProfile = Encoding.UTF8.GetBytes("<x:xmpmeta xmlns:x='adobe:ns:meta/'/>");
        if ((families & 8) != 0) { byte[] icc = new byte[132]; icc[3] = 132; Encoding.ASCII.GetBytes("acsp", 0, 4, icc, 36); metadata.IccProfile = icc; }
        return metadata;
    }
    private static void AssertPixels(byte[] before, byte[] after) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(before, out OfficeRasterImage? expected));
        Assert.True(OfficeRasterImageDecoder.TryDecode(after, out OfficeRasterImage? actual));
        Assert.Equal(expected!.Width, actual!.Width); Assert.Equal(expected.Height, actual.Height); Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }
    private static void AssertFirstJfif(byte[] bytes) {
        Assert.Equal(0xFF, bytes[2]); Assert.Equal(0xE0, bytes[3]); Assert.Equal("JFIF\0", Encoding.ASCII.GetString(bytes, 6, 5));
    }
    private static byte[] Jfif(byte[] bytes) {
        AssertFirstJfif(bytes); return Slice(bytes, 2, 2 + (bytes[4] << 8 | bytes[5]));
    }
    private static byte[] AddJfifThumbnail(byte[] bytes) {
        byte[] segment = Jfif(bytes); int length = segment.Length;
        Array.Resize(ref segment, length + 3); segment[2] = (byte)((length + 1) >> 8); segment[3] = (byte)(length + 1);
        segment[16] = 1; segment[17] = 1; segment[length] = 10; segment[length + 1] = 80; segment[length + 2] = 160;
        using var output = new MemoryStream(); output.Write(bytes, 0, 2); output.Write(segment, 0, segment.Length); output.Write(bytes, 2 + length, bytes.Length - 2 - length); return output.ToArray();
    }
    private static byte[] Scan(byte[] bytes) {
        for (int cursor = 2; cursor < bytes.Length;) { int marker = bytes[cursor + 1]; if (marker == 0xDA) return Slice(bytes, cursor, bytes.Length - cursor); cursor += 2 + (bytes[cursor + 2] << 8 | bytes[cursor + 3]); }
        throw new InvalidDataException("Missing scan.");
    }
    private static byte[] Slice(byte[] bytes, int at, int length) { var result = new byte[length]; Buffer.BlockCopy(bytes, at, result, 0, length); return result; }
    private static void ZeroRationalWord(byte[] bytes, uint tag, bool denominator) {
        int root = bytes[4] | bytes[5] << 8 | bytes[6] << 16 | bytes[7] << 24;
        int count = bytes[root] | bytes[root + 1] << 8;
        for (int i = 0; i < count; i++) { int entry = root + 2 + i * 12; if ((bytes[entry] | bytes[entry + 1] << 8) != tag) continue; int at = bytes[entry + 8] | bytes[entry + 9] << 8 | bytes[entry + 10] << 16 | bytes[entry + 11] << 24; Array.Clear(bytes, at + (denominator ? 4 : 0), 4); return; }
        throw new InvalidDataException("Missing density.");
    }
}
