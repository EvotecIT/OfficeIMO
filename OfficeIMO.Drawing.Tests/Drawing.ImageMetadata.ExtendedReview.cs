using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class ExtendedMetadataReviewTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void TiffIptcPreservesExactRecordBytesIncludingTrailingZerosDuringJpegTransfer(int extraBytes) {
        byte[] iptc = new byte[5 + extraBytes]; iptc[0] = 0x1C; iptc[1] = 2; iptc[2] = 120; iptc[4] = (byte)extraBytes;
        byte[] tiff = OfficeTiffCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Blue));
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(tiff); metadata.IptcProfile = iptc;
        byte[] annotated = OfficeImageMetadata.Apply(tiff, metadata);
        metadata = OfficeImageMetadata.Read(annotated); Assert.Equal(iptc, metadata.IptcProfile);
        metadata.SetExifValue(OfficeExifTag.Artist, "Preserved profile");
        annotated = OfficeImageMetadata.Apply(annotated, metadata); Assert.Equal(iptc, OfficeImageMetadata.Read(annotated).IptcProfile);
        byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Blue));
        Assert.Equal(iptc, OfficeImageMetadata.Read(OfficeImageMetadata.Apply(jpeg, metadata)).IptcProfile);
    }

    [Theory]
    [InlineData(OfficeImageResolutionUnit.PixelsPerCentimeter, false)]
    [InlineData(OfficeImageResolutionUnit.AspectRatio, false)]
    [InlineData(OfficeImageResolutionUnit.PixelsPerCentimeter, true)]
    public void BoundedAndStreamingTiffEncodingPreserveNativeResolutionAndExplicitOverrides(OfficeImageResolutionUnit unit, bool explicitDpi) {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Blue);
        var options = new OfficeRasterEncodingOptions { Tiff = new OfficeTiffEncodeOptions { Resolution = new OfficeImageResolution(60, 50, unit) } };
        if (explicitDpi) options.Resolution = new OfficeImageResolution(144D, 127D);
        byte[] unbounded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options);
        byte[] bounded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, options, 1024 * 1024);
        using var stream = new MemoryStream(); OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Tiff, stream, options, 1024 * 1024);
        Assert.Equal(unbounded, bounded); Assert.Equal(unbounded, stream.ToArray());
        var read = OfficeImageMetadata.Read(bounded);
        Assert.Equal(explicitDpi ? OfficeImageResolutionUnit.PixelsPerInch : unit, read.ResolutionUnits);
        Assert.Equal(explicitDpi ? 144D : 60D, read.HorizontalResolution);
        Assert.Equal(explicitDpi ? 127D : 50D, read.VerticalResolution);
        Assert.Equal(unit, options.Tiff.Resolution!.Unit); Assert.Equal(60D, options.Tiff.Resolution.Horizontal);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Webp)]
    [InlineData(OfficeImageExportFormat.Bmp)]
    public void OtherEncoderOverloadsRetainFormatSpecificDensityAndCallerSettings(OfficeImageExportFormat format) {
        var options = new OfficeRasterEncodingOptions(); options.Png.DpiX = 144; options.Png.DpiY = 120; options.Jpeg.DpiX = 144; options.Jpeg.DpiY = 120; options.Webp.DpiX = 144; options.Webp.DpiY = 120;
        if (format == OfficeImageExportFormat.Bmp) { options.Resolution = new OfficeImageResolution(144D, 120D); }
        var image = new OfficeRasterImage(2, 1, OfficeColor.Blue);
        byte[] first = OfficeRasterImageEncoder.Encode(image, format, options);
        byte[] second = OfficeRasterImageEncoder.Encode(image, format, options, 1024 * 1024);
        using var stream = new MemoryStream(); OfficeRasterImageEncoder.EncodeTo(image, format, stream, options, 1024 * 1024);
        foreach (byte[] result in new[] { first, second, stream.ToArray() }) {
            Assert.Equal(144D, OfficeImageMetadata.Read(result).PhysicalDpiX!.Value, 1);
            Assert.Equal(120D, OfficeImageMetadata.Read(result).PhysicalDpiY!.Value, 1);
            Assert.True(OfficeRasterImageDecoder.TryDecode(result, out OfficeRasterImage? decoded));
            Assert.Equal(2, decoded!.Width); Assert.Equal(1, decoded.Height);
        }
        Assert.Equal(120D, options.Png.DpiY);
    }

    [Theory]
    [InlineData("header")]
    [InlineData("truncated")]
    [InlineData("size")]
    [InlineData("offset")]
    [InlineData("palette")]
    [InlineData("rle_end")]
    [InlineData("rle_run")]
    [InlineData("rle_delta")]
    [InlineData("rle_absolute")]
    public void BmpMetadataRejectsIncompleteAndOutOfBoundsPayloads(string defect) {
        byte[] bmp = BitmapFixture(8, 1, 40);
        if (defect == "header") { bmp = new byte[54]; bmp[0] = 66; bmp[1] = 77; Put(bmp, 2, 54, 4); Put(bmp, 10, 54, 4); Put(bmp, 14, 40, 4); Put(bmp, 18, 3, 4); Put(bmp, 22, 2, 4); Put(bmp, 26, 1, 2); Put(bmp, 28, 24, 2); }
        else if (defect == "truncated") { bmp = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red), OfficeImageExportFormat.Bmp); Array.Resize(ref bmp, bmp.Length - 1); Put(bmp, 2, (uint)bmp.Length, 4); }
        else if (defect == "size") Put(bmp, 2, (uint)bmp.Length + 1, 4);
        else if (defect == "offset") Put(bmp, 10, 20, 4);
        else if (defect == "palette") Put(bmp, 46, 256, 4);
        else if (defect == "rle_end") bmp[bmp.Length - 1] = 0;
        else if (defect == "rle_run") bmp[62] = 4;
        else if (defect == "rle_delta") { bmp[62] = 0; bmp[63] = 2; bmp[64] = 4; bmp[65] = 0; }
        else if (defect == "rle_absolute") { bmp[62] = 0; bmp[63] = 255; }
        Assert.True(OfficeImageReader.TryIdentifyByContent(bmp, null, out _));
        Assert.Throws<FormatException>(() => OfficeImageMetadata.Read(bmp));
        Assert.Throws<FormatException>(() => OfficeImageMetadata.Apply(bmp, new OfficeImageMetadata()));
        Assert.Throws<FormatException>(() => OfficeImageMetadata.Remove(bmp, OfficeImageMetadataProfileKinds.All));
    }

    [Theory]
    [InlineData(1, 0, 12)]
    [InlineData(4, 0, 40)]
    [InlineData(8, 0, 40)]
    [InlineData(4, 2, 40)]
    [InlineData(8, 1, 40)]
    [InlineData(16, 0, 40)]
    [InlineData(16, 3, 40)]
    [InlineData(32, 6, 56)]
    [InlineData(24, 0, 64)]
    [InlineData(24, 0, 108)]
    public void BmpMetadataPreservesCompleteLayoutsOutsideTheManagedPixelDecoder(int bits, int compression, int header) {
        byte[] bmp = BitmapFixture(bits, compression, header);
        Assert.True(OfficeImageReader.TryIdentifyByContent(bmp, null, out _));
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(bmp);
        byte[] removed = OfficeImageMetadata.Remove(bmp, OfficeImageMetadataProfileKinds.All).EncodedBytes;
        Assert.Equal(bmp, removed);
        if (header >= 40) {
            metadata.HorizontalResolution = 6000; metadata.VerticalResolution = 5000; metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerMeter;
            byte[] applied = OfficeImageMetadata.Apply(bmp, metadata);
            Assert.Equal(Slice(bmp, (int)Read(bmp, 10, 4), bmp.Length - (int)Read(bmp, 10, 4)), Slice(applied, (int)Read(applied, 10, 4), applied.Length - (int)Read(applied, 10, 4)));
            Assert.Equal(6000D, OfficeImageMetadata.Read(applied).HorizontalResolution);
        }
    }

    [Theory]
    [InlineData(4)]
    [InlineData(5)]
    public void BmpMetadataChecksEmbeddedJpegAndPngStorageWithoutChangingTheirPayloads(int compression) {
        var raster = new OfficeRasterImage(2, 1, OfficeColor.Blue);
        byte[] payload = OfficeRasterImageEncoder.Encode(raster, compression == 4 ? OfficeImageExportFormat.Jpeg : OfficeImageExportFormat.Png);
        var bmp = new byte[54 + payload.Length]; bmp[0] = 66; bmp[1] = 77;
        Put(bmp, 2, (uint)bmp.Length, 4); Put(bmp, 10, 54, 4); Put(bmp, 14, 40, 4); Put(bmp, 18, 2, 4); Put(bmp, 22, 1, 4); Put(bmp, 26, 1, 2);
        Put(bmp, 30, (uint)compression, 4); Put(bmp, 34, (uint)payload.Length, 4); Buffer.BlockCopy(payload, 0, bmp, 54, payload.Length);
        OfficeImageMetadata metadata = OfficeImageMetadata.Read(bmp); metadata.HorizontalResolution = 144;
        byte[] updated = OfficeImageMetadata.Apply(bmp, metadata);
        Assert.Equal(payload, Slice(updated, 54, payload.Length));
        Assert.Equal(updated, OfficeImageMetadata.Remove(updated, OfficeImageMetadataProfileKinds.All).EncodedBytes);
        Array.Resize(ref bmp, bmp.Length - 1); Put(bmp, 2, (uint)bmp.Length, 4); Put(bmp, 34, (uint)(payload.Length - 1), 4);
        Assert.True(OfficeImageReader.TryIdentifyByContent(bmp, null, out _));
        Assert.Throws<FormatException>(() => OfficeImageMetadata.Read(bmp));
        Assert.Throws<FormatException>(() => OfficeImageMetadata.Apply(bmp, metadata));
    }

    [Fact]
    public void BmpRleWithoutDeclaredImageSizePreservesProfileAndCompleteCommandRanges() {
        byte[] original = BitmapFixture(8, 1, 40);
        int pixels = (int)Read(original, 10, 4); Array.Resize(ref original, pixels + 4);
        original[pixels] = 3; original[pixels + 1] = 1; original[pixels + 2] = 0; original[pixels + 3] = 1;
        Put(original, 2, (uint)original.Length, 4); Put(original, 34, 0, 4);
        var metadata = OfficeImageMetadata.Read(original);
        var icc = new byte[132]; icc[3] = 132; icc[36] = 97; icc[37] = 99; icc[38] = 115; icc[39] = 112; metadata.IccProfile = icc;
        byte[] annotated = OfficeImageMetadata.Apply(original, metadata);
        Assert.Equal(icc, OfficeImageMetadata.Read(annotated).IccProfile);
        byte[] removed = OfficeImageMetadata.Remove(annotated, OfficeImageMetadataProfileKinds.Icc).EncodedBytes;
        Assert.Null(OfficeImageMetadata.Read(removed).IccProfile);
        Assert.Equal(new byte[] { 3, 1, 0, 1 }, Slice(removed, (int)Read(removed, 10, 4), 4));
    }

    private static byte[] BitmapFixture(int bits, int compression, int header) {
        int colors = bits <= 8 ? 2 : 0;
        int masks = header == 40 && compression is 3 or 6 ? compression == 6 ? 16 : 12 : 0;
        int offset = 14 + header + masks + colors * (header == 12 ? 3 : 4);
        byte[] payload = compression == 1 ? new byte[] { 3, 1, 0, 0, 0, 3, 0, 1, 0, 0, 0, 1 } : compression == 2 ? new byte[] { 3, 0x10, 0, 0, 0, 3, 0x01, 0, 0, 1 } : new byte[((3 * bits + 31) / 32) * 4 * 2];
        var bmp = new byte[offset + payload.Length]; bmp[0] = 66; bmp[1] = 77; Put(bmp, 2, (uint)bmp.Length, 4); Put(bmp, 10, (uint)offset, 4); Put(bmp, 14, (uint)header, 4);
        Put(bmp, 18, 3, header == 12 ? 2 : 4); Put(bmp, header == 12 ? 20 : 22, 2, header == 12 ? 2 : 4); Put(bmp, header == 12 ? 22 : 26, 1, 2); Put(bmp, header == 12 ? 24 : 28, (uint)bits, 2);
        if (header >= 40) { Put(bmp, 30, (uint)compression, 4); Put(bmp, 34, (uint)payload.Length, 4); Put(bmp, 46, (uint)colors, 4); }
        if (compression is 3 or 6) { Put(bmp, 54, bits == 16 ? 0xF800U : 0x00FF0000U, 4); Put(bmp, 58, bits == 16 ? 0x07E0U : 0x0000FF00U, 4); Put(bmp, 62, bits == 16 ? 0x001FU : 0x000000FFU, 4); if (compression == 6) Put(bmp, 66, 0xFF000000, 4); }
        Buffer.BlockCopy(payload, 0, bmp, offset, payload.Length); return bmp;
    }
    private static byte[] Slice(byte[] bytes, int at, int count) { var result = new byte[count]; Buffer.BlockCopy(bytes, at, result, 0, count); return result; }
    private static uint Read(byte[] bytes, int at, int size) { uint result = 0; for (int i = 0; i < size; i++) result |= (uint)bytes[at + i] << (8 * i); return result; }
    private static void Put(byte[] bytes, int at, uint value, int size) { for (int i = 0; i < size; i++) bytes[at + i] = (byte)(value >> (8 * i)); }
}
