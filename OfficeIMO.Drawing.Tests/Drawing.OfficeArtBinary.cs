using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeArtPropertyTableReader_DecodesFixedAndComplexEntries() {
        byte[] payload = {
            0x81, 0x01, 0x33, 0x22, 0x11, 0x00,
            0x80, 0x83, 0x06, 0x00, 0x00, 0x00,
            0x4E, 0x00, 0x61, 0x00, 0x6D, 0x00
        };

        IReadOnlyList<OfficeArtProperty> properties = OfficeArtPropertyTableReader.Read(payload, 2);

        Assert.Equal(2, properties.Count);
        Assert.Equal("fillColor", properties[0].PropertyName);
        Assert.Equal("Fill", properties[0].PropertyGroupName);
        Assert.Equal(0x00112233U, properties[0].Value);
        Assert.True(properties[1].IsComplex);
        Assert.Equal("wzName", properties[1].PropertyName);
        Assert.Equal(6, properties[1].AvailableComplexDataLength);
        Assert.Equal("Nam", properties[1].ComplexText);
        Assert.Equal(new byte[] { 0x4E, 0x00, 0x61, 0x00, 0x6D, 0x00 },
            properties[1].CopyComplexData());
    }

    [Fact]
    public void OfficeArtShapeStyle_DecodesVisibilityColorsAndLineGeometry() {
        byte[] payload = {
            0x81, 0x01, 0x00, 0x00, 0x00, 0x08,
            0x82, 0x01, 0x00, 0x80, 0x00, 0x00,
            0xBF, 0x01, 0x10, 0x00, 0x10, 0x00,
            0xC0, 0x01, 0x33, 0x22, 0x11, 0x00,
            0xCB, 0x01, 0x00, 0x7F, 0x00, 0x00,
            0xCE, 0x01, 0x03, 0x00, 0x00, 0x00,
            0xFF, 0x01, 0x08, 0x00, 0x08, 0x00
        };

        OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(
            OfficeArtPropertyTableReader.Read(payload, 7));

        Assert.True(style.FillEnabled);
        Assert.Equal(0.5D, style.FillOpacity);
        Assert.True(style.LineEnabled);
        Assert.Equal(32512, style.LineWidthEmus);
        Assert.Equal(3U, style.LineDashing);
        Assert.True(style.FillColor!.Value.TryResolve(
            index => index == 0 ? OfficeColor.FromRgb(0xAA, 0xBB, 0xCC) : null,
            out OfficeColor fill));
        Assert.Equal(OfficeColor.FromRgb(0xAA, 0xBB, 0xCC), fill);
        Assert.True(style.LineColor!.Value.TryResolve(null, out OfficeColor line));
        Assert.Equal(OfficeColor.FromRgb(0x33, 0x22, 0x11), line);
    }

    [Fact]
    public void OfficeArtShapeStyle_DecodesProjectableSignedOffsetShadow() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x0200, 0U),
            new OfficeArtProperty(1, 0x0201, 0x00332211U),
            new OfficeArtProperty(2, 0x0204, 32768U),
            new OfficeArtProperty(3, 0x0205, unchecked((uint)-12700)),
            new OfficeArtProperty(4, 0x0206, 25400U),
            new OfficeArtProperty(5, 0x021C, 50800U),
            new OfficeArtProperty(6, 0x023F, 0x00020002U)
        };

        OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(properties);

        Assert.True(style.HasProjectableShadow);
        Assert.False(style.HasUnprojectedVisualStyle);
        Assert.Equal(0U, style.ShadowType);
        Assert.Equal(new OfficeArtColorReference(0x00332211U), style.ShadowColor);
        Assert.Equal(0.5D, style.ShadowOpacity);
        Assert.Equal(-12700, style.ShadowOffsetXEmus);
        Assert.Equal(25400, style.ShadowOffsetYEmus);
        Assert.Equal(50800, style.ShadowSoftnessEmus);
        Assert.Equal("shadowOffsetX", properties[3].PropertyName);
        Assert.Equal("Shadow", properties[3].PropertyGroupName);
    }

    [Fact]
    public void OfficeArtShapeTransform_DecodesSignedRotationAndFspFlips() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x0004, unchecked((uint)(-45 * 65536)))
        };

        OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(
            (1U << 6) | (1U << 7), properties);

        Assert.Equal(-45D, transform.RotationDegrees);
        Assert.True(transform.FlipHorizontal);
        Assert.True(transform.FlipVertical);
        Assert.True(transform.HasTransform);
        Assert.Equal("rotation", properties[0].PropertyName);
        Assert.Equal("Transform", properties[0].PropertyGroupName);
    }

    [Fact]
    public void OfficeArtShapeTransform_HonorsExplicitFlipOverrides() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x033F,
                (1U << 8) | (1U << 9) | (1U << 25))
        };

        OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(
            (1U << 6) | (1U << 7), properties);

        Assert.False(transform.FlipHorizontal);
        Assert.True(transform.FlipVertical);
    }

    [Fact]
    public void OfficeArtShapeGeometry_PreservesSignedShapeSpecificAdjustmentSlots() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x0147, 5400U),
            new OfficeArtProperty(1, 0x0149, unchecked((uint)-2700)),
            new OfficeArtProperty(2, 0x0147, 7200U)
        };

        OfficeArtShapeGeometry geometry = OfficeArtShapeGeometry.Decode(properties);

        Assert.True(geometry.HasAdjustments);
        Assert.Equal(7200, geometry.AdjustmentValues[0]);
        Assert.Null(geometry.AdjustmentValues[1]);
        Assert.Equal(-2700, geometry.AdjustmentValues[2]);
        Assert.Null(geometry.AdjustmentValues[7]);
        Assert.Equal("adjustValue", properties[0].PropertyName);
        Assert.Equal("Geometry", properties[0].PropertyGroupName);
    }

    [Fact]
    public void OfficeArtPictureProperties_DecodesSignedCropFractionsAndUsesLastValue() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x0102, 8192U),
            new OfficeArtProperty(1, 0x0100, unchecked((uint)-4096)),
            new OfficeArtProperty(2, 0x0102, 16384U),
            new OfficeArtProperty(3, 0x0103, 0U)
        };

        OfficeArtPictureProperties picture = OfficeArtPictureProperties.Decode(properties);

        Assert.True(picture.HasExplicitCrop);
        Assert.True(picture.HasCrop);
        Assert.Equal(16384, picture.CropFromLeftRaw);
        Assert.Equal(0.25D, picture.CropFromLeft);
        Assert.Equal(-0.0625D, picture.CropFromTop);
        Assert.Equal(0D, picture.CropFromRight);
        Assert.Null(picture.CropFromBottom);
        Assert.Equal("cropFromLeft", properties[0].PropertyName);
        Assert.Equal("Blip", properties[0].PropertyGroupName);
    }

    [Fact]
    public void OfficeArtPictureProperties_DecodesPictureEffectsAndBooleanUseBits() {
        var properties = new[] {
            new OfficeArtProperty(0, 0x0107, 0x00030201U),
            new OfficeArtProperty(1, 0x0108, 45875U),
            new OfficeArtProperty(2, 0x0109, unchecked((uint)-8192)),
            new OfficeArtProperty(3, 0x011A, 0x00060504U),
            new OfficeArtProperty(4, 0x013F, (1U << 18) | (1U << 2) | (1U << 1))
        };

        OfficeArtPictureProperties picture = OfficeArtPictureProperties.Decode(properties);

        Assert.True(picture.HasPictureEffect);
        Assert.Equal(new OfficeArtColorReference(0x00030201U), picture.TransparentColor);
        Assert.Equal(45875, picture.ContrastRaw);
        Assert.Equal(-8192, picture.BrightnessRaw);
        Assert.Equal(-0.3D, picture.ContrastAdjustment!.Value, 4);
        Assert.Equal(-0.25D, picture.BrightnessAdjustment);
        Assert.Equal(new OfficeArtColorReference(0x00060504U), picture.RecolorColor);
        Assert.True(picture.Grayscale);
        Assert.Null(picture.BiLevel);
        Assert.Equal("pictureBrightness", properties[2].PropertyName);
    }

    [Fact]
    public void OfficeArtPropertyTableReader_RejectsTruncatedFixedTableWithoutOverread() {
        byte[] payload = { 0x81, 0x01, 0x33, 0x22, 0x11 };

        Assert.Empty(OfficeArtPropertyTableReader.Read(payload, 1));
    }

    [Fact]
    public void OfficeArtBlipStoreEntryReader_ResolvesEmbeddedAndDelayedPngRecords() {
        byte[] png = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] blip = BuildBlipRecord(0x06E0, 0xF01E, png);

        byte[] embeddedFbse = BuildFbse(blip, uint.MaxValue);
        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(embeddedFbse, 0x0006,
            out OfficeArtBlipStoreEntry? embedded));
        Assert.NotNull(embedded);
        Assert.Equal(OfficeArtBlipStorage.Embedded, embedded!.Storage);
        Assert.Equal(OfficeArtBlipType.Png, embedded.RecordInstanceBlipType);
        Assert.Equal("OfficeArtBlipPNG", embedded.BlipRecordTypeName);
        Assert.Equal("image/png", embedded.ContentType);
        Assert.Equal(png, embedded.ImageBytes);
        Assert.True(embedded.HasImportableImage);

        byte[] delayedFbse = BuildFbse(Array.Empty<byte>(), 0);
        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(delayedFbse, 0,
            delayedFbse.Length, 0x0006, blip, out OfficeArtBlipStoreEntry? delayed));
        Assert.NotNull(delayed);
        Assert.Equal(OfficeArtBlipStorage.Delayed, delayed!.Storage);
        Assert.Equal(png, delayed.ImageBytes);
        Assert.Equal(25, delayed.BlipPayloadAvailableLength);
        Assert.False(delayed.IsPayloadTruncated);
    }

    [Fact]
    public void OfficeArtBlipStoreEntryReader_UsesTwoUidRasterPrefixAndBuildsBmpFileHeader() {
        byte[] dib = new byte[44];
        WriteOfficeArtUInt32(dib, 0, 40);
        WriteOfficeArtUInt32(dib, 4, 1);
        WriteOfficeArtUInt32(dib, 8, 1);
        dib[12] = 1;
        dib[14] = 24;
        dib[40] = 0x11;
        dib[41] = 0x22;
        dib[42] = 0x33;
        byte[] blip = BuildBlipRecord(0x07A9, 0xF01F, dib, twoUids: true);

        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(BuildFbse(blip, uint.MaxValue), 0x0007,
            out OfficeArtBlipStoreEntry? entry));

        Assert.NotNull(entry);
        Assert.Equal("image/bmp", entry!.ContentType);
        Assert.Equal(58, entry.ImageBytes.Length);
        Assert.Equal(0x42, entry.ImageBytes[0]);
        Assert.Equal(0x4D, entry.ImageBytes[1]);
        Assert.Equal(54U, ReadOfficeArtUInt32(entry.ImageBytes, 10));
    }

    [Fact]
    public void OfficeArtBlipStoreEntryReader_ExtractsUncompressedMetafilePayload() {
        byte[] metafile = { 0x01, 0x02, 0x03, 0x04 };
        byte[] header = new byte[34];
        WriteOfficeArtUInt32(header, 0, unchecked((uint)metafile.Length));
        WriteOfficeArtUInt32(header, 28, unchecked((uint)metafile.Length));
        header[32] = 0xFE;
        header[33] = 0xFE;
        byte[] blip = BuildBlipRecord(0x03D4, 0xF01A, header.Concat(metafile).ToArray(),
            includeTag: false);

        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(BuildFbse(blip, uint.MaxValue), 0x0002,
            out OfficeArtBlipStoreEntry? entry));

        Assert.NotNull(entry);
        Assert.Equal("image/x-emf", entry!.ContentType);
        Assert.Equal(metafile, entry.ImageBytes);
    }

    [Theory]
    [InlineData(0xF01D, 0x046A, "image/jpeg")]
    [InlineData(0xF02A, 0x06E2, "image/jpeg")]
    [InlineData(0xF029, 0x06E4, "image/tiff")]
    public void OfficeArtBlipStoreEntryReader_ExtractsSupportedRasterFormats(
        ushort recordType, ushort recordInstance, string contentType) {
        byte[] image = recordType == 0xF029
            ? new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x01 }
            : new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 };
        byte[] blip = BuildBlipRecord(recordInstance, recordType, image);

        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(BuildFbse(blip, uint.MaxValue), 0x0005,
            out OfficeArtBlipStoreEntry? entry));

        Assert.NotNull(entry);
        Assert.Equal(contentType, entry!.ContentType);
        Assert.Equal(image, entry.ImageBytes);
    }

    private static byte[] BuildFbse(byte[] embeddedBlip, uint delayedOffset) {
        byte[] payload = new byte[36 + embeddedBlip.Length];
        payload[0] = 0x06;
        payload[1] = 0x06;
        for (int index = 0; index < 16; index++) payload[2 + index] = unchecked((byte)index);
        WriteOfficeArtUInt32(payload, 20, unchecked((uint)embeddedBlip.Length));
        WriteOfficeArtUInt32(payload, 24, 1);
        WriteOfficeArtUInt32(payload, 28, delayedOffset);
        Buffer.BlockCopy(embeddedBlip, 0, payload, 36, embeddedBlip.Length);
        return payload;
    }

    private static byte[] BuildBlipRecord(ushort instance, ushort type, byte[] imageData,
        bool twoUids = false, bool includeTag = true) {
        int uidLength = twoUids ? 32 : 16;
        byte[] payload = new byte[uidLength + (includeTag ? 1 : 0) + imageData.Length];
        if (includeTag) payload[uidLength] = 0xFF;
        Buffer.BlockCopy(imageData, 0, payload, uidLength + (includeTag ? 1 : 0), imageData.Length);
        byte[] record = new byte[8 + payload.Length];
        ushort versionAndInstance = unchecked((ushort)(instance << 4));
        record[0] = unchecked((byte)versionAndInstance);
        record[1] = unchecked((byte)(versionAndInstance >> 8));
        record[2] = unchecked((byte)type);
        record[3] = unchecked((byte)(type >> 8));
        WriteOfficeArtUInt32(record, 4, unchecked((uint)payload.Length));
        Buffer.BlockCopy(payload, 0, record, 8, payload.Length);
        return record;
    }

    private static uint ReadOfficeArtUInt32(byte[] source, int offset) => unchecked((uint)(
        source[offset] | source[offset + 1] << 8 | source[offset + 2] << 16 | source[offset + 3] << 24));

    private static void WriteOfficeArtUInt32(byte[] target, int offset, uint value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
        target[offset + 2] = unchecked((byte)(value >> 16));
        target[offset + 3] = unchecked((byte)(value >> 24));
    }
}
