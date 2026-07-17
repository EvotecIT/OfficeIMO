using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingOfficeArtBlipTests {
    [Fact]
    public void PropertyTableReaderDecodesFixedAndComplexEntries() {
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
    public void ShapeStyleDecodesVisibilityColorsAndLineGeometry() {
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
    public void PropertyTableReaderRejectsTruncatedFixedTableWithoutOverread() {
        byte[] payload = { 0x81, 0x01, 0x33, 0x22, 0x11 };

        Assert.Empty(OfficeArtPropertyTableReader.Read(payload, 1));
    }

    [Fact]
    public void ReaderResolvesEmbeddedAndDelayedPngRecords() {
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
    public void ReaderRejectsImageBeforeAllocatingPastDecodedLimit() {
        byte[] png = {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A
        };
        byte[] blip = BuildBlipRecord(0x06E0, 0xF01E, png);

        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(
            BuildFbse(blip, uint.MaxValue), 0x0006,
            out OfficeArtBlipStoreEntry? entry,
            maximumDecodedImageBytes: png.Length - 1));

        Assert.NotNull(entry);
        Assert.Empty(entry!.ImageBytes);
        Assert.True(entry.WasImageRejectedBySizeLimit);
    }

    [Fact]
    public void ReaderUsesTwoUidRasterPrefixAndBuildsBmpFileHeader() {
        byte[] dib = new byte[44];
        WriteUInt32(dib, 0, 40);
        WriteUInt32(dib, 4, 1);
        WriteUInt32(dib, 8, 1);
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
        Assert.Equal(54U, ReadUInt32(entry.ImageBytes, 10));
    }

    [Fact]
    public void ReaderExtractsUncompressedMetafilePayload() {
        byte[] metafile = { 0x01, 0x02, 0x03, 0x04 };
        byte[] header = new byte[34];
        WriteUInt32(header, 0, unchecked((uint)metafile.Length));
        WriteUInt32(header, 28, unchecked((uint)metafile.Length));
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
    public void ReaderExtractsSupportedRasterFormats(
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

    [Theory]
    [InlineData("", "31D6CFE0D16AE931B73C59D7E0C089C0")]
    [InlineData("a", "BDE52CB31DE33E46245E05FBDBD6FB24")]
    [InlineData("abc", "A448017AAF21D8525FC10AE87AA6729D")]
    [InlineData("message digest", "D9130A8164549FE818874806E1C7014B")]
    public void Md4MatchesRfc1320Vectors(string value, string expected) {
        byte[] digest = OfficeArtMd4.Compute(System.Text.Encoding.ASCII.GetBytes(value));

        Assert.Equal(expected, BitConverter.ToString(digest).Replace("-", string.Empty));
    }

    [Fact]
    public void WriterCreatesEmbeddedPngFbse() {
        byte[] png = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] record = OfficeArtBlipStoreEntryWriter.CreateEmbedded(
            png, "image/png", referenceCount: 3);

        Assert.Equal(2, record[0] & 0x0F);
        Assert.Equal(0xF007, record[2] | record[3] << 8);
        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(record, 8,
            checked((int)ReadUInt32(record, 4)), 0x0006,
            delayStream: null, out OfficeArtBlipStoreEntry? entry));
        Assert.NotNull(entry);
        Assert.Equal(OfficeArtBlipStorage.Embedded, entry!.Storage);
        Assert.Equal(OfficeArtBlipType.Png, entry.RecordInstanceBlipType);
        Assert.Equal(3U, entry.ReferenceCount);
        Assert.Equal(33U, entry.SizeBytes);
        Assert.Equal(png, entry.ImageBytes);
        Assert.Equal("image/png", entry.ContentType);
        Assert.Equal(BitConverter.ToString(OfficeArtMd4.Compute(png)).Replace("-", string.Empty),
            entry.UidHex);
    }

    [Fact]
    public void WriterCreatesDelayedPngFbseAndBlip() {
        byte[] png = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] blip = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(png, "image/png");
        byte[] record = OfficeArtBlipStoreEntryWriter.CreateDelayed(
            png, "image/png", delayedStreamOffset: 0, referenceCount: 3);

        Assert.Equal(36U, ReadUInt32(record, 4));
        Assert.Equal(0xF01E, blip[2] | blip[3] << 8);
        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(record, 8,
            checked((int)ReadUInt32(record, 4)), 0x0006,
            blip, out OfficeArtBlipStoreEntry? entry));
        Assert.NotNull(entry);
        Assert.Equal(OfficeArtBlipStorage.Delayed, entry!.Storage);
        Assert.Equal(0U, entry.DelayedStreamOffset);
        Assert.Equal(3U, entry.ReferenceCount);
        Assert.Equal(checked((uint)blip.Length), entry.SizeBytes);
        Assert.Equal(png, entry.ImageBytes);
        Assert.Equal(BitConverter.ToString(OfficeArtMd4.Compute(png)).Replace("-", string.Empty),
            entry.UidHex);
    }

    [Fact]
    public void WriterCreatesCompressedEmfBlip() {
        byte[] emf = BuildMinimalEmf();
        byte[] blip = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(emf, "image/x-emf");
        byte[] fbse = OfficeArtBlipStoreEntryWriter.CreateDelayed(
            emf, "image/x-emf", delayedStreamOffset: 0);

        Assert.Equal(0xF01A, blip[2] | blip[3] << 8);
        Assert.Equal(0x03D4, (blip[0] | blip[1] << 8) >> 4);
        Assert.Equal(0x02, fbse[8]);
        Assert.Equal(0x04, fbse[9]);
        Assert.Equal(0x00, blip[8 + 16 + 32]);
        Assert.Equal(0xFE, blip[8 + 16 + 33]);
        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(fbse, 8,
            checked((int)ReadUInt32(fbse, 4)), 0x0002,
            blip, out OfficeArtBlipStoreEntry? entry));
        Assert.NotNull(entry);
        Assert.Equal(OfficeArtBlipType.Emf, entry!.RecordInstanceBlipType);
        Assert.Equal("image/x-emf", entry.ContentType);
        Assert.Equal(emf, entry.ImageBytes);
    }

    [Fact]
    public void WriterCreatesCompressedPlaceableWmfBlip() {
        byte[] wmf = BuildMinimalPlaceableWmf();
        byte[] blip = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(wmf, "image/x-wmf");
        byte[] fbse = OfficeArtBlipStoreEntryWriter.CreateDelayed(
            wmf, "image/x-wmf", delayedStreamOffset: 0);

        Assert.Equal(0xF01B, blip[2] | blip[3] << 8);
        Assert.Equal(0x0216, (blip[0] | blip[1] << 8) >> 4);
        Assert.Equal(0x03, fbse[8]);
        Assert.Equal(0x04, fbse[9]);
        Assert.Equal(0x00, blip[8 + 16 + 32]);
        Assert.Equal(0xFE, blip[8 + 16 + 33]);
        Assert.True(OfficeArtBlipStoreEntryReader.TryRead(fbse, 8,
            checked((int)ReadUInt32(fbse, 4)), 0x0003,
            blip, out OfficeArtBlipStoreEntry? entry));
        Assert.NotNull(entry);
        Assert.Equal(OfficeArtBlipType.Wmf, entry!.RecordInstanceBlipType);
        Assert.Equal("image/x-wmf", entry.ContentType);
        Assert.Equal(wmf, entry.ImageBytes);
    }

    [Fact]
    public void WriterRejectsMismatchedOrUnsupportedPayload() {
        Assert.Throws<NotSupportedException>(() =>
            OfficeArtBlipStoreEntryWriter.CreateEmbedded(new byte[] { 1, 2, 3 }, "image/png"));
        Assert.Throws<NotSupportedException>(() =>
            OfficeArtBlipStoreEntryWriter.CreateEmbedded(
                new byte[] { (byte)'G', (byte)'I', (byte)'F' }, "image/gif"));
    }

    private static byte[] BuildFbse(byte[] embeddedBlip, uint delayedOffset) {
        byte[] payload = new byte[36 + embeddedBlip.Length];
        payload[0] = 0x06;
        payload[1] = 0x06;
        for (int index = 0; index < 16; index++) {
            payload[2 + index] = unchecked((byte)index);
        }
        WriteUInt32(payload, 20, unchecked((uint)embeddedBlip.Length));
        WriteUInt32(payload, 24, 1);
        WriteUInt32(payload, 28, delayedOffset);
        Buffer.BlockCopy(embeddedBlip, 0, payload, 36, embeddedBlip.Length);
        return payload;
    }

    private static byte[] BuildMinimalEmf() {
        var result = new byte[108];
        WriteUInt32(result, 0, 1U);
        WriteUInt32(result, 4, 88U);
        WriteUInt32(result, 16, 100U);
        WriteUInt32(result, 20, 100U);
        WriteUInt32(result, 32, 2540U);
        WriteUInt32(result, 36, 2540U);
        WriteUInt32(result, 40, 0x464D4520U);
        WriteUInt32(result, 44, 0x00010000U);
        WriteUInt32(result, 48, checked((uint)result.Length));
        WriteUInt32(result, 52, 2U);
        result[56] = 1;
        WriteUInt32(result, 72, 100U);
        WriteUInt32(result, 76, 100U);
        WriteUInt32(result, 80, 25U);
        WriteUInt32(result, 84, 25U);
        WriteUInt32(result, 88, 14U);
        WriteUInt32(result, 92, 20U);
        WriteUInt32(result, 104, 20U);
        return result;
    }

    private static byte[] BuildMinimalPlaceableWmf() {
        var result = new byte[46];
        WriteUInt32(result, 0, 0x9AC6CDD7U);
        WriteUInt16(result, 10, 1440);
        WriteUInt16(result, 12, 720);
        WriteUInt16(result, 14, 1440);
        ushort checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= unchecked((ushort)(result[offset] | result[offset + 1] << 8));
        }
        WriteUInt16(result, 20, checksum);

        WriteUInt16(result, 22, 1);
        WriteUInt16(result, 24, 9);
        WriteUInt16(result, 26, 0x0300);
        WriteUInt32(result, 28, 12U);
        WriteUInt32(result, 34, 3U);
        WriteUInt32(result, 40, 3U);
        return result;
    }

    private static byte[] BuildBlipRecord(ushort instance, ushort type, byte[] imageData,
        bool twoUids = false, bool includeTag = true) {
        int uidLength = twoUids ? 32 : 16;
        byte[] payload = new byte[uidLength + (includeTag ? 1 : 0) + imageData.Length];
        if (includeTag) {
            payload[uidLength] = 0xFF;
        }
        Buffer.BlockCopy(imageData, 0, payload, uidLength + (includeTag ? 1 : 0), imageData.Length);
        byte[] record = new byte[8 + payload.Length];
        ushort versionAndInstance = unchecked((ushort)(instance << 4));
        record[0] = unchecked((byte)versionAndInstance);
        record[1] = unchecked((byte)(versionAndInstance >> 8));
        record[2] = unchecked((byte)type);
        record[3] = unchecked((byte)(type >> 8));
        WriteUInt32(record, 4, unchecked((uint)payload.Length));
        Buffer.BlockCopy(payload, 0, record, 8, payload.Length);
        return record;
    }

    private static uint ReadUInt32(byte[] source, int offset) => unchecked((uint)(
        source[offset] | source[offset + 1] << 8 | source[offset + 2] << 16
        | source[offset + 3] << 24));

    private static void WriteUInt32(byte[] target, int offset, uint value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
        target[offset + 2] = unchecked((byte)(value >> 16));
        target[offset + 3] = unchecked((byte)(value >> 24));
    }

    private static void WriteUInt16(byte[] target, int offset, ushort value) {
        target[offset] = unchecked((byte)value);
        target[offset + 1] = unchecked((byte)(value >> 8));
    }
}
