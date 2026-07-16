using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageReaderIdentifiesStandardWmfWithoutPlaceableHeader() {
        byte[] wmf = CreateStandardWmf();

        Assert.True(OfficeImageReader.TryIdentifyByContent(wmf, "diagram.wmf", out OfficeImageInfo image));

        Assert.Equal(OfficeImageFormat.Wmf, image.Format);
        Assert.Equal(0, image.Width);
        Assert.Equal(0, image.Height);
        Assert.Equal("image/x-wmf", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderRejectsStandardWmfWithRecordPastDeclaredPayload() {
        byte[] wmf = CreateStandardWmf();
        WriteInt32LittleEndian(wmf, 12, 10);
        WriteInt32LittleEndian(wmf, 18, 10);

        Assert.False(OfficeImageReader.TryIdentifyByContent(wmf, "truncated.wmf", out _));
    }

    [Fact]
    public void OfficeImageReaderRejectsStandardWmfWithInaccurateMaximumRecordSize() {
        byte[] wmf = CreateStandardWmf();
        WriteInt32LittleEndian(wmf, 12, 6);

        Assert.False(OfficeImageReader.TryIdentifyByContent(wmf, "invalid-maximum.wmf", out _));
    }

    [Fact]
    public void OfficeImageReaderRejectsPlaceableWmfWithoutStandardPayload() {
        byte[] wmf = CreatePlaceableWmfHeader();

        Assert.False(OfficeImageReader.TryIdentifyByContent(wmf, "truncated-placeable.wmf", out _));
    }

    private static byte[] CreatePlaceableWmf() {
        byte[] header = CreatePlaceableWmfHeader();
        byte[] payload = CreateStandardWmf();
        var wmf = new byte[header.Length + payload.Length];
        Array.Copy(header, wmf, header.Length);
        Array.Copy(payload, 0, wmf, header.Length, payload.Length);
        return wmf;
    }

    private static byte[] CreatePlaceableWmfHeader() {
        var wmf = new byte[22];
        WriteInt32LittleEndian(wmf, 0, unchecked((int)0x9AC6CDD7));
        WriteInt16LittleEndian(wmf, 10, 2880);
        WriteInt16LittleEndian(wmf, 12, 1440);
        WriteUInt16LittleEndian(wmf, 14, 1440);
        WritePlaceableWmfChecksum(wmf);
        return wmf;
    }

    private static byte[] CreateStandardWmf() {
        var wmf = new byte[34];
        WriteUInt16LittleEndian(wmf, 0, 1);
        WriteUInt16LittleEndian(wmf, 2, 9);
        WriteUInt16LittleEndian(wmf, 4, 0x0300);
        WriteInt32LittleEndian(wmf, 6, 17);
        WriteInt32LittleEndian(wmf, 12, 5);
        WriteInt32LittleEndian(wmf, 18, 5);
        WriteUInt16LittleEndian(wmf, 22, 0x0201);
        WriteInt32LittleEndian(wmf, 28, 3);
        return wmf;
    }
}
