using System;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderMediaAdapterTests {
    [Fact]
    public void ImageAdapter_IdentifiesStandardWmfWithoutPlaceableHeader() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(CreateStandardWmf(), "diagram.wmf");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/x-wmf", asset.MediaType);
        Assert.Null(asset.Width);
        Assert.Null(asset.Height);
        Assert.Contains("Format: Wmf", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(OfficeDocumentReaderBuilderImageExtensions.HandlerId, result.CapabilitiesUsed);
    }

    [Fact]
    public void ImageAdapter_IdentifiesPlaceableWmfWithStandardPayload() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(CreatePlaceableWmf(), "diagram.wmf");

        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image/x-wmf", asset.MediaType);
        Assert.Equal(192, asset.Width);
        Assert.Equal(96, asset.Height);
    }

    [Fact]
    public void ImageAdapter_RejectsPlaceableWmfWithoutStandardPayload() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(CreatePlaceableWmfHeader(), "truncated.wmf"));
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
        WriteUInt32LittleEndian(wmf, 0, unchecked((int)0x9AC6CDD7));
        WriteUInt16LittleEndian(wmf, 10, 2880);
        WriteUInt16LittleEndian(wmf, 12, 1440);
        WriteUInt16LittleEndian(wmf, 14, 1440);
        WritePlaceableWmfChecksum(wmf);
        return wmf;
    }

    private static void WritePlaceableWmfChecksum(byte[] data) {
        int checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= data[offset] | (data[offset + 1] << 8);
        }

        WriteUInt16LittleEndian(data, 20, checksum);
    }

    private static byte[] CreateStandardWmf() {
        var wmf = new byte[34];
        WriteUInt16LittleEndian(wmf, 0, 1);
        WriteUInt16LittleEndian(wmf, 2, 9);
        WriteUInt16LittleEndian(wmf, 4, 0x0300);
        WriteUInt32LittleEndian(wmf, 6, 17);
        WriteUInt32LittleEndian(wmf, 12, 5);
        WriteUInt32LittleEndian(wmf, 18, 5);
        WriteUInt16LittleEndian(wmf, 22, 0x0201);
        WriteUInt32LittleEndian(wmf, 28, 3);
        return wmf;
    }
}
