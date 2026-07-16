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
