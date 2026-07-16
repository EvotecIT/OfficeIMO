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
