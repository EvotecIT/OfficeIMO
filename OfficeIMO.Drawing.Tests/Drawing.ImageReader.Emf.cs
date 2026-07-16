using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageReaderRejectsEmfWithoutAnEofRecord() {
        byte[] emf = CreateCompleteEmf(2, 2);
        Array.Resize(ref emf, 88);
        WriteInt32LittleEndian(emf, 48, 88);

        Assert.False(OfficeImageReader.TryIdentifyByContent(emf, "header-only.emf", out _));
    }

    [Fact]
    public void OfficeImageReaderRejectsEmfWithAnInaccurateRecordCount() {
        byte[] emf = CreateCompleteEmf(2, 2);
        WriteInt32LittleEndian(emf, 52, 3);

        Assert.False(OfficeImageReader.TryIdentifyByContent(emf, "bad-count.emf", out _));
    }

    private static byte[] CreateCompleteEmf(int width, int height) {
        var emf = new byte[108];
        WriteInt32LittleEndian(emf, 0, 1);
        WriteInt32LittleEndian(emf, 4, 88);
        WriteInt32LittleEndian(emf, 16, width);
        WriteInt32LittleEndian(emf, 20, height);
        WriteInt32LittleEndian(emf, 40, 0x464D4520);
        WriteInt32LittleEndian(emf, 44, 0x00010000);
        WriteInt32LittleEndian(emf, 48, emf.Length);
        WriteInt32LittleEndian(emf, 52, 2);
        WriteUInt16LittleEndian(emf, 56, 1);
        WriteInt32LittleEndian(emf, 88, 14);
        WriteInt32LittleEndian(emf, 92, 20);
        WriteInt32LittleEndian(emf, 104, 20);
        return emf;
    }
}
