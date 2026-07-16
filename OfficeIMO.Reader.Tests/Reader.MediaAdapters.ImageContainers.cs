using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderMediaAdapterTests {
    [Fact]
    public void ImageAdapter_RejectsBmpCompressionThatDoesNotMatchBitDepth() {
        var bmp = new byte[54];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteUInt32LittleEndian(bmp, 14, 40);
        WriteUInt32LittleEndian(bmp, 18, 2);
        WriteUInt32LittleEndian(bmp, 22, 2);
        WriteUInt16LittleEndian(bmp, 26, 1);
        WriteUInt16LittleEndian(bmp, 28, 24);
        WriteUInt32LittleEndian(bmp, 30, 1);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(bmp, "invalid-compression.bmp"));
    }

    [Fact]
    public void ImageAdapter_RejectsPcxWithoutEncodedImageData() {
        var pcx = new byte[128];
        pcx[0] = 0x0A;
        pcx[1] = 0x05;
        pcx[2] = 0x01;
        pcx[3] = 0x08;
        pcx[8] = 1;
        pcx[10] = 1;
        pcx[65] = 1;
        WriteUInt16LittleEndian(pcx, 66, 2);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(pcx, "header-only.pcx"));
    }

    [Fact]
    public void ImageAdapter_RejectsEmfWithoutAnEofRecord() {
        byte[] emf = CreateCompleteEmf(2, 2);
        Array.Resize(ref emf, 88);
        WriteUInt32LittleEndian(emf, 48, 88);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();

        Assert.Throws<NotSupportedException>(() =>
            reader.ReadDocument(emf, "header-only.emf"));
    }

    private static byte[] CreateCompleteEmf(int width, int height) {
        var emf = new byte[108];
        WriteUInt32LittleEndian(emf, 0, 1);
        WriteUInt32LittleEndian(emf, 4, 88);
        WriteUInt32LittleEndian(emf, 16, width);
        WriteUInt32LittleEndian(emf, 20, height);
        WriteUInt32LittleEndian(emf, 40, 0x464D4520);
        WriteUInt32LittleEndian(emf, 44, 0x00010000);
        WriteUInt32LittleEndian(emf, 48, emf.Length);
        WriteUInt32LittleEndian(emf, 52, 2);
        WriteUInt16LittleEndian(emf, 56, 1);
        WriteUInt32LittleEndian(emf, 88, 14);
        WriteUInt32LittleEndian(emf, 92, 20);
        WriteUInt32LittleEndian(emf, 104, 20);
        return emf;
    }
}
