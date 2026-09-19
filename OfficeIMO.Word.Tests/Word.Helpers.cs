using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Provides helper methods for Word tests.
/// </summary>
public partial class Word {
    [Theory]
    [InlineData("snail.bmp", OfficeImageFormat.Bmp)]
    [InlineData("example.gif", OfficeImageFormat.Gif)]
    [InlineData("Kulek.jpg", OfficeImageFormat.Jpeg)]
    [InlineData("BackgroundImage.png", OfficeImageFormat.Png)]
    [InlineData("saturn.tif", OfficeImageFormat.Tiff)]
    public void Test_GetImageСharacteristics(string filename, OfficeImageFormat expectedType) {
        var filePath = Path.Combine(_directoryWithImages, filename);
        using var imageStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var imageСharacteristics = Helpers.GetImageCharacteristics(imageStream, filename);
        Assert.Equal(expectedType, imageСharacteristics.Type);
    }

    [Fact]
    public void Test_GetImageCharacteristics_ForCompleteEmf() {
        using var imageStream = new MemoryStream(CreateCompleteEmf());

        var imageCharacteristics = Helpers.GetImageCharacteristics(imageStream, "sample.emf");

        Assert.Equal(OfficeImageFormat.Emf, imageCharacteristics.Type);
        Assert.Equal(2, imageCharacteristics.Width);
        Assert.Equal(2, imageCharacteristics.Height);
    }

    [Fact]
    public void Test_GetImageCharacteristics_ForCompletePlaceableWmf() {
        using var imageStream = new MemoryStream(CreatePlaceableWmf());

        var imageCharacteristics = Helpers.GetImageCharacteristics(imageStream, "sample.wmf");

        Assert.Equal(OfficeImageFormat.Wmf, imageCharacteristics.Type);
        Assert.Equal(192, imageCharacteristics.Width);
        Assert.Equal(96, imageCharacteristics.Height);
    }

    [Fact]
    public void Test_GetImageCharacteristics_RejectsUnsupportedWebpWordImagePart() {
        using var imageStream = new MemoryStream(new byte[] { 1, 2, 3, 4 });

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => Helpers.GetImageCharacteristics(imageStream, "preview.webp"));

        Assert.Contains("Webp", exception.Message);
    }

    [Fact]
    public void Test_GetImageCharacteristics_FromNonSeekableStream() {
        var filePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
        using var imageStream = new NonSeekableReadStream(File.ReadAllBytes(filePath));

        var imageCharacteristics = Helpers.GetImageCharacteristics(imageStream, "Kulek.jpg");

        Assert.Equal(OfficeImageFormat.Jpeg, imageCharacteristics.Type);
        Assert.True(imageCharacteristics.Width > 0);
        Assert.True(imageCharacteristics.Height > 0);
    }

    [Fact]
    public void Test_GetImageCharacteristics_RejectsIncompleteJpeg2000Container() {
        byte[] payload = CreateIncompleteJpeg2000Container();
        Assert.True(OfficeImageReader.TryIdentify(payload, "truncated.jp2", out OfficeImageInfo identified));
        Assert.Equal(OfficeImageFormat.Jpeg2000, identified.Format);
        using var imageStream = new MemoryStream(payload);

        Assert.Throws<InvalidDataException>(() =>
            Helpers.GetImageCharacteristics(imageStream, "truncated.jp2"));

        using WordDocument document = WordDocument.Create();
        using var publicApiStream = new MemoryStream(payload);
        Assert.Throws<InvalidDataException>(() =>
            document.AddParagraph().AddImage(publicApiStream, "truncated.jp2", width: null, height: null));
    }

    private static byte[] CreateIncompleteJpeg2000Container() {
        byte[] signature = { 0, 0, 0, 12, (byte)'j', (byte)'P', (byte)' ', (byte)' ', 13, 10, 135, 10 };
        byte[] fileType = CreateJpeg2000Box("ftyp", new byte[] {
            (byte)'j', (byte)'p', (byte)'2', (byte)' ', 0, 0, 0, 0,
            (byte)'j', (byte)'p', (byte)'2', (byte)' '
        });
        byte[] imageHeader = CreateJpeg2000Box("ihdr", new byte[] {
            0, 0, 0, 1, 0, 0, 0, 1, 0, 3, 7, 7, 0, 0
        });
        byte[] colorSpecification = CreateJpeg2000Box("colr", new byte[] {
            1, 0, 0, 0, 0, 0, 16
        });
        byte[] header = CreateJpeg2000Box("jp2h", imageHeader.Concat(colorSpecification).ToArray());
        byte[] truncatedCodestream = {
            0xFF, 0x4F, 0xFF, 0x51, 0x00, 0x2F, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x03, 0x07, 0x01, 0x01, 0x07, 0x01, 0x01,
            0x07, 0x01, 0x01
        };
        byte[] codestream = CreateJpeg2000Box("jp2c", truncatedCodestream);
        return signature.Concat(fileType).Concat(header).Concat(codestream).ToArray();
    }

    private static byte[] CreateJpeg2000Box(string type, byte[] contents) {
        byte[] box = new byte[8 + contents.Length];
        WriteJpeg2000UInt32BigEndian(box, 0, box.Length);
        System.Text.Encoding.ASCII.GetBytes(type, 0, 4, box, 4);
        Buffer.BlockCopy(contents, 0, box, 8, contents.Length);
        return box;
    }

    private static void WriteJpeg2000UInt32BigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static byte[] CreatePlaceableWmf() {
        var wmf = new byte[56];
        WriteInt32LittleEndian(wmf, 0, unchecked((int)0x9AC6CDD7));
        WriteInt16LittleEndian(wmf, 10, 2880);
        WriteInt16LittleEndian(wmf, 12, 1440);
        WriteUInt16LittleEndian(wmf, 14, 1440);
        WritePlaceableWmfChecksum(wmf);
        WriteUInt16LittleEndian(wmf, 22, 1);
        WriteUInt16LittleEndian(wmf, 24, 9);
        WriteUInt16LittleEndian(wmf, 26, 0x0300);
        WriteInt32LittleEndian(wmf, 28, 17);
        WriteInt32LittleEndian(wmf, 34, 5);
        WriteInt32LittleEndian(wmf, 40, 5);
        WriteUInt16LittleEndian(wmf, 44, 0x0201);
        WriteInt32LittleEndian(wmf, 50, 3);
        return wmf;
    }

    private static byte[] CreateCompleteEmf() {
        var emf = new byte[108];
        WriteInt32LittleEndian(emf, 0, 1);
        WriteInt32LittleEndian(emf, 4, 88);
        WriteInt32LittleEndian(emf, 16, 2);
        WriteInt32LittleEndian(emf, 20, 2);
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

    private static void WriteInt16LittleEndian(byte[] data, int offset, short value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteUInt16LittleEndian(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteInt32LittleEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
        data[offset + 2] = (byte)((value >> 16) & 0xFF);
        data[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private static void WritePlaceableWmfChecksum(byte[] data) {
        ushort checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        WriteUInt16LittleEndian(data, 20, checksum);
    }
}
