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
    [InlineData("sample.emf", OfficeImageFormat.Emf)]
    public void Test_GetImageСharacteristics(string filename, OfficeImageFormat expectedType) {
        var filePath = Path.Combine(_directoryWithImages, filename);
        using var imageStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var imageСharacteristics = Helpers.GetImageCharacteristics(imageStream, filename);
        Assert.Equal(expectedType, imageСharacteristics.Type);
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
