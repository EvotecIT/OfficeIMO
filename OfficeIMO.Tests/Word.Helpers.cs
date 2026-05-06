using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Provides helper methods for Word tests.
/// </summary>
public partial class Word {
    [Theory]
    [InlineData("snail.bmp", CustomImagePartType.Bmp)]
    [InlineData("example.gif", CustomImagePartType.Gif)]
    [InlineData("Kulek.jpg", CustomImagePartType.Jpeg)]
    [InlineData("BackgroundImage.png", CustomImagePartType.Png)]
    [InlineData("saturn.tif", CustomImagePartType.Tiff)]
    [InlineData("sample.emf", CustomImagePartType.Emf)]
    public void Test_GetImageСharacteristics(string filename, CustomImagePartType expectedType) {
        var filePath = Path.Combine(_directoryWithImages, filename);
        using var imageStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var imageСharacteristics = Helpers.GetImageCharacteristics(imageStream, filename);
        Assert.Equal(expectedType, imageСharacteristics.Type);
    }

    [Fact]
    public void Test_GetImageCharacteristics_ForPlaceableWmfHeader() {
        using var imageStream = new MemoryStream(CreatePlaceableWmfHeader());

        var imageCharacteristics = Helpers.GetImageCharacteristics(imageStream, "sample.wmf");

        Assert.Equal(CustomImagePartType.Wmf, imageCharacteristics.Type);
        Assert.Equal(192, imageCharacteristics.Width);
        Assert.Equal(96, imageCharacteristics.Height);
    }

    [Theory]
    [InlineData(CustomImagePartType.Emf, "image/x-emf")]
    [InlineData(CustomImagePartType.Wmf, "image/x-wmf")]
    public void Test_CustomImagePartType_ToOpenXmlImagePartType(CustomImagePartType imagePartType, string expectedContentType) {
        Assert.Equal(expectedContentType, imagePartType.ToOpenXmlImagePartType());
    }

    [Fact]
    public void Test_GetImageCharacteristics_FromNonSeekableStream() {
        var filePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
        using var imageStream = new NonSeekableReadStream(File.ReadAllBytes(filePath));

        var imageCharacteristics = Helpers.GetImageCharacteristics(imageStream, "Kulek.jpg");

        Assert.Equal(CustomImagePartType.Jpeg, imageCharacteristics.Type);
        Assert.True(imageCharacteristics.Width > 0);
        Assert.True(imageCharacteristics.Height > 0);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Open_WithInvalidFilePath_ThrowsArgumentException(string? path) {
        Assert.Throws<ArgumentException>(() => Helpers.Open(path!, true));
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
