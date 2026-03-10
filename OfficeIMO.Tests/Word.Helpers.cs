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
}
