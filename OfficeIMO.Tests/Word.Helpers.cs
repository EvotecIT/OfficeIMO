using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("snail.bmp", CustomImagePartType.Bmp)]
    [InlineData("example.gif", CustomImagePartType.Gif)]
    [InlineData("Kulek.jpg", CustomImagePartType.Jpeg)]
    [InlineData("BackgroundImage.png", CustomImagePartType.Png)]
    [InlineData("saturn.tif", CustomImagePartType.Tiff)]
    public void Test_GetImageСharacteristics(string filename, CustomImagePartType expectedType) {
        var filePath = Path.Combine(_directoryWithImages, filename);
        using var imageStream = new FileStream(filePath, FileMode.Open);
        var imageСharacteristics = Helpers.GetImageCharacteristics(imageStream);
        Assert.Equal(expectedType, imageСharacteristics.Type);
    }
}
