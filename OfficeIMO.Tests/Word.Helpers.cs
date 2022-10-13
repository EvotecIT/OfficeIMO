using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("snail.bmp", ImagePartType.Bmp)]
    [InlineData("example.gif", ImagePartType.Gif)]
    [InlineData("Kulek.jpg", ImagePartType.Jpeg)]
    [InlineData("BackgroundImage.png", ImagePartType.Png)]
    [InlineData("saturn.tif", ImagePartType.Tiff)]
    public void Test_GetImageСharacteristics(string filename, ImagePartType expectedType) {
        var filePath = Path.Combine(_directoryWithImages, filename);
        using var imageStream = new FileStream(filePath, FileMode.Open);
        var imageСharacteristics = Helpers.GetImageСharacteristics(imageStream);
        Assert.Equal(expectedType, imageСharacteristics.Type);
    }
}
