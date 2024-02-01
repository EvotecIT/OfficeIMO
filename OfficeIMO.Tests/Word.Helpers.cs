using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("snail.bmp", "image/bmp")]
    [InlineData("example.gif", "image/gif")]
    [InlineData("Kulek.jpg", "image/jpeg")]
    [InlineData("BackgroundImage.png", "image/png")]
    [InlineData("saturn.tif", "image/tiff")]
    public void Test_GetImageСharacteristics(string filename, string expectedType) {
        var filePath = Path.Combine(_directoryWithImages, filename);
        using var imageStream = new FileStream(filePath, FileMode.Open);
        var imageСharacteristics = Helpers.GetImageСharacteristics(imageStream);
        Assert.Equal(expectedType, imageСharacteristics.Type);
    }
}
