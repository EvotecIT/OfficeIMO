using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgTransformedImageTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageOutsideLocalViewportCanBeTransformedIntoView(bool pattern) {
        string image = "<image x='-10' y='0' width='4' height='4' href='data:image/png;base64," + Convert.ToBase64String(OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.Red))) + "'/>";
        string content = "<g transform='translate(60 0) scale(5)'>" + image + "</g>";
        if (pattern) content = "<defs><pattern id='p' patternUnits='userSpaceOnUse' width='40' height='40'>" + content + "</pattern></defs><path d='M0,0H80V40H0Z' fill='url(#p)'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='40'>" + content + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.True(OfficeRasterImageDecoder.TryDecode(drawing!.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(15, 10));
        if (pattern) Assert.Equal(OfficeColor.Red, raster.GetPixel(55, 10));
    }
}
