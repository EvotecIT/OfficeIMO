using System;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocImageValidationTests {
    [Fact]
    public void Image_WithNullBytes_ThrowsArgumentNullException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentNullException>(() => doc.Image(null!, 24, 24));

        Assert.Equal("jpegBytes", exception.ParamName);
        Assert.Contains("Parameter 'jpegBytes' cannot be null.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithEmptyBytes_ThrowsArgumentException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentException>(() => doc.Image(Array.Empty<byte>(), 24, 24));

        Assert.Equal("jpegBytes", exception.ParamName);
        Assert.Contains("Parameter 'jpegBytes' cannot be empty.", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void Image_WithInvalidWidth_ThrowsArgumentOutOfRangeException(double invalidWidth) {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => doc.Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, invalidWidth, 10));

        Assert.Equal("width", exception.ParamName);
        Assert.Contains("Parameter 'width' must be a finite positive number.", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void Image_WithInvalidHeight_ThrowsArgumentOutOfRangeException(double invalidHeight) {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => doc.Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, 10, invalidHeight));

        Assert.Equal("height", exception.ParamName);
        Assert.Contains("Parameter 'height' must be a finite positive number.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnImage_WithNullBytes_ThrowsArgumentNullException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentNullException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(null!, 24, 24)))))));

        Assert.Equal("jpegBytes", exception.ParamName);
    }

    [Fact]
    public void RowColumnImage_WithEmptyBytes_ThrowsArgumentException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(Array.Empty<byte>(), 24, 24)))))));

        Assert.Equal("jpegBytes", exception.ParamName);
        Assert.Contains("Parameter 'jpegBytes' cannot be empty.", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    public void RowColumnImage_WithNonPositiveWidth_ThrowsArgumentOutOfRangeException(double invalidWidth) {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, invalidWidth, 24)))))));

        Assert.Equal("width", exception.ParamName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    public void RowColumnImage_WithNonPositiveHeight_ThrowsArgumentOutOfRangeException(double invalidHeight) {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, 24, invalidHeight)))))));

        Assert.Equal("height", exception.ParamName);
    }

    [Fact]
    public void ItemComposeImage_WithNullBytes_ThrowsArgumentNullException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentNullException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Image(null!, 24, 24))))));

        Assert.Equal("jpegBytes", exception.ParamName);
    }

    [Fact]
    public void ItemComposeImage_WithEmptyBytes_ThrowsArgumentException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Image(Array.Empty<byte>(), 24, 24))))));

        Assert.Equal("jpegBytes", exception.ParamName);
        Assert.Contains("Parameter 'jpegBytes' cannot be empty.", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-2)]
    public void ItemComposeImage_WithNonPositiveWidth_ThrowsArgumentOutOfRangeException(double invalidWidth) {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, invalidWidth, 24))))));

        Assert.Equal("width", exception.ParamName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-2)]
    public void ItemComposeImage_WithNonPositiveHeight_ThrowsArgumentOutOfRangeException(double invalidHeight) {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, 24, invalidHeight))))));

        Assert.Equal("height", exception.ParamName);
    }
}
