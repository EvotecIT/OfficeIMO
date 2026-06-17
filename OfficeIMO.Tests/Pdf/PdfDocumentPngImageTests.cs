using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentPngImageTests {
    [Fact]
    public void Image_With16BitRgbPng_Writes8BitRgbImageObject() {
        byte[] bytes = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitRgbPng(), 24, 24)
            .ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 1 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/BitsPerComponent 8", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.DoesNotContain("/SMask", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_With16BitRgbaPng_WritesSoftMaskImageObject() {
        byte[] bytes = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitRgbaPng(), 24, 24)
            .ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 1 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/BitsPerComponent 8", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithExpandedPng_WritesCompressedFlateImageStreams() {
        byte[] bytes = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitRgbaPng(), 24, 24)
            .ToBytes();

        var imageStreams = GetImageStreams(bytes);

        Assert.Equal(2, imageStreams.Count);
        foreach (PdfStream imageStream in imageStreams) {
            Assert.False(imageStream.DecodingFailed, imageStream.DecodingError);
            Assert.True(imageStream.Data.Length >= 2);
            Assert.Equal(0x78, imageStream.Data[0]);
            Assert.Equal(0x9C, imageStream.Data[1]);
            Assert.Equal("FlateDecode", Assert.IsType<PdfName>(imageStream.Dictionary.Items["Filter"]).Name);
        }

        var images = PdfImageExtractor.ExtractImages(bytes);
        Assert.Single(images);
    }

    [Fact]
    public void Image_WithInterlacedRgbPng_WritesNonInterlacedRgbImageObject() {
        byte[] sourcePng = PdfPngTestImages.CreateInterlacedRgbPng();

        Assert.True(PdfDocument.TryValidateImageBytes(sourcePng, out OfficeImageInfo? imageInfo, out string? unsupportedReason), unsupportedReason);
        Assert.NotNull(imageInfo);
        Assert.Equal(OfficeImageFormat.Png, imageInfo.Format);

        byte[] bytes = PdfDocument.Create()
            .Image(sourcePng, 32, 32)
            .ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 8 /Height 8", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.DoesNotContain("/SMask", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);

        var images = PdfImageExtractor.ExtractImages(bytes);
        var image = Assert.Single(images);
        Assert.Equal(8, image.Width);
        Assert.Equal(8, image.Height);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(0, PdfPngTestImages.ReadPngInterlaceMethod(image.Bytes));
        Assert.Equal(PdfPngTestImages.CreateInterlacedRgbExpectedScanlines(), PdfPngTestImages.DecodePngIdat(image.Bytes));
    }

    [Fact]
    public void Image_WithInterlacedRgbaPng_PreservesSoftMaskPixels() {
        byte[] sourcePng = PdfPngTestImages.CreateInterlacedRgbaPng();

        Assert.True(PdfDocument.TryValidateImageBytes(sourcePng, out OfficeImageInfo? imageInfo, out string? unsupportedReason), unsupportedReason);
        Assert.NotNull(imageInfo);
        Assert.Equal(OfficeImageFormat.Png, imageInfo.Format);

        byte[] bytes = PdfDocument.Create()
            .Image(sourcePng, 32, 32)
            .ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 8 /Height 8", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);

        var images = PdfImageExtractor.ExtractImages(bytes);
        var image = Assert.Single(images);
        Assert.Equal(8, image.Width);
        Assert.Equal(8, image.Height);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(0, PdfPngTestImages.ReadPngInterlaceMethod(image.Bytes));
        Assert.Equal(PdfPngTestImages.CreateInterlacedRgbaExpectedScanlines(), PdfPngTestImages.DecodePngIdat(image.Bytes));
    }

    [Fact]
    public void Image_WithInterlacedIndexedPng_PreservesPaletteTransparencyPixels() {
        byte[] sourcePng = PdfPngTestImages.CreateInterlacedIndexedPng();

        Assert.True(PdfDocument.TryValidateImageBytes(sourcePng, out OfficeImageInfo? imageInfo, out string? unsupportedReason), unsupportedReason);
        Assert.NotNull(imageInfo);
        Assert.Equal(OfficeImageFormat.Png, imageInfo.Format);

        byte[] bytes = PdfDocument.Create()
            .Image(sourcePng, 32, 32)
            .ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 8 /Height 8", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);

        var images = PdfImageExtractor.ExtractImages(bytes);
        var image = Assert.Single(images);
        Assert.Equal(8, image.Width);
        Assert.Equal(8, image.Height);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(0, PdfPngTestImages.ReadPngInterlaceMethod(image.Bytes));
        Assert.Equal(PdfPngTestImages.CreateInterlacedIndexedRgbaExpectedScanlines(), PdfPngTestImages.DecodePngIdat(image.Bytes));
    }

    [Fact]
    public void Image_WithOversizedInterlacedPng_RejectsImageBytesWithoutExpanding() {
        Assert.False(PdfDocument.TryValidateImageBytes(
            PdfPngTestImages.CreateOversizedInterlacedGrayscalePng(),
            out OfficeImageInfo? imageInfo,
            out string? unsupportedReason));
        Assert.Null(imageInfo);
        Assert.Contains("PNG dimensions exceed", unsupportedReason, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithInvalidPngCrc_RejectsImageBytes() {
        Assert.False(PdfDocument.TryValidateImageBytes(
            PdfPngTestImages.CreatePngWithInvalidCrc(),
            out OfficeImageInfo? imageInfo,
            out string? unsupportedReason));
        Assert.Null(imageInfo);
        Assert.Contains("PNG chunk CRC is invalid.", unsupportedReason, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithExtraDecodedPngScanlines_RejectsPassThroughImageBytes() {
        Assert.False(PdfDocument.TryValidateImageBytes(
            PdfPngTestImages.CreateRgbPngWithExtraDecodedScanlines(),
            out OfficeImageInfo? imageInfo,
            out string? unsupportedReason));
        Assert.Null(imageInfo);
        Assert.Contains("PNG image data length does not match the expected scanline size.", unsupportedReason, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithOverflowingPngChunkLength_RejectsPdfGenerationWithoutThrowingParserErrors() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create()
                .Image(PdfPngTestImages.CreatePngWithOverflowingChunkLength(), 24, 24)
                .ToBytes());

        Assert.Contains("PNG chunk length is invalid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_With16BitRgbTransparency_WritesSoftMaskImageObject() {
        byte[] bytes = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitRgbPng(includeTransparency: true), 24, 24)
            .ToBytes();

        string pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/BitsPerComponent 8", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithInvalidRgbaTransparencyChunk_RejectsImageBytes() {
        Assert.False(PdfDocument.TryValidateImageBytes(
            PdfPngTestImages.CreateRgbaPngWithInvalidTransparencyChunk(),
            out OfficeImageInfo? imageInfo,
            out string? unsupportedReason));
        Assert.Null(imageInfo);
        Assert.Contains("PNG transparency chunks are not valid for grayscale-alpha or RGBA PNG images.", unsupportedReason, StringComparison.Ordinal);
    }

    [Fact]
    public void StampImage_With16BitRgbaPng_PreservesSoftMaskImageObject() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Stamp source"))
            .ToBytes();

        byte[] stamped = PdfStamper.StampImage(source, PdfPngTestImages.Create16BitRgbaPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhen16BitRgbaCreatesSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitRgbaPng(), 24, 24)
            .Paragraph(p => p.Text("16-bit RGBA image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 18, 128, 255, 64 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsGrayAlphaPngWhen16BitGrayscaleAlphaCreatesSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitGrayscaleAlphaPng(), 24, 24)
            .Paragraph(p => p.Text("16-bit grayscale-alpha image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        Assert.Equal(4, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 128, 64 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhen16BitRgbTransparencyCreatesSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.Create16BitRgbPng(includeTransparency: true), 24, 24)
            .Paragraph(p => p.Text("16-bit RGB transparency image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 18, 128, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    private static List<PdfStream> GetImageStreams(byte[] bytes) {
        var (objects, _) = PdfSyntax.ParseObjects(bytes);
        var imageStreams = new List<PdfStream>();
        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is PdfStream stream &&
                stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) &&
                subtype is PdfName name &&
                string.Equals(name.Name, "Image", StringComparison.Ordinal)) {
                imageStreams.Add(stream);
            }
        }

        return imageStreams;
    }
}
