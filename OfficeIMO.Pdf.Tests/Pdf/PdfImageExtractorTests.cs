using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImageExtractorTests {
    [Fact]
    public void ExtractImages_ReturnsPngFilesByPageFromGeneratedPdf() {
        byte[] source = PdfDocument.Create()
            .Image(CreateMinimalRgbPng(), 24, 24)
            .Paragraph(p => p.Text("Image page marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal(8, image.BitsPerComponent);
        Assert.Equal("DeviceRGB", image.ColorSpace);
        Assert.Equal("FlateDecode", image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenImageHasSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(CreateMinimalRgbaPng(), 24, 24)
            .Paragraph(p => p.Text("Transparent image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsJpegWithUnresolvedTransparencyMetadataWhenDctImageHasSoftMask() {
        byte[] source = BuildDeviceRgbJpegSoftMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("DeviceRGB", image.ColorSpace);
        Assert.Equal("DCTDecode", image.Filter);
        Assert.Equal("jpg", image.FileExtension);
        Assert.Equal("image/jpeg", image.MimeType);
        Assert.True(image.IsImageFile);
        Assert.True(image.HasTransparencyMask);
        Assert.True(image.HasUnresolvedTransparencyMask);
        Assert.False(image.TransparencyMaskResolved);
        Assert.Equal("soft-mask", image.TransparencyMaskKind);
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Jpeg, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesUnfilteredDeviceRgbImageStreamsToPngFiles() {
        byte[] source = BuildUnfilteredDeviceRgbImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal(8, image.BitsPerComponent);
        Assert.Equal("DeviceRGB", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(new byte[] { 0, 97, 98, 99 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesSupportedFilterChainDeviceRgbImageStreamsToPngFiles() {
        byte[] source = BuildAsciiHexFlateDeviceRgbImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("ASCIIHexDecode,FlateDecode", image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(new byte[] { 0, 97, 98, 99 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenDeviceRgbImageHasColorKeyMask() {
        byte[] source = BuildDeviceRgbColorKeyMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("DeviceRGB", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 97, 98, 99, 0, 100, 101, 102, 255 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReportsExplicitMaskImageAsUnresolvedWhenRgbStreamNormalizesToPng() {
        byte[] source = BuildDeviceRgbExplicitMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("png", image.FileExtension);
        Assert.True(image.IsImageFile);
        Assert.True(image.HasTransparencyMask);
        Assert.True(image.HasUnresolvedTransparencyMask);
        Assert.False(image.TransparencyMaskResolved);
        Assert.Equal("explicit-mask-image", image.TransparencyMaskKind);
        AssertPngSignature(image.Bytes);
    }

    [Fact]
    public void ExtractImages_AppliesImageDecodeArrayWhenNormalizingDeviceGrayStreamsToPngFiles() {
        byte[] source = BuildUnfilteredDeviceGrayInvertedDecodeImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("DeviceGray", image.ColorSpace);
        Assert.Equal(8, image.BitsPerComponent);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(0, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 158 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesUnfilteredDeviceCmykImageStreamsToRgbPngFiles() {
        byte[] source = BuildUnfilteredDeviceCmykImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal(8, image.BitsPerComponent);
        Assert.Equal("DeviceCMYK", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 126, 125, 124 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesIccBasedRgbImageStreamsToPngFiles() {
        byte[] source = BuildUnfilteredIccBasedRgbImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("ICCBased", image.ColorSpace);
        Assert.Equal(8, image.BitsPerComponent);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 97, 98, 99 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesIccBasedCmykImageStreamsToRgbPngFiles() {
        byte[] source = BuildUnfilteredIccBasedCmykImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("ICCBased", image.ColorSpace);
        Assert.Equal(8, image.BitsPerComponent);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 126, 125, 124 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesSupportedFilterChainDeviceCmykImageStreamsToRgbPngFiles() {
        byte[] source = BuildAsciiHexFlateDeviceCmykImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("DeviceCMYK", image.ColorSpace);
        Assert.Equal("ASCIIHexDecode,FlateDecode", image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 126, 125, 124 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenDeviceCmykImageHasSoftMask() {
        byte[] source = BuildDeviceCmykSoftMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("DeviceCMYK", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 126, 125, 124, 126 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenIccBasedCmykImageHasSoftMask() {
        byte[] source = BuildIccBasedCmykSoftMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("ICCBased", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 126, 125, 124, 126 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesPackedIndexedRgbImageStreamsToRgbPngFiles() {
        byte[] source = BuildUnfilteredIndexedRgbImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(2, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal(1, image.BitsPerComponent);
        Assert.Equal("Indexed", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0, 255, 0 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_NormalizesSupportedFilterChainIndexedRgbImageStreamsToRgbPngFiles() {
        byte[] source = BuildAsciiHexFlateIndexedRgbImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("Indexed", image.ColorSpace);
        Assert.Equal("ASCIIHexDecode,FlateDecode", image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0, 255, 0 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenIndexedImageHasSoftMask() {
        byte[] source = BuildIndexedSoftMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("Indexed", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 126, 0, 255, 0, 64 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenIndexedImageHasColorKeyMask() {
        byte[] source = BuildIndexedColorKeyMaskImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("Indexed", image.ColorSpace);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0, 0, 255, 0, 255 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsGrayAlphaPngWhenImageMaskStreamIsSupported() {
        byte[] source = BuildImageMaskPdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("ImageMask", image.ColorSpace);
        Assert.Equal(1, image.BitsPerComponent);
        Assert.Equal(string.Empty, image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(4, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 0, 0, 255 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_RejectsImageMaskDimensionsThatOverflowDecodedBuffers() {
        byte[] source = BuildImagePdfWithColorSpace(
            "/DeviceGray",
            int.MaxValue,
            int.MaxValue,
            1,
            "@",
            " /ImageMask true");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(source));

        Assert.False(image.IsImageFile);
        Assert.NotEqual("png", image.FileExtension);
    }

    [Fact]
    public void ExtractImages_AppliesImageDecodeArrayWhenNormalizingIndexedStreamsToRgbPngFiles() {
        byte[] source = BuildUnfilteredIndexedRgbInvertedDecodeImagePdf();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("Indexed", image.ColorSpace);
        Assert.Equal(1, image.BitsPerComponent);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(2, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 0, 255, 0, 0 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsRgbaPngWhenRgbTransparencyCreatesSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(CreateMinimalRgbTransparencyPng(), 24, 24)
            .Paragraph(p => p.Text("RGB tRNS image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(6, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsGrayAlphaPngWhenGrayscaleTransparencyCreatesSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(CreateMinimalGrayscaleTransparencyPng(), 24, 24)
            .Paragraph(p => p.Text("Grayscale tRNS image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(4, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 128, 0 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsGrayAlphaPngWhenPackedGrayscaleTransparencyCreatesSoftMask() {
        byte[] source = PdfDocument.Create()
            .Image(CreateMinimalPackedGrayscaleTransparencyPng(), 24, 12)
            .Paragraph(p => p.Text("Packed grayscale tRNS image marker"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        var image = Assert.Single(images);
        Assert.Equal(2, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("png", image.FileExtension);
        Assert.Equal("image/png", image.MimeType);
        Assert.True(image.IsImageFile);
        AssertPngSignature(image.Bytes);
        Assert.Equal(4, ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 17, 0 }, DecodeStoredPngIdat(image.Bytes));
        Assert.True(OfficeImageReader.TryIdentify(image.Bytes, null, out var info));
        Assert.Equal(OfficeImageFormat.Png, info.Format);
    }

    [Fact]
    public void ExtractImages_ReturnsStampedImageOnSelectedPage() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 2 },
            Width = 24,
            Height = 24
        });

        var images = PdfImageExtractor.ExtractImages(stamped);

        var image = Assert.Single(images);
        Assert.Equal(2, image.PageNumber);
        Assert.Equal("png", image.FileExtension);
        AssertPngSignature(image.Bytes);
    }

    [Fact]
    public void ExtractImagesByPageRanges_ReturnsSelectedPagesInRangeOrder() {
        byte[] stamped = PdfStamper.StampImage(BuildThreePagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1, 3 },
            Width = 24,
            Height = 24
        });

        var images = PdfImageExtractor.ExtractImagesByPageRanges(stamped, PdfPageRange.ParseMany("3,1-2,3"));

        Assert.Equal(3, images.Count);
        Assert.Equal(3, images[0].PageNumber);
        Assert.Equal(1, images[1].PageNumber);
        Assert.Equal(3, images[2].PageNumber);
        AssertPngSignature(images[0].Bytes);
        AssertPngSignature(images[1].Bytes);
        AssertPngSignature(images[2].Bytes);
    }

    [Fact]
    public void ExtractImages_ReturnsEmptyListWhenPdfHasNoImages() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("No images here"))
            .ToBytes();

        var images = PdfImageExtractor.ExtractImages(source);

        Assert.Empty(images);
    }

    [Fact]
    public void ExtractImages_WritesPathInput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-images-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24).ToBytes());

            var images = PdfImageExtractor.ExtractImages(inputPath);

            var image = Assert.Single(images);
            Assert.Equal("png", image.FileExtension);
            Assert.True(image.IsImageFile);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractImages_ReadsFromCurrentStreamPosition() {
        byte[] pdf = PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24).ToBytes();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        var images = PdfImageExtractor.ExtractImages(stream);

        var image = Assert.Single(images);
        Assert.Equal("png", image.FileExtension);
        AssertPngSignature(image.Bytes);
    }

    [Fact]
    public void ExtractImages_WritesImageFilesToDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-image-files-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "images");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24).ToBytes());

            IReadOnlyList<string> paths = PdfImageExtractor.ExtractImages(inputPath, outputDirectory);

            string expectedPath = Path.Combine(outputDirectory, "source-page-0001-image-0001.png");
            string path = Assert.Single(paths);
            Assert.Equal(expectedPath, path);
            Assert.True(File.Exists(path));
            AssertPngSignature(File.ReadAllBytes(path));

            using var stream = new MemoryStream(PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24).ToBytes());
            string streamOutputDirectory = Path.Combine(directory, "stream-images");
            IReadOnlyList<string> streamPaths = PdfImageExtractor.ExtractImages(stream, streamOutputDirectory, "stream-source.pdf");

            string expectedStreamPath = Path.Combine(streamOutputDirectory, "stream-source-page-0001-image-0001.png");
            string streamPath = Assert.Single(streamPaths);
            Assert.Equal(expectedStreamPath, streamPath);
            Assert.True(File.Exists(streamPath));
            AssertPngSignature(File.ReadAllBytes(streamPath));

            string byteOutputDirectory = Path.Combine(directory, "byte-images");
            IReadOnlyList<string> bytePaths = PdfImageExtractor.ExtractImages(
                PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24).ToBytes(),
                byteOutputDirectory,
                "byte-source.pdf");

            string expectedBytePath = Path.Combine(byteOutputDirectory, "byte-source-page-0001-image-0001.png");
            string bytePath = Assert.Single(bytePaths);
            Assert.Equal(expectedBytePath, bytePath);
            Assert.True(File.Exists(bytePath));
            AssertPngSignature(File.ReadAllBytes(bytePath));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractImagesByPageRanges_WritesSelectedSourcePageImageFiles() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-image-range-files-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "images");

        try {
            Directory.CreateDirectory(directory);
            byte[] stamped = PdfStamper.StampImage(BuildThreePagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
                PageNumbers = new[] { 1, 3 },
                Width = 24,
                Height = 24
            });
            File.WriteAllBytes(inputPath, stamped);

            IReadOnlyList<string> paths = PdfImageExtractor.ExtractImagesByPageRanges(inputPath, outputDirectory, PdfPageRange.ParseMany("3"));

            string expectedPath = Path.Combine(outputDirectory, "source-page-0003-image-0001.png");
            string path = Assert.Single(paths);
            Assert.Equal(expectedPath, path);
            Assert.True(File.Exists(path));
            AssertPngSignature(File.ReadAllBytes(path));

            using var stream = new MemoryStream(stamped);
            string streamOutputDirectory = Path.Combine(directory, "stream-images");
            IReadOnlyList<string> streamPaths = PdfImageExtractor.ExtractImagesByPageRanges(stream, streamOutputDirectory, "stream-source.pdf", PdfPageRange.ParseMany("1"));

            string expectedStreamPath = Path.Combine(streamOutputDirectory, "stream-source-page-0001-image-0001.png");
            string streamPath = Assert.Single(streamPaths);
            Assert.Equal(expectedStreamPath, streamPath);
            Assert.True(File.Exists(streamPath));
            AssertPngSignature(File.ReadAllBytes(streamPath));

            string byteOutputDirectory = Path.Combine(directory, "byte-images");
            IReadOnlyList<string> bytePaths = PdfImageExtractor.ExtractImagesByPageRanges(
                stamped,
                byteOutputDirectory,
                "byte-source.pdf",
                PdfPageRange.ParseMany("3"));

            string expectedBytePath = Path.Combine(byteOutputDirectory, "byte-source-page-0003-image-0001.png");
            string bytePath = Assert.Single(bytePaths);
            Assert.Equal(expectedBytePath, bytePath);
            Assert.True(File.Exists(bytePath));
            AssertPngSignature(File.ReadAllBytes(bytePath));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractImagesByPageRanges_WritesRepeatedSourcePageImageFilesWithoutCollisions() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-image-range-repeat-files-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "images");

        try {
            Directory.CreateDirectory(directory);
            byte[] stamped = PdfStamper.StampImage(BuildThreePagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
                PageNumbers = new[] { 1, 3 },
                Width = 24,
                Height = 24
            });
            File.WriteAllBytes(inputPath, stamped);

            IReadOnlyList<string> paths = PdfImageExtractor.ExtractImagesByPageRanges(inputPath, outputDirectory, PdfPageRange.ParseMany("3,1,3"));

            Assert.Equal(3, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003-image-0001.png"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001-image-0002.png"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003-image-0003.png"), paths[2]);
            Assert.Equal(3, new HashSet<string>(paths, StringComparer.OrdinalIgnoreCase).Count);
            foreach (string path in paths) {
                Assert.True(File.Exists(path));
                AssertPngSignature(File.ReadAllBytes(path));
            }
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractImages_RejectsNullInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages((string)null!));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages((PdfReadDocument)null!));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages(null!, "out"));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages("input.pdf", null!));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges((PdfReadDocument)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges(CreateImagePdf(), (PdfPageRange[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges(null!, "out", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImagesByPageRanges("input.pdf", null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfImageExtractor.ExtractImages("input.pdf", " "));
        Assert.Throws<ArgumentException>(() => PdfImageExtractor.ExtractImagesByPageRanges(CreateImagePdf()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfImageExtractor.ExtractImagesByPageRanges(CreateImagePdf(), default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfImageExtractor.ExtractImagesByPageRanges(CreateImagePdf(), PdfPageRange.From(2, 2)));

        using var unreadable = new WriteOnlyStream();
        Assert.Throws<ArgumentException>(() => PdfImageExtractor.ExtractImages(unreadable));
        Assert.Throws<ArgumentException>(() => PdfImageExtractor.ExtractImagesByPageRanges(unreadable, PdfPageRange.From(1, 1)));
    }

    [Fact]
    public void ExtractImages_RejectsFileOutputDirectoryBeforeReadingInput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-image-output-directory-" + Guid.NewGuid().ToString("N"));
        string outputFile = Path.Combine(directory, "not-a-directory");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllText(outputFile, "existing file");

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfImageExtractor.ExtractImages("missing.pdf", outputFile));

            Assert.Equal("outputDirectory", exception.ParamName);
            Assert.Contains("Output directory refers to a file; a directory path is required.", exception.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] BuildTwoPagePdf() {
        var doc = PdfDocument.Create();
        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page body"))));
            });

            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page body"))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] BuildThreePagePdf() {
        var doc = PdfDocument.Create();
        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page body"))));
            });

            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page body"))));
            });

            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Third page body"))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] CreateImagePdf() {
        return PdfDocument.Create().Image(CreateMinimalRgbPng(), 24, 24).ToBytes();
    }

    private static void AssertPngSignature(byte[] bytes) {
        Assert.True(bytes.Length > 8);
        Assert.Equal(137, bytes[0]);
        Assert.Equal(80, bytes[1]);
        Assert.Equal(78, bytes[2]);
        Assert.Equal(71, bytes[3]);
    }

    private static int ReadPngColorType(byte[] bytes) {
        Assert.True(bytes.Length > 25);
        Assert.Equal((byte)'I', bytes[12]);
        Assert.Equal((byte)'H', bytes[13]);
        Assert.Equal((byte)'D', bytes[14]);
        Assert.Equal((byte)'R', bytes[15]);
        return bytes[25];
    }

    private static byte[] DecodeStoredPngIdat(byte[] bytes) {
        using var idat = new MemoryStream();
        int offset = 8;
        while (offset + 12 <= bytes.Length) {
            int length = ReadInt32BigEndian(bytes, offset);
            Assert.True(length >= 0);
            Assert.True(offset + 12 + length <= bytes.Length);
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == "IDAT") {
                idat.Write(bytes, offset + 8, length);
            }

            if (type == "IEND") {
                break;
            }

            offset += 12 + length;
        }

        byte[] compressed = idat.ToArray();
        Assert.True(compressed.Length >= 6);
        Assert.Equal(0x78, compressed[0]);
        using var decoded = new MemoryStream();
        int compressedOffset = 2;
        bool finalBlock;
        do {
            Assert.True(compressedOffset + 5 <= compressed.Length);
            byte header = compressed[compressedOffset++];
            finalBlock = (header & 1) != 0;
            Assert.Equal(0, (header >> 1) & 0x03);

            int length = compressed[compressedOffset] | (compressed[compressedOffset + 1] << 8);
            int nlen = compressed[compressedOffset + 2] | (compressed[compressedOffset + 3] << 8);
            compressedOffset += 4;
            Assert.Equal(0xFFFF, length ^ nlen);
            Assert.True(compressedOffset + length <= compressed.Length - 4);
            decoded.Write(compressed, compressedOffset, length);
            compressedOffset += length;
        } while (!finalBlock);

        return decoded.ToArray();
    }

    private static int ReadInt32BigEndian(byte[] buffer, int offset) {
        return (buffer[offset] << 24) |
               (buffer[offset + 1] << 16) |
               (buffer[offset + 2] << 8) |
               buffer[offset + 3];
    }

    private static MemoryStream BuildPrefixedStream(byte[] pdf) {
        var data = new byte[pdf.Length + 5];
        data[0] = 1;
        data[1] = 2;
        data[2] = 3;
        data[3] = 4;
        data[4] = 5;
        Array.Copy(pdf, 0, data, 5, pdf.Length);
        return new MemoryStream(data);
    }

    private static byte[] BuildUnfilteredDeviceRgbImagePdf() {
        return BuildDeviceRgbImagePdf("abc", string.Empty);
    }

    private static byte[] BuildUnfilteredDeviceCmykImagePdf() {
        return BuildImagePdf("DeviceCMYK", "abc ", string.Empty);
    }

    private static byte[] BuildUnfilteredDeviceGrayInvertedDecodeImagePdf() {
        return BuildImagePdf("DeviceGray", "a", " /Decode [1 0]");
    }

    private static byte[] BuildUnfilteredIccBasedRgbImagePdf() {
        return BuildIccBasedImagePdf(3, "DeviceRGB", "abc");
    }

    private static byte[] BuildUnfilteredIccBasedCmykImagePdf() {
        return BuildIccBasedImagePdf(4, "DeviceCMYK", "abc ");
    }

    private static byte[] BuildUnfilteredIndexedRgbImagePdf() {
        return BuildImagePdfWithColorSpace("[/Indexed /DeviceRGB 1 <FF000000FF00>]", 2, 1, 1, "@", string.Empty);
    }

    private static byte[] BuildUnfilteredIndexedRgbInvertedDecodeImagePdf() {
        return BuildImagePdfWithColorSpace("[/Indexed /DeviceRGB 1 <FF000000FF00>]", 2, 1, 1, "@", " /Decode [1 0]");
    }

    private static byte[] BuildAsciiHexFlateDeviceRgbImagePdf() {
        string encoded = EncodeAsciiHex(BuildStoredZlib(new byte[] { 97, 98, 99 }));
        return BuildDeviceRgbImagePdf(encoded, " /Filter [/ASCIIHexDecode /FlateDecode]");
    }

    private static byte[] BuildDeviceRgbColorKeyMaskImagePdf() {
        return BuildImagePdfWithColorSpace("/DeviceRGB", 2, 1, 8, "abcdef", " /Mask [97 97 98 98 99 99]");
    }

    private static byte[] BuildDeviceRgbExplicitMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "/DeviceRGB",
            2,
            1,
            8,
            "abcdef",
            " /Mask 6 0 R",
            new[] { BuildImageMaskObject(6) },
            7);
    }

    private static byte[] BuildAsciiHexFlateDeviceCmykImagePdf() {
        string encoded = EncodeAsciiHex(BuildStoredZlib(new byte[] { 97, 98, 99, 32 }));
        return BuildImagePdf("DeviceCMYK", encoded, " /Filter [/ASCIIHexDecode /FlateDecode]");
    }

    private static byte[] BuildDeviceCmykSoftMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "/DeviceCMYK",
            1,
            1,
            8,
            "abc ",
            " /SMask 6 0 R",
            new[] { BuildSoftMaskObject(6, 126) },
            7);
    }

    private static byte[] BuildDeviceRgbJpegSoftMaskImagePdf() {
        byte[] jpeg = CreateMinimalJpeg(2, 1);
        return BuildImagePdfWithColorSpace(
            "/DeviceRGB",
            2,
            1,
            8,
            PdfEncoding.Latin1GetString(jpeg),
            " /Filter /DCTDecode /SMask 6 0 R",
            new[] { BuildSoftMaskObject(6, new byte[] { 126, 64 }, 2, 1) },
            7);
    }

    private static byte[] BuildIccBasedCmykSoftMaskImagePdf() {
        string profileObject = BuildIccProfileObject(6, 4, "DeviceCMYK");
        return BuildImagePdfWithColorSpace(
            "[/ICCBased 6 0 R]",
            1,
            1,
            8,
            "abc ",
            " /SMask 7 0 R",
            new[] { profileObject, BuildSoftMaskObject(7, 126) },
            8);
    }

    private static byte[] BuildAsciiHexFlateIndexedRgbImagePdf() {
        string encoded = EncodeAsciiHex(BuildStoredZlib(new byte[] { 64 }));
        return BuildImagePdfWithColorSpace("[/Indexed /DeviceRGB 1 <FF000000FF00>]", 2, 1, 1, encoded, " /Filter [/ASCIIHexDecode /FlateDecode]");
    }

    private static byte[] BuildIndexedSoftMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "[/Indexed /DeviceRGB 1 <FF000000FF00>]",
            2,
            1,
            1,
            "@",
            " /SMask 6 0 R",
            new[] { BuildSoftMaskObject(6, new byte[] { 126, 64 }, 2, 1) },
            7);
    }

    private static byte[] BuildIndexedColorKeyMaskImagePdf() {
        return BuildImagePdfWithColorSpace(
            "[/Indexed /DeviceRGB 1 <FF000000FF00>]",
            2,
            1,
            1,
            "@",
            " /Mask [0 0]");
    }

    private static byte[] BuildImageMaskPdf() {
        string content = string.Join("\n", new[] {
            "q",
            "24 0 0 24 36 160 cm",
            "/Im1 Do",
            "Q"
        });

        string imageStreamData = "@";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 220 220] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ImageMask true /Length " + Encoding.ASCII.GetByteCount(imageStreamData).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            imageStreamData,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildDeviceRgbImagePdf(string imageStreamData, string imageDictionarySuffix) {
        return BuildImagePdf("DeviceRGB", imageStreamData, imageDictionarySuffix);
    }

    private static byte[] BuildImagePdf(string colorSpace, string imageStreamData, string imageDictionarySuffix) {
        return BuildImagePdfWithColorSpace("/" + colorSpace, 1, 1, 8, imageStreamData, imageDictionarySuffix);
    }

    private static byte[] BuildImagePdfWithColorSpace(
        string colorSpaceObject,
        int width,
        int height,
        int bitsPerComponent,
        string imageStreamData,
        string imageDictionarySuffix) {
        return BuildImagePdfWithColorSpace(colorSpaceObject, width, height, bitsPerComponent, imageStreamData, imageDictionarySuffix, Array.Empty<string>(), 6);
    }

    private static byte[] BuildIccBasedImagePdf(int componentCount, string alternateColorSpace, string imageStreamData) {
        string profileObject = BuildIccProfileObject(6, componentCount, alternateColorSpace);
        return BuildImagePdfWithColorSpace("[/ICCBased 6 0 R]", 1, 1, 8, imageStreamData, string.Empty, new[] { profileObject }, 7);
    }

    private static string BuildIccProfileObject(int objectNumber, int componentCount, string alternateColorSpace) {
        return string.Join("\n", new[] {
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj",
            "<< /N " + componentCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Alternate /" + alternateColorSpace + " /Length 0 >>",
            "stream",
            string.Empty,
            "endstream",
            "endobj"
        });
    }

    private static string BuildSoftMaskObject(int objectNumber, byte alpha) {
        return BuildSoftMaskObject(objectNumber, new[] { alpha }, 1, 1);
    }

    private static string BuildSoftMaskObject(int objectNumber, byte[] alpha, int width, int height) {
        string encodedAlpha = EncodeAsciiHex(BuildStoredZlib(alpha));
        return string.Join("\n", new[] {
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj",
            "<< /Type /XObject /Subtype /Image /Width "
                + width.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /Height "
                + height.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter [/ASCIIHexDecode /FlateDecode] /Length "
                + Encoding.ASCII.GetByteCount(encodedAlpha).ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " >>",
            "stream",
            encodedAlpha,
            "endstream",
            "endobj"
        });
    }

    private static string BuildImageMaskObject(int objectNumber) {
        string maskData = PdfEncoding.Latin1GetString(new byte[] { 0x80 });
        return string.Join("\n", new[] {
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ImageMask true /BitsPerComponent 1 /Length "
                + maskData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " >>",
            "stream",
            maskData,
            "endstream",
            "endobj"
        });
    }

    private static byte[] BuildImagePdfWithColorSpace(
        string colorSpaceObject,
        int width,
        int height,
        int bitsPerComponent,
        string imageStreamData,
        string imageDictionarySuffix,
        IReadOnlyList<string> additionalObjects,
        int trailerSize) {
        string content = string.Join("\n", new[] {
            "q",
            "24 0 0 24 36 160 cm",
            "/Im1 Do",
            "Q"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 220 220] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width "
                + width.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /Height "
                + height.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /ColorSpace "
                + colorSpaceObject
                + " /BitsPerComponent "
                + bitsPerComponent.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " /Length "
                + imageStreamData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + imageDictionarySuffix
                + " >>",
            "stream",
            imageStreamData,
            "endstream",
            "endobj",
            string.Join("\n", additionalObjects),
            "trailer",
            "<< /Root 1 0 R /Size " + trailerSize.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "%%EOF"
        }) + "\n";

        return PdfEncoding.Latin1GetBytes(pdf);
    }

    private static string EncodeAsciiHex(byte[] data) {
        var builder = new StringBuilder(data.Length * 2 + 1);
        for (int i = 0; i < data.Length; i++) {
            builder.Append(data[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }

        builder.Append('>');
        return builder.ToString();
    }

    private sealed class WriteOnlyStream : Stream {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => 0;

        public override long Position {
            get => 0;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin) {
            throw new NotSupportedException();
        }

        public override void SetLength(long value) {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count) {
        }
    }

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(255, 0, 0);

    private static byte[] CreateMinimalRgbaPng() => PdfPngTestImages.CreateRgbaPng(255, 0, 0, 128);

    private static byte[] CreateMinimalJpeg(int width, int height) {
        return new byte[] {
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            (byte)(height >> 8), (byte)(height & 0xFF),
            (byte)(width >> 8), (byte)(width & 0xFF),
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        };
    }

    private static byte[] CreateMinimalRgbTransparencyPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0
        });
        WritePngChunk(ms, "tRNS", new byte[] {
            0, 255,
            0, 0,
            0, 0
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] { 0, 255, 0, 0 }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] CreateMinimalGrayscaleTransparencyPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 0, 0, 0, 0
        });
        WritePngChunk(ms, "tRNS", new byte[] { 0, 128 });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] { 0, 128 }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] CreateMinimalPackedGrayscaleTransparencyPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 2,
            0, 0, 0, 1,
            4, 0, 0, 0, 0
        });
        WritePngChunk(ms, "tRNS", new byte[] { 0, 1 });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] { 0, 0x01 }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static void WritePngChunk(Stream stream, string type, byte[] data) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        var length = new byte[4];
        WriteInt32BigEndian(length, 0, data.Length);
        stream.Write(length, 0, length.Length);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        uint crc = ComputeCrc32(typeBytes, data);
        var crcBytes = new byte[4];
        WriteUInt32BigEndian(crcBytes, 0, crc);
        stream.Write(crcBytes, 0, crcBytes.Length);
    }

    private static byte[] BuildStoredZlib(byte[] scanline) {
        using var ms = new MemoryStream();
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);
        ms.WriteByte(0x01);
        ms.WriteByte((byte)(scanline.Length & 0xFF));
        ms.WriteByte((byte)((scanline.Length >> 8) & 0xFF));
        int nlen = scanline.Length ^ 0xFFFF;
        ms.WriteByte((byte)(nlen & 0xFF));
        ms.WriteByte((byte)((nlen >> 8) & 0xFF));
        ms.Write(scanline, 0, scanline.Length);
        uint adler = Adler32(scanline);
        ms.WriteByte((byte)((adler >> 24) & 0xFF));
        ms.WriteByte((byte)((adler >> 16) & 0xFF));
        ms.WriteByte((byte)((adler >> 8) & 0xFF));
        ms.WriteByte((byte)(adler & 0xFF));
        return ms.ToArray();
    }

    private static void WriteInt32BigEndian(byte[] buffer, int offset, int value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static void WriteUInt32BigEndian(byte[] buffer, int offset, uint value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static uint ComputeCrc32(byte[] typeBytes, byte[] data) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < typeBytes.Length; i++) {
            crc = UpdateCrc32(crc, typeBytes[i]);
        }

        for (int i = 0; i < data.Length; i++) {
            crc = UpdateCrc32(crc, data[i]);
        }

        return crc ^ 0xFFFFFFFF;
    }

    private static uint UpdateCrc32(uint crc, byte value) {
        crc ^= value;
        for (int i = 0; i < 8; i++) {
            if ((crc & 1) != 0) {
                crc = (crc >> 1) ^ 0xEDB88320;
            } else {
                crc >>= 1;
            }
        }

        return crc;
    }
}
