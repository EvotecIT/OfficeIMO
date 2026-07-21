using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentImageValidationTests {
    [Fact]
    public void ImageApis_PreservePreAlternativeTextClrSignatures() {
        Assert.NotNull(typeof(PdfDocument).GetMethod(nameof(PdfDocument.Image), new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign?),
            typeof(OfficeClipPath),
            typeof(OfficeImageFit?),
            typeof(double?),
            typeof(double?),
            typeof(PdfImageStyle),
            typeof(string),
            typeof(string)
        }));
        Assert.NotNull(typeof(PdfItemCompose).GetMethod(nameof(PdfItemCompose.Image), new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign?),
            typeof(OfficeClipPath),
            typeof(OfficeImageFit?),
            typeof(double?),
            typeof(double?),
            typeof(PdfImageStyle),
            typeof(string),
            typeof(string)
        }));
        Assert.NotNull(typeof(PdfElementCompose).GetMethod(nameof(PdfElementCompose.Image), new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign?),
            typeof(OfficeClipPath),
            typeof(OfficeImageFit?),
            typeof(double?),
            typeof(double?),
            typeof(PdfImageStyle),
            typeof(string),
            typeof(string)
        }));
        Assert.NotNull(typeof(PdfRowColumnCompose).GetMethod(nameof(PdfRowColumnCompose.Image), new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign?),
            typeof(OfficeClipPath),
            typeof(OfficeImageFit?),
            typeof(double?),
            typeof(double?),
            typeof(PdfImageStyle),
            typeof(string),
            typeof(string)
        }));
        Assert.NotNull(typeof(PdfHeaderCompose).GetMethod(nameof(PdfHeaderCompose.Image), new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign),
            typeof(OfficeImageFit)
        }));
        Assert.NotNull(typeof(PdfFooterCompose).GetMethod(nameof(PdfFooterCompose.Image), new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign),
            typeof(OfficeImageFit)
        }));
        Assert.NotNull(typeof(PdfHeaderFooterImage).GetConstructor(new[] {
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(PdfAlign),
            typeof(OfficeImageFit)
        }));
    }

    [Fact]
    public void Image_WithNullBytes_ThrowsArgumentNullException() {
        var doc = PdfDocument.Create();

        var exception = Assert.Throws<ArgumentNullException>(() => doc.Image(null!, 24, 24));

        Assert.Equal("jpegBytes", exception.ParamName);
        Assert.Contains("Parameter 'jpegBytes' cannot be null.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithEmptyBytes_ThrowsArgumentException() {
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => doc.Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, 10, invalidHeight));

        Assert.Equal("height", exception.ParamName);
        Assert.Contains("Parameter 'height' must be a finite positive number.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageBlock_SnapshotsImageBytes() {
        var data = new byte[] { 1, 2, 3 };
        var imageInfo = new OfficeImageInfo(OfficeImageFormat.Png, 1, 1);

        var block = new ImageBlock(data, 24, 24, PdfAlign.Left, imageInfo);

        data[0] = 9;

        Assert.Equal(1, block.Data[0]);
    }

    [Fact]
    public void TableCellImage_RendersInspectableImageInsideCell() {
        byte[] png = CreateMinimalRgbPng();

        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] {
                    PdfTableCell.WithImages(
                        "Logo",
                        new[] { new PdfTableCellImage(png, 24, 24) })
                }
            })
            .ToBytes();

        string content = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", content);
        Assert.Single(PdfImageExtractor.ExtractImages(bytes));

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Logo", pdf.GetPage(1).Text);
    }

    [Fact]
    public void RepeatedIdenticalImages_ReuseImageXObjectAcrossPages() {
        byte[] png = CreateMinimalRgbPng();

        byte[] bytes = PdfDocument.Create()
            .Image(png, 24, 24)
            .Image(png, 12, 12)
            .PageBreak()
            .Image(png, 18, 18)
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        int imageObjectCount = CountOccurrences(pdfContent, "/Subtype /Image");
        int imageDrawCount = CountOccurrences(pdfContent, "/Im1 Do");

        Assert.Equal(1, imageObjectCount);
        Assert.True(imageDrawCount >= 3, "Expected all repeated image placements to draw the shared XObject.");
    }

    [Fact]
    public void TableCellImage_WithScaleDownToFit_ReducesOversizedImageIntoCellFrame() {
        byte[] jpeg = CreateMinimalJpeg(400, 200);
        var tableStyle = new PdfTableStyle {
            HeaderRowCount = 0,
            BorderColor = null,
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 80D }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.WithImages(
                        (string?)null,
                        new[] {
                            new PdfTableCellImage(jpeg, 144, 72, new PdfImageStyle { ScaleDownToFit = true })
                        })
                }
            }, style: tableStyle)
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n72 0 0 36 24 122 cm\n/Im1 Do\nQ", pdfContent);
    }

    [Fact]
    public void RowColumnTableCellImage_WithScaleDownToFit_ReducesOversizedImageIntoCellFrame() {
        byte[] jpeg = CreateMinimalJpeg(400, 200);
        var tableStyle = new PdfTableStyle {
            HeaderRowCount = 0,
            BorderColor = null,
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 80D }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.WithImages(
                                            (string?)null,
                                            new[] {
                                                new PdfTableCellImage(jpeg, 144, 72, new PdfImageStyle { ScaleDownToFit = true })
                                            })
                                    }
                                }, style: tableStyle))))))
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n72 0 0 36 24 122 cm\n/Im1 Do\nQ", pdfContent);
    }

    [Fact]
    public void Options_SnapshotDefaultImageStyle() {
        var style = new PdfImageStyle {
            Align = PdfAlign.Center,
            Fit = OfficeImageFit.Contain,
            ClipPath = OfficeClipPath.Rectangle(12, 10),
            SpacingBefore = 4,
            SpacingAfter = 9,
            KeepWithNext = true,
            ScaleDownToFit = true,
            RotationAngle = 12
        };
        var options = new PdfOptions {
            DefaultImageStyle = style
        };

        style.Align = PdfAlign.Right;
        style.Fit = OfficeImageFit.Cover;
        style.ClipPath = OfficeClipPath.Rectangle(4, 4);
        style.SpacingBefore = 1;
        style.SpacingAfter = 2;
        style.KeepWithNext = false;
        style.ScaleDownToFit = false;
        style.RotationAngle = 0;

        PdfImageStyle readback = options.DefaultImageStyle!;
        readback.Align = PdfAlign.Left;
        readback.ClipPath = OfficeClipPath.Rectangle(2, 2);
        readback.ScaleDownToFit = false;
        readback.RotationAngle = 3;

        PdfOptions clone = options.Clone();

        Assert.Equal(PdfAlign.Center, options.DefaultImageStyle!.Align);
        Assert.Equal(OfficeImageFit.Contain, options.DefaultImageStyle.Fit);
        Assert.Equal(12, options.DefaultImageStyle.ClipPath!.Width);
        Assert.Equal(4, options.DefaultImageStyle.SpacingBefore);
        Assert.Equal(9, options.DefaultImageStyle.SpacingAfter);
        Assert.True(options.DefaultImageStyle.KeepWithNext);
        Assert.True(options.DefaultImageStyle.ScaleDownToFit);
        Assert.Equal(12, options.DefaultImageStyle.RotationAngle);
        Assert.Equal(PdfAlign.Center, clone.DefaultImageStyle!.Align);
        Assert.Equal(12, clone.DefaultImageStyle.ClipPath!.Width);
        Assert.True(clone.DefaultImageStyle.KeepWithNext);
        Assert.True(clone.DefaultImageStyle.ScaleDownToFit);
        Assert.Equal(12, clone.DefaultImageStyle.RotationAngle);
    }

    [Fact]
    public void Options_ApplyThemeSnapshotsDefaultImageStyle() {
        var imageStyle = new PdfImageStyle {
            Align = PdfAlign.Center,
            Fit = OfficeImageFit.Contain,
            SpacingBefore = 3,
            SpacingAfter = 8,
            KeepWithNext = true,
            ScaleDownToFit = true
        };
        var theme = new PdfTheme {
            ImageStyle = imageStyle
        };
        var options = new PdfOptions().ApplyTheme(theme);

        imageStyle.Align = PdfAlign.Right;
        imageStyle.Fit = OfficeImageFit.Cover;
        imageStyle.SpacingAfter = 1;
        imageStyle.KeepWithNext = false;
        imageStyle.ScaleDownToFit = false;

        PdfOptions clone = options.Clone();

        Assert.Equal(PdfAlign.Center, options.DefaultImageStyle!.Align);
        Assert.Equal(OfficeImageFit.Contain, options.DefaultImageStyle.Fit);
        Assert.Equal(8, options.DefaultImageStyle.SpacingAfter);
        Assert.True(options.DefaultImageStyle.KeepWithNext);
        Assert.True(options.DefaultImageStyle.ScaleDownToFit);
        Assert.Equal(PdfAlign.Center, clone.DefaultImageStyle!.Align);
        Assert.True(clone.DefaultImageStyle.KeepWithNext);
        Assert.True(clone.DefaultImageStyle.ScaleDownToFit);
    }

    [Fact]
    public void ImageXObjectDictionaryBuilder_EmitsGeneratedImageDictionariesAndSoftMasks() {
        var image = new PdfWriter.PdfImageStream {
            Data = new byte[] { 1, 2, 3 },
            PixelWidth = 16,
            PixelHeight = 8,
            DictionarySuffix = " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode"
        };

        Assert.Equal(
            "<< /Type /XObject /Subtype /Image /Width 16 /Height 8 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /SMask 5 0 R /Length 3 >>",
            PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(image, 5));

        PdfStream stream = PdfImageXObjectDictionaryBuilder.BuildStreamObject(image, -1001);
        Assert.Equal("XObject", Assert.IsType<PdfName>(stream.Dictionary.Items["Type"]).Name);
        Assert.Equal("Image", Assert.IsType<PdfName>(stream.Dictionary.Items["Subtype"]).Name);
        Assert.Equal(16, Assert.IsType<PdfNumber>(stream.Dictionary.Items["Width"]).Value);
        Assert.Equal(8, Assert.IsType<PdfNumber>(stream.Dictionary.Items["Height"]).Value);
        Assert.Equal("DeviceRGB", Assert.IsType<PdfName>(stream.Dictionary.Items["ColorSpace"]).Name);
        Assert.Equal("DCTDecode", Assert.IsType<PdfName>(stream.Dictionary.Items["Filter"]).Name);
        var softMask = Assert.IsType<PdfReference>(stream.Dictionary.Items["SMask"]);
        Assert.Equal(-1001, softMask.ObjectNumber);
    }

    [Fact]
    public void ImageXObjectDictionaryBuilder_EmitsPngPredictorDecodeParmsAndRejectsInvalidStreams() {
        var image = new PdfWriter.PdfImageStream {
            Data = new byte[] { 1, 2, 3 },
            PixelWidth = 10,
            PixelHeight = 5,
            DictionarySuffix = " /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors 1 /BitsPerComponent 8 /Columns 10 >>"
        };

        PdfStream stream = PdfImageXObjectDictionaryBuilder.BuildStreamObject(image);
        Assert.Equal("DeviceGray", Assert.IsType<PdfName>(stream.Dictionary.Items["ColorSpace"]).Name);
        Assert.Equal("FlateDecode", Assert.IsType<PdfName>(stream.Dictionary.Items["Filter"]).Name);
        var decodeParms = Assert.IsType<PdfDictionary>(stream.Dictionary.Items["DecodeParms"]);
        Assert.Equal(15, Assert.IsType<PdfNumber>(decodeParms.Items["Predictor"]).Value);
        Assert.Equal(1, Assert.IsType<PdfNumber>(decodeParms.Items["Colors"]).Value);
        Assert.Equal(10, Assert.IsType<PdfNumber>(decodeParms.Items["Columns"]).Value);

        Assert.Throws<ArgumentException>(() =>
            PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(new PdfWriter.PdfImageStream {
                Data = Array.Empty<byte>(),
                PixelWidth = 10,
                PixelHeight = 5,
                DictionarySuffix = image.DictionarySuffix
            }));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(new PdfWriter.PdfImageStream {
                Data = new byte[] { 1 },
                PixelWidth = 0,
                PixelHeight = 5,
                DictionarySuffix = image.DictionarySuffix
            }));
    }

    [Fact]
    public void ImageBlock_RejectsInvalidModelState() {
        var imageInfo = new OfficeImageInfo(OfficeImageFormat.Png, 1, 1);

        Assert.Throws<ArgumentNullException>(() =>
            new ImageBlock(null!, 24, 24, PdfAlign.Left, imageInfo));

        Assert.Throws<ArgumentException>(() =>
            new ImageBlock(Array.Empty<byte>(), 24, 24, PdfAlign.Left, imageInfo));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ImageBlock(new byte[] { 1 }, 0, 24, PdfAlign.Left, imageInfo));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ImageBlock(new byte[] { 1 }, 24, double.PositiveInfinity, PdfAlign.Left, imageInfo));

        Assert.Throws<ArgumentException>(() =>
            new ImageBlock(new byte[] { 1 }, 24, 24, PdfAlign.Justify, imageInfo));

        Assert.Throws<ArgumentNullException>(() =>
            new ImageBlock(new byte[] { 1 }, 24, 24, PdfAlign.Left, null!));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ImageBlock(new byte[] { 1 }, 24, 24, PdfAlign.Left, imageInfo, fit: (OfficeImageFit)42));
    }

    [Theory]
    [MemberData(nameof(NormalizedRasterPayloads))]
    public void Image_WithDrawingSupportedRaster_NormalizesAndEmbeds(
        byte[] source,
        OfficeImageFormat sourceFormat) {
        Assert.True(PdfDocument.TryPrepareImageBytes(
            source,
            out byte[] prepared,
            out OfficeImageInfo? imageInfo,
            out bool wasTranscoded,
            out string? unsupportedReason), unsupportedReason);
        Assert.True(wasTranscoded);
        Assert.NotNull(imageInfo);
        Assert.Equal(OfficeImageFormat.Png, imageInfo!.Format);
        Assert.Equal(sourceFormat, OfficeImageReader.Identify(source).Format);
        Assert.True(OfficeImageReader.TryIdentify(prepared, null, out OfficeImageInfo preparedInfo));
        Assert.Equal(OfficeImageFormat.Png, preparedInfo.Format);

        byte[] pdf = PdfDocument.Create()
            .Image(source, 24, 24)
            .ToBytes();

        Assert.Single(PdfImageExtractor.ExtractImages(pdf));
    }

    [Fact]
    public void RowColumnImage_WithNullBytes_ThrowsArgumentNullException() {
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

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
    public void RowColumnImage_WithDrawingSupportedRaster_NormalizesAndEmbeds() {
        var doc = PdfDocument.Create();

        doc.Compose(compose =>
            compose.Page(page =>
                page.Content(content =>
                    content.Row(row =>
                        row.Column(100, column =>
                            column.Image(CreateMinimalGif(), 24, 24))))));

        Assert.Single(PdfImageExtractor.ExtractImages(doc.ToBytes()));
    }

    [Fact]
    public void WordPdfAdapterUsesCanonicalRasterPreparation() {
        using WordDocument document = WordDocument.Create();
        using var image = new MemoryStream(CreateMinimalGif());
        document.AddParagraph().AddImage(image, "pixel.gif", 24, 24);

        byte[] pdf = document.ToPdf();

        PdfExtractedImage embedded = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal("image/png", embedded.MimeType);
    }

    [Fact]
    public void ExcelPdfAdapterUsesCanonicalRasterPreparation() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Raster");
        sheet.AddImage(1, 1, CreateMinimalGif(), "image/gif", 24, 24);

        byte[] pdf = workbook.ToPdf();

        PdfExtractedImage embedded = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal("image/png", embedded.MimeType);
    }

    [Fact]
    public void Image_WithUnrecognizedPayload_ReportsTheSharedContract() {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create().Image(new byte[] { 1, 2, 3, 4 }, 24, 24));

        Assert.Contains("raster formats decoded by OfficeIMO.Drawing", exception.Message, StringComparison.Ordinal);
        Assert.Contains("not recognized", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterNormalizationPreservesPhysicalResolution() {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Red);
        byte[] tiff = OfficeRasterImageEncoder.Encode(
            image,
            OfficeImageExportFormat.Tiff,
            new OfficeRasterEncodingOptions { DpiX = 144, DpiY = 120 });

        Assert.True(PdfDocument.TryPrepareImageBytes(
            tiff,
            out byte[] prepared,
            out OfficeImageInfo? info,
            out bool wasTranscoded,
            out string? unsupportedReason), unsupportedReason);

        Assert.True(wasTranscoded);
        Assert.NotNull(info);
        Assert.InRange(info!.DpiX, 143.98D, 144.02D);
        Assert.InRange(info.DpiY, 119.98D, 120.02D);
        OfficeImageInfo preparedInfo = OfficeImageReader.Identify(prepared);
        Assert.Equal(info.DpiX, preparedInfo.DpiX);
        Assert.Equal(info.DpiY, preparedInfo.DpiY);
    }

    [Fact]
    public void ItemComposeImage_WithNullBytes_ThrowsArgumentNullException() {
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

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
        var doc = PdfDocument.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Image(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, 24, invalidHeight))))));

        Assert.Equal("height", exception.ParamName);
    }

    [Fact]
    public void Image_WithJpegHeader_UsesOfficeImageReaderMetadata() {
        byte[] jpeg = CreateMinimalJpeg(32, 16);

        Assert.True(OfficeImageReader.TryIdentify(jpeg, null, out var info));
        Assert.Equal(OfficeImageFormat.Jpeg, info.Format);
        Assert.Equal(32, info.Width);
        Assert.Equal(16, info.Height);

        byte[] bytes = PdfDocument.Create().Image(jpeg, 24, 12).ToBytes();

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void Image_WithHeightExceedingContentArea_ThrowsArgumentException() {
        byte[] jpeg = CreateMinimalJpeg(32, 16);
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 140,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Image(jpeg, 80, 130)
                .ToBytes());

        Assert.Contains("Image height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithWidthExceedingContentArea_ThrowsArgumentException() {
        byte[] jpeg = CreateMinimalJpeg(32, 16);
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Image(jpeg, 190, 40)
                .ToBytes());

        Assert.Contains("Image width exceeds the available page content width.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithScaleDownToFit_ReducesOversizedImageIntoContentFrame() {
        byte[] jpeg = CreateMinimalJpeg(400, 200);
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        byte[] bytes = PdfDocument.Create(options)
            .Image(jpeg, 360, 180, style: new PdfImageStyle { ScaleDownToFit = true })
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n180 0 0 90 20 70 cm\n/Im1 Do\nQ", pdfContent);
    }

    [Fact]
    public void Image_WithSimpleRgbPng_WritesFlatePngPredictorImageObject() {
        byte[] png = CreateMinimalRgbPng();

        byte[] bytes = PdfDocument.Create().Image(png, 24, 24).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 1 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/Predictor 15", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithRgbaPng_WritesSoftMaskImageObject() {
        byte[] png = CreateMinimalRgbaPng();

        byte[] bytes = PdfDocument.Create().Image(png, 24, 24).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 1 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithRgbPngTransparency_WritesSoftMaskImageObject() {
        byte[] png = CreateMinimalRgbTransparencyPng();

        byte[] bytes = PdfDocument.Create().Image(png, 24, 24).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 1 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithGrayscalePngTransparency_WritesSoftMaskImageObject() {
        byte[] png = CreateMinimalGrayscaleTransparencyPng();

        byte[] bytes = PdfDocument.Create().Image(png, 24, 24).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 1 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    public void Image_WithPackedGrayscalePngBitDepth_WritesDeviceGrayImageObject(int bitDepth) {
        byte[] png = CreateMinimalPackedGrayscalePng(bitDepth, includeTransparency: false);

        byte[] bytes = PdfDocument.Create().Image(png, 24, 12).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 2 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/BitsPerComponent 8", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);
        Assert.DoesNotContain("/SMask", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithPackedGrayscalePngTransparency_WritesSoftMaskImageObject() {
        byte[] png = CreateMinimalPackedGrayscalePng(4, includeTransparency: true);

        byte[] bytes = PdfDocument.Create().Image(png, 24, 12).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 2 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/BitsPerComponent 8", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithIndexedColorPng_WritesRgbImageObject() {
        byte[] png = CreateMinimalIndexedColorPng(includeTransparency: false);

        Assert.True(PdfDocument.TryValidateImageBytes(png, out OfficeImageInfo? imageInfo, out string? unsupportedReason));
        Assert.Null(unsupportedReason);
        Assert.NotNull(imageInfo);
        Assert.Equal(OfficeImageFormat.Png, imageInfo!.Format);

        byte[] bytes = PdfDocument.Create().Image(png, 24, 12).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 2 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.DoesNotContain("/SMask", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    [InlineData(8)]
    public void Image_WithIndexedColorPngBitDepth_WritesRgbImageObject(int bitDepth) {
        byte[] png = CreateMinimalIndexedColorPng(includeTransparency: false, bitDepth: bitDepth);

        byte[] bytes = PdfDocument.Create().Image(png, 24, 12).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Width 2 /Height 1", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.DoesNotContain("/SMask", pdfContent);
    }

    [Fact]
    public void Image_WithIndexedColorPngTransparency_WritesSoftMaskImageObject() {
        byte[] png = CreateMinimalIndexedColorPng(includeTransparency: true);

        byte[] bytes = PdfDocument.Create().Image(png, 24, 12).ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/Width 2 /Height 1", pdfContent);
        Assert.Contains("/Filter /FlateDecode", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/Colors 3", pdfContent);
        Assert.Contains("/Colors 1", pdfContent);
        Assert.DoesNotContain("/Filter /DCTDecode", pdfContent);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Image_WithContainFit_PreservesAspectRatioInsideTargetBox() {
        byte[] jpeg = CreateMinimalJpeg(200, 100);

        byte[] bytes = PdfDocument.Create()
            .Image(jpeg, 100, 100, fit: OfficeImageFit.Contain)
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n100 0 0 50 72 645 cm\n/Im1 Do\nQ", pdfContent);
        Assert.DoesNotContain("W n\nq\n100 0 0 50", pdfContent);
    }

    [Fact]
    public void Image_WithCoverFit_PreservesAspectRatioAndClipsToTargetBox() {
        byte[] jpeg = CreateMinimalJpeg(200, 100);

        byte[] bytes = PdfDocument.Create()
            .Image(jpeg, 100, 100, fit: OfficeImageFit.Cover)
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n72 620 100 100 re W n\nq\n200 0 0 100 22 620 cm\n/Im1 Do\nQ\nQ", pdfContent);
    }

    [Fact]
    public void Image_WithInvalidFit_ThrowsArgumentOutOfRangeException() {
        byte[] png = CreateMinimalRgbPng();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().Image(png, 24, 24, fit: (OfficeImageFit)42));

        Assert.Equal("fit", exception.ParamName);
        Assert.Contains("Unsupported image fit mode.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_WithClipPath_WritesClippingPathAroundImageXObject() {
        byte[] png = CreateMinimalRgbPng();

        byte[] bytes = PdfDocument.Create()
            .Image(png, 24, 24, clipPath: OfficeClipPath.Rectangle(12, 12))
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n72 708 12 12 re W n\nq\n24 0 0 24 72 696 cm\n/Im1 Do\nQ\nQ", pdfContent);
        Assert.Contains("/Subtype /Image", pdfContent);
    }

    [Fact]
    public void Image_WithClipPathLargerThanImage_ThrowsArgumentOutOfRangeException() {
        byte[] png = CreateMinimalRgbPng();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().Image(png, 24, 24, clipPath: OfficeClipPath.Rectangle(25, 24)));

        Assert.Equal("clipPath", exception.ParamName);
        Assert.Contains("Clip paths must fit inside the image width and height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DefaultImageStyle_AppliesAlignmentFitAndSpacingToFollowingImages() {
        byte[] jpeg = CreateMinimalJpeg(200, 100);
        var style = new PdfImageStyle {
            Align = PdfAlign.Center,
            Fit = OfficeImageFit.Contain,
            SpacingBefore = 4,
            SpacingAfter = 12
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .DefaultImageStyle(style)
            .Image(jpeg, 100, 100)
            .Paragraph(p => p.Text("AfterDefaultImage"))
            .ToBytes();

        style.Align = PdfAlign.Right;
        style.Fit = OfficeImageFit.Cover;
        style.SpacingAfter = 0;

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        double imageBottomY = 60;
        double paragraphTopY = FindWordStartY(pdf.GetPage(1), "AfterDefaultImage") + 10 * 0.74;
        double clearance = imageBottomY - paragraphTopY;

        Assert.Contains("q\n100 0 0 50 70 85 cm\n/Im1 Do\nQ", pdfContent);
        Assert.True(clearance >= 11, $"Expected default image spacing to leave visible rhythm. Clearance: {clearance:0.##}pt.");
    }

    [Fact]
    public void RowColumnImage_UsesDefaultImageStyleWhenStyleIsNotProvided() {
        byte[] png = CreateMinimalRgbPng();
        var style = new PdfImageStyle {
            Align = PdfAlign.Center,
            SpacingBefore = 4,
            SpacingAfter = 8
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .DefaultImageStyle(style)
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(png, 24, 24))))))
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n24 0 0 24 108 136 cm\n/Im1 Do\nQ", pdfContent);
    }

    [Fact]
    public void RowColumnImage_WithScaleDownToFit_ReducesOversizedImageIntoColumnFrame() {
        byte[] jpeg = CreateMinimalJpeg(400, 200);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(jpeg, 360, 180, style: new PdfImageStyle { ScaleDownToFit = true }))))))
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n180 0 0 90 20 70 cm\n/Im1 Do\nQ", pdfContent);
    }

    [Fact]
    public void RowColumnImage_WithClipPath_WritesClippingPathAroundImageXObject() {
        byte[] png = CreateMinimalRgbPng();

        byte[] bytes = PdfDocument.Create()
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(60, column =>
                                column.Image(png, 24, 24, clipPath: OfficeClipPath.RoundedRectangle(12, 12, 3)))))))
            .ToBytes();

        string pdfContent = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("75 708 m", pdfContent);
        Assert.Contains("W n\nq\n24 0 0 24", pdfContent);
        Assert.Contains("/Im1 Do\nQ\nQ", pdfContent);
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static int CountOccurrences(string text, string value) =>
        text.Split(new[] { value }, StringSplitOptions.None).Length - 1;

    private static byte[] CreateMinimalRgbPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", new byte[] { 0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00 });
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] CreateMinimalRgbaPng() {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 6, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", new byte[] {
            0x78, 0x01, 0x01, 0x05, 0x00, 0xFA, 0xFF, 0x00,
            0xFF, 0x00, 0x00, 0x80, 0x04, 0x81, 0x01, 0x80
        });
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
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

    private static byte[] CreateMinimalPackedGrayscalePng(int bitDepth, bool includeTransparency) {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 2,
            0, 0, 0, 1,
            (byte)bitDepth, 0, 0, 0, 0
        });
        if (includeTransparency) {
            WritePngChunk(ms, "tRNS", new byte[] { 0, 1 });
        }

        WritePngChunk(ms, "IDAT", BuildPackedGrayscalePngIdat(bitDepth));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] BuildPackedGrayscalePngIdat(int bitDepth) {
        byte packedPixels = bitDepth switch {
            1 => 0x40,
            2 => 0x10,
            4 => 0x01,
            _ => throw new ArgumentOutOfRangeException(nameof(bitDepth))
        };

        return BuildStoredZlib(new byte[] { 0, packedPixels });
    }

    private static byte[] CreateMinimalIndexedColorPng(bool includeTransparency, int bitDepth = 8) {
        using var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 2,
            0, 0, 0, 1,
            (byte)bitDepth, 3, 0, 0, 0
        });
        WritePngChunk(ms, "PLTE", new byte[] {
            0xE6, 0x39, 0x46,
            0x2B, 0x7D, 0xD8
        });
        if (includeTransparency) {
            WritePngChunk(ms, "tRNS", new byte[] { 255, 64 });
        }

        WritePngChunk(ms, "IDAT", BuildIndexedPngIdat(bitDepth));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] BuildIndexedPngIdat(int bitDepth) {
        byte packedPixels = bitDepth switch {
            1 => 0x40,
            2 => 0x10,
            4 => 0x01,
            8 => 0x00,
            _ => throw new ArgumentOutOfRangeException(nameof(bitDepth))
        };
        byte[] scanline = bitDepth == 8
            ? new byte[] { 0, 0, 1 }
            : new byte[] { 0, packedPixels };

        return BuildStoredZlib(scanline);
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

    private static void WritePngChunk(Stream stream, string type, byte[] data) {
        stream.WriteByte((byte)((data.Length >> 24) & 0xFF));
        stream.WriteByte((byte)((data.Length >> 16) & 0xFF));
        stream.WriteByte((byte)((data.Length >> 8) & 0xFF));
        stream.WriteByte((byte)(data.Length & 0xFF));
        byte[] typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);
        uint crc = ComputeCrc32(typeBytes, data);
        stream.WriteByte((byte)((crc >> 24) & 0xFF));
        stream.WriteByte((byte)((crc >> 16) & 0xFF));
        stream.WriteByte((byte)((crc >> 8) & 0xFF));
        stream.WriteByte((byte)(crc & 0xFF));
    }

    private static byte[] CreateMinimalGif() =>
        Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

    public static System.Collections.Generic.IEnumerable<object[]> NormalizedRasterPayloads() {
        var image = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
        image.SetPixel(0, 0, OfficeColor.Red);
        image.SetPixel(1, 0, OfficeColor.Blue);
        yield return new object[] { CreateMinimalGif(), OfficeImageFormat.Gif };
        yield return new object[] { CreateMinimalBmp(), OfficeImageFormat.Bmp };
        yield return new object[] {
            OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff),
            OfficeImageFormat.Tiff
        };
        yield return new object[] {
            OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Webp),
            OfficeImageFormat.Webp
        };
    }

    private static byte[] CreateMinimalBmp() {
        byte[] bytes = new byte[58];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        WriteInt32LittleEndian(bytes, 2, bytes.Length);
        WriteInt32LittleEndian(bytes, 10, 54);
        WriteInt32LittleEndian(bytes, 14, 40);
        WriteInt32LittleEndian(bytes, 18, 1);
        WriteInt32LittleEndian(bytes, 22, 1);
        WriteUInt16LittleEndian(bytes, 26, 1);
        WriteUInt16LittleEndian(bytes, 28, 24);
        WriteInt32LittleEndian(bytes, 34, 4);
        bytes[54] = 0x33;
        bytes[55] = 0x22;
        bytes[56] = 0x11;
        return bytes;
    }

    private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

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
        for (int bit = 0; bit < 8; bit++) {
            crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }

        return crc;
    }

}
