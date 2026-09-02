using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageDocumentTests {
    [Fact]
    public void CreateFromImagesPreservesCallerOrderAndUsesImagePhysicalSize() {
        byte[] first = PdfPngTestImages.CreateRgbPng(96, 48);
        byte[] second = PdfPngTestImages.CreateRgbPng(48, 96);

        PdfDocument document = PdfDocument.CreateFromImages([
            new PdfImageDocumentSource(first, "wide.png"),
            new PdfImageDocumentSource(second, "tall.png")
        ]);

        PdfDocumentInfo info = document.Inspect();
        Assert.Equal(2, info.PageCount);
        Assert.Equal(72D, info.Pages[0].Width, 2);
        Assert.Equal(36D, info.Pages[0].Height, 2);
        Assert.Equal(36D, info.Pages[1].Width, 2);
        Assert.Equal(72D, info.Pages[1].Height, 2);
    }

    [Fact]
    public void CreateFromImagesUsesPerAxisDpiForNaturalPageAndPlacement() {
        var raster = new OfficeRasterImage(72, 72);
        byte[] image = OfficePngWriter.Encode(raster, new OfficePngEncodeOptions {
            DpiX = 72D,
            DpiY = 144D
        });

        PdfDocument document = PdfDocument.CreateFromImages([
            new PdfImageDocumentSource(image, "non-square-dpi.png")
        ]);

        PdfPageInfo page = Assert.Single(document.Inspect().Pages);
        PdfImagePlacement placement = Assert.Single(document.Images.Placements());
        Assert.InRange(page.Width, 71.99D, 72.01D);
        Assert.InRange(page.Height, 35.99D, 36.01D);
        Assert.InRange(placement.Width, 71.99D, 72.01D);
        Assert.InRange(placement.Height, 35.99D, 36.01D);
        Assert.InRange(placement.X, -0.01D, 0.01D);
    }

    [Fact]
    public void CreateFromImagesCanFitToAutoOrientedFixedPaper() {
        byte[] image = PdfPngTestImages.CreateRgbPng(96, 48);

        PdfDocument document = PdfDocument.CreateFromImages(
            [new PdfImageDocumentSource(image, "landscape.png")],
            new PdfImageDocumentOptions {
                FixedPageSize = PageSizes.A4,
                AutoOrientPage = true,
                Margin = 24D,
                Fit = OfficeImageFit.Contain
            });

        PdfDocumentInfo info = document.Inspect();
        Assert.Single(info.Pages);
        Assert.Equal(PageSizes.A4.Height, info.Pages[0].Width, 2);
        Assert.Equal(PageSizes.A4.Width, info.Pages[0].Height, 2);
    }

    [Fact]
    public void CreateFromImagesUsesPhysicalDimensionsWhenAutoOrientingFixedPaper() {
        var raster = new OfficeRasterImage(72, 72);
        byte[] image = OfficePngWriter.Encode(raster, new OfficePngEncodeOptions {
            DpiX = 72D,
            DpiY = 144D
        });

        PdfDocument document = PdfDocument.CreateFromImages(
            [new PdfImageDocumentSource(image, "non-square-dpi.png")],
            new PdfImageDocumentOptions {
                FixedPageSize = PageSizes.A4,
                AutoOrientPage = true,
                Fit = OfficeImageFit.Contain
            });

        PdfPageInfo page = Assert.Single(document.Inspect().Pages);
        Assert.Equal(PageSizes.A4.Height, page.Width, 2);
        Assert.Equal(PageSizes.A4.Width, page.Height, 2);
    }

    [Fact]
    public void CreateFromImagesAllowsDynamicPageMarginsLargerThanUnusedFallbackPaper() {
        byte[] image = PdfPngTestImages.CreateRgbPng(96, 48);

        PdfDocument document = PdfDocument.CreateFromImages(
            [new PdfImageDocumentSource(image, "large-margin.png")],
            new PdfImageDocumentOptions { Margin = 300D });

        PdfPageInfo page = Assert.Single(document.Inspect().Pages);
        Assert.Equal(672D, page.Width, 2);
        Assert.Equal(636D, page.Height, 2);
    }

    [Fact]
    public void CreateFromImagesAppliesEmbeddedJpegOrientationBeforeSizingAndDrawing() {
        var raster = new OfficeRasterImage(2, 1);
        raster.SetPixel(0, 0, OfficeColor.Red);
        raster.SetPixel(1, 0, OfficeColor.Blue);
        byte[] image = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444,
            Metadata = new OfficeJpegMetadata(exif: CreateExifOrientation(6))
        });

        PdfDocument document = PdfDocument.CreateFromImages(
            [new PdfImageDocumentSource(image, "phone-photo.jpg")],
            new PdfImageDocumentOptions {
                FixedPageSize = PageSizes.A4,
                AutoOrientPage = true,
                Fit = OfficeImageFit.Contain
            });

        PdfDocumentInfo info = document.Inspect();
        Assert.Single(info.Pages);
        Assert.Equal(PageSizes.A4.Width, info.Pages[0].Width, 2);
        Assert.Equal(PageSizes.A4.Height, info.Pages[0].Height, 2);
        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(document.ToBytes()));
        Assert.Equal(1, extracted.Width);
        Assert.Equal(2, extracted.Height);
    }

    [Fact]
    public void CreateFromImagesRejectsUnsupportedPayloadWithoutPublishingPartialDocument() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.CreateFromImages([
                new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(10, 20, 30), "valid.png"),
                new PdfImageDocumentSource([1, 2, 3, 4], "invalid.bin")
            ]));

        Assert.Contains("image", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateFromImagesHonorsCancellationWhileEnumeratingSources() {
        using var cancellation = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() =>
            PdfDocument.CreateFromImages(EnumerateSources(cancellation), options: null, cancellation.Token));

        static IEnumerable<PdfImageDocumentSource> EnumerateSources(CancellationTokenSource cancellation) {
            yield return new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(10, 20, 30), "first.png");
            cancellation.Cancel();
            yield return new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(40, 50, 60), "second.png");
        }
    }

    [Fact]
    public void CreateFromImagePathsHonorsCancellationBeforeFileIo() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            PdfDocument.CreateFromImages(
                new[] { "missing-image.png" },
                options: null,
                cancellation.Token));
    }

    private static byte[] CreateExifOrientation(ushort orientation) => [
        (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
        0x01, 0x00,
        0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00,
        (byte)orientation, (byte)(orientation >> 8), 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00
    ];
}
