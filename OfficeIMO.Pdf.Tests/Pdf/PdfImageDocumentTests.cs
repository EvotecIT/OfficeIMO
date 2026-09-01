using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
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
    public void CreateFromImagesRejectsUnsupportedPayloadWithoutPublishingPartialDocument() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.CreateFromImages([
                new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(10, 20, 30), "valid.png"),
                new PdfImageDocumentSource([1, 2, 3, 4], "invalid.bin")
            ]));

        Assert.Contains("image", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
