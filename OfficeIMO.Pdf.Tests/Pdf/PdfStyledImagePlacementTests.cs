using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStyledImagePlacementTests {
    [Theory]
    [InlineData("inline")]
    [InlineData("flow")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("canvas")]
    public void ImageStyleSnapshotRetainsCropRotationAndMirrorsAcrossPlacementOwners(string path) {
        byte[] source = PdfPngTestImages.CreateRgbPng(4, 2);
        var style = new PdfImageStyle { Fit = OfficeImageFit.Stretch,
            SourceCrop = new PdfImageSourceCrop(.25D, .125D, .125D, .25D),
            RotationAngle = -90D, HorizontalFlip = true, AlternativeText = "Mirrored crop" };
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions { PageWidth = 300, PageHeight = 300,
            MarginLeft = 40, MarginRight = 40, MarginTop = 90, MarginBottom = 90 });
        if (path == "inline") document.Paragraph(paragraph => paragraph.Inline(new PdfInlineImage(source, 60, 30, style, 0)));
        else if (path == "flow") document.Image(source, 60, 30, style: style);
        else if (path == "canvas") document.Canvas(canvas => canvas.Image(source, 100, 100, 60, 30, style));
        else {
            var image = new PdfHeaderFooterImage(source, 60, 30, PdfAlign.Left, style);
            document.Compose(compose => compose.Page(page => {
                if (path == "header") page.Header(header => header.Image(image));
                else page.Footer(footer => footer.Image(image));
                page.Content(content => content.Paragraph(paragraph => paragraph.Text("Body")));
            }));
        }
        style.SourceCrop = null;
        style.RotationAngle = 0;
        style.HorizontalFlip = false;
        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(document.ToBytes()));
        Assert.Equal(0, placement.A, 3);
        Assert.Equal(96, placement.B, 3);
        Assert.Equal(48, placement.C, 3);
        Assert.Equal(0, placement.D, 3);
        Assert.NotNull(placement.Clip);
    }
}
