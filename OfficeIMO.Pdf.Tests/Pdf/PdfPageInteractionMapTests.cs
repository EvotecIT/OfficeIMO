using OfficeIMO.Pdf;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageInteractionMapTests {
    [Fact]
    public void InteractionKind_PreservesPublishedNumericValues() {
        Assert.Equal(0, (int) PdfInteractionKind.Text);
        Assert.Equal(1, (int) PdfInteractionKind.Link);
        Assert.Equal(2, (int) PdfInteractionKind.Annotation);
        Assert.Equal(3, (int) PdfInteractionKind.FormWidget);
        Assert.Equal(4, (int) PdfInteractionKind.Image);
    }

    [Fact]
    public void InteractionMap_ProjectsTextLinksAnnotationsAndWidgets() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Selectable text ").Link("project", "https://officeimo.net/"))
            .TextField("Person.Name", value: "Ada")
            .ToBytes();
        source = PdfAnnotationEditor.AddStampAnnotation(
            source,
            new PdfStampAnnotationOptions {
                X = 72,
                Y = 500,
                Width = 120,
                Height = 40,
                Contents = "Review stamp"
            }).Bytes;

        PdfPageInteractionMap map = PdfDocument.Load(source).Render.Interactions(1);

        Assert.Contains(map.Regions, region => region.Kind == PdfInteractionKind.Text && region.Text == "S");
        PdfPageInteractionRegion link = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.Link);
        PdfPageInteractionRegion annotation = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.Annotation && region.Subtype == "Stamp");
        PdfPageInteractionRegion widget = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.FormWidget);
        Assert.Equal("https://officeimo.net/", link.Target);
        Assert.Equal("Review stamp", annotation.Text);
        Assert.Equal("Person.Name", widget.FieldName);
        Assert.Contains(link, map.HitTest((link.Quad.Left + link.Quad.Right) / 2D, (link.Quad.Top + link.Quad.Bottom) / 2D));
        Assert.Contains(annotation, map.HitTest((annotation.Quad.Left + annotation.Quad.Right) / 2D, (annotation.Quad.Top + annotation.Quad.Bottom) / 2D));
        Assert.Contains(widget, map.HitTest((widget.Quad.Left + widget.Quad.Right) / 2D, (widget.Quad.Top + widget.Quad.Bottom) / 2D));
        Assert.Contains("Selectable text", map.GetSelectedText(0, 0, map.Width, map.Height), StringComparison.Ordinal);
    }

    [Fact]
    public void InteractionMap_AppliesCropAndPageRotationToVisualCoordinates() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Link("rotated link", "https://example.com/rotated"))
            .ToBytes();
        source = PdfPageEditor.SetCropBox(source, 0, 200, 595, 842);
        source = PdfPageEditor.RotatePages(source, 90);

        PdfPageInteractionMap map = PdfPageInteractionMap.Create(source, 1);
        PdfPageInteractionRegion link = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.Link);

        Assert.Equal(642, map.Width);
        Assert.Equal(595, map.Height);
        Assert.InRange(link.Quad.Left, 0, map.Width);
        Assert.InRange(link.Quad.Right, 0, map.Width);
        Assert.InRange(link.Quad.Top, 0, map.Height);
        Assert.InRange(link.Quad.Bottom, 0, map.Height);
        Assert.Contains(link, map.HitTest((link.Quad.Left + link.Quad.Right) / 2D, (link.Quad.Top + link.Quad.Bottom) / 2D));
    }

    [Fact]
    public void InteractionMap_ProjectsExactEditableImagePlacement() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Image interaction"))
            .ToBytes();
        PdfDocument withImage = PdfDocument.Load(source).Images.Add(
            new PdfPageRegion(1, 50D, 60D, 40D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;

        PdfPageInteractionMap map = withImage.Reader.Interactions(1);
        PdfPageInteractionRegion image = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.Image);

        Assert.NotNull(image.ImagePlacement);
        Assert.Equal(50D, image.ImagePlacement!.X, 3);
        Assert.Equal(60D, image.ImagePlacement.Y, 3);
        Assert.Equal(40D, image.ImagePlacement.Width, 3);
        Assert.Equal(20D, image.ImagePlacement.Height, 3);
        Assert.Contains(image, map.HitTest(
            (image.Quad.Left + image.Quad.Right) / 2D,
            (image.Quad.Top + image.Quad.Bottom) / 2D));

        PdfImageEditResult removed = withImage.Images.Remove(image.ImagePlacement);
        Assert.Empty(removed.Document.Images.Placements());
    }

    [Fact]
    public void InteractionMap_ProjectsImageThroughCropAndPageRotation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Rotated image interaction"))
            .ToBytes();
        byte[] withImage = PdfDocument.Load(source).Images.Add(
            new PdfPageRegion(1, 75D, 90D, 45D, 30D),
            PdfPngTestImages.CreateRgbPng(0, 128, 255)).Document.ToBytes();
        withImage = PdfPageEditor.SetCropBox(withImage, 20D, 40D, 400D, 600D);
        withImage = PdfPageEditor.RotatePages(withImage, 90);

        PdfPageInteractionMap map = PdfPageInteractionMap.Create(withImage, 1);
        PdfPageInteractionRegion image = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.Image);

        Assert.InRange(image.Quad.Left, 0D, map.Width);
        Assert.InRange(image.Quad.Right, 0D, map.Width);
        Assert.InRange(image.Quad.Top, 0D, map.Height);
        Assert.InRange(image.Quad.Bottom, 0D, map.Height);
        Assert.Contains(image, map.HitTest(
            (image.Quad.Left + image.Quad.Right) / 2D,
            (image.Quad.Top + image.Quad.Bottom) / 2D));
    }

    [Fact]
    public void InteractionMap_ClipsImageHitRegionToExactVisibleRectangle() {
        const string content = "q 50 50 50 50 re W n 100 0 0 100 0 0 cm /Im1 Do Q";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "RGB", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));

        PdfImagePlacement placement = Assert.Single(PdfDocument.Load(source).Images.Placements());
        Assert.NotNull(placement.Clip);
        Assert.True(placement.Clip!.IsRectangle);
        Assert.True(placement.Clip.IsExact);
        PdfPageInteractionMap map = PdfPageInteractionMap.Create(source, 1);
        PdfPageInteractionRegion image = Assert.Single(map.Regions, static region => region.Kind == PdfInteractionKind.Image);

        Assert.Equal(50D, image.Quad.Left, 3);
        Assert.Equal(100D, image.Quad.Right, 3);
        Assert.Equal(100D, image.Quad.Top, 3);
        Assert.Equal(150D, image.Quad.Bottom, 3);
        Assert.DoesNotContain(image, map.HitTest(25D, 125D));
        Assert.Contains(image, map.HitTest(75D, 125D));
    }

    [Theory]
    [InlineData(0, 30D, 60D, 80D, 110D)]
    [InlineData(90, 60D, 80D, 110D, 130D)]
    [InlineData(270, 10D, 30D, 60D, 80D)]
    public void InteractionMap_ClipsImageUsingSourceCoordinatesOnOffsetRotatedCrop(
        int rotation,
        double expectedLeft,
        double expectedTop,
        double expectedRight,
        double expectedBottom) {
        const string content = "q 50 50 50 50 re W n 100 0 0 100 0 0 cm /Im1 Do Q";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /CropBox [20 40 180 160] /Rotate " + rotation.ToString(CultureInfo.InvariantCulture) + " /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "RGB", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));

        PdfPageInteractionMap map = PdfPageInteractionMap.Create(source, 1);
        PdfPageInteractionRegion image = Assert.Single(map.Regions, static region => region.Kind == PdfInteractionKind.Image);

        Assert.Equal(expectedLeft, image.Quad.Left, 3);
        Assert.Equal(expectedTop, image.Quad.Top, 3);
        Assert.Equal(expectedRight, image.Quad.Right, 3);
        Assert.Equal(expectedBottom, image.Quad.Bottom, 3);
        Assert.DoesNotContain(image, map.HitTest(Math.Max(0D, expectedLeft - 10D), (expectedTop + expectedBottom) / 2D));
        Assert.Contains(image, map.HitTest((expectedLeft + expectedRight) / 2D, (expectedTop + expectedBottom) / 2D));
    }

    [Fact]
    public void InteractionMap_EnforcesTextRegionBudget() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("More than one glyph"))
            .ToBytes();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageInteractionMap.Create(source, 1, new PdfPageInteractionOptions { MaxTextRegions = 1 }));

        Assert.Equal(PdfReadLimitKind.InteractionRegions, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void InteractionMap_EnforcesImageRegionBudget() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Image budget"))
            .ToBytes();
        PdfDocument first = PdfDocument.Load(source).Images.Add(
            new PdfPageRegion(1, 20D, 20D, 20D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        byte[] second = first.Images.Add(
            new PdfPageRegion(1, 60D, 20D, 20D, 20D),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)).Document.ToBytes();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageInteractionMap.Create(second, 1, new PdfPageInteractionOptions { MaxImageRegions = 1 }));

        Assert.Equal(PdfReadLimitKind.InteractionRegions, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void InteractionMap_CountsOnlyTextRegionsThatIntersectThePage() {
        const string content = "BT /F1 12 Tf 10000 10000 Td (off-page text) Tj ET";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));

        PdfPageInteractionMap map = PdfPageInteractionMap.Create(
            source,
            1,
            new PdfPageInteractionOptions { MaxTextRegions = 1 });

        Assert.Empty(map.TextRegions);
    }

    [Fact]
    public void InteractionMap_CountsOnlyImageRegionsThatIntersectThePage() {
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(200D, 200D))).ToBytes();
        PdfDocument offPage = PdfDocument.Load(source).Images.Add(
            new PdfPageRegion(1, 10000D, 10000D, 20D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        byte[] withVisibleImage = offPage.Images.Add(
            new PdfPageRegion(1, 20D, 20D, 20D, 20D),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)).Document.ToBytes();

        PdfPageInteractionMap map = PdfPageInteractionMap.Create(
            withVisibleImage,
            1,
            new PdfPageInteractionOptions { MaxImageRegions = 1 });

        PdfPageInteractionRegion image = Assert.Single(map.Regions, region => region.Kind == PdfInteractionKind.Image);
        Assert.Equal(20D, image.ImagePlacement!.X, 3);
        Assert.Equal(20D, image.ImagePlacement.Y, 3);
    }
}
