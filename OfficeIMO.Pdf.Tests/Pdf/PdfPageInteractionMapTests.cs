using OfficeIMO.Pdf;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageInteractionMapTests {
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

        PdfPageInteractionMap map = PdfDocument.Open(source).Read.Interactions(1);

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
}
