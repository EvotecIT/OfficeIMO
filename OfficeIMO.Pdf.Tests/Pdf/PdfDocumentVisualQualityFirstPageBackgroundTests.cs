using System;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void BackgroundShape_FirstPageOnlyIsDrawnOnTheFirstPageOnly() {
        var shape = OfficeShape.Rectangle(500, 60);
        shape.FillColor = PdfColor.FromRgb(234, 244, 255).ToOfficeColor();
        var band = new PdfPageBackgroundShape(shape, 41, 700) { FirstPageOnly = true };

        byte[] bytes = PdfDocument.Create()
            .BackgroundShape(band)
            .H1("Title page")
            .PageBreak()
            .Paragraph(p => p.Text("Continuation page"))
            .ToBytes();

        string firstPage = string.Concat(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPage = string.Concat(GetPageContentStreams(bytes, pageNumber: 2));
        Assert.Contains("41 700", firstPage, StringComparison.Ordinal);
        Assert.DoesNotContain("41 700", secondPage, StringComparison.Ordinal);
        Assert.True(band.Clone().FirstPageOnly);
    }
}
