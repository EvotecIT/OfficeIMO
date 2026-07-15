using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEmbeddedFontRenderingTests {
    [Fact]
    public void RenderPage_RetainsSupportedEmbeddedTrueTypeProgramInDrawingScene() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) return;

        byte[] fontBytes = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("Managed Embedded", fontBytes);
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .UseFontFamily(family)
            .Paragraph(paragraph => paragraph.Text("Document font outlines"))
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg }));

        OfficeFontFace face = Assert.Single(drawing.Fonts.Faces);
        Assert.Equal("ManagedEmbedded", face.FamilyName);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.font-substitution");
        Assert.Contains("font-family=\"ManagedEmbedded\"", System.Text.Encoding.UTF8.GetString(result.Bytes!), StringComparison.Ordinal);
    }

    [Fact]
    public void RenderPage_ReportsSubstitutionForUnembeddedStandardFont() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Standard font"))
            .ToBytes();

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg }));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.font-substitution");
    }
}
