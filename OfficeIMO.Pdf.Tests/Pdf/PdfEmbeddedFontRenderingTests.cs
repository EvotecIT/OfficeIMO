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
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .UseFontFamily(family)
            .Paragraph(paragraph => paragraph.Text("Document font outlines"))
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Svg }));

        OfficeFontFace face = Assert.Single(drawing.Fonts.Faces);
        // Generated subsets carry a six-letter tag before the font name.
        string familyName = System.Text.RegularExpressions.Regex.Replace(face.FamilyName, "^[A-Z]{6}\\+", string.Empty);
        Assert.StartsWith("ManagedEmbedded-", familyName, StringComparison.Ordinal);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.font-substitution");
        var svg = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(result.Bytes!));
        // Proven PDF glyphs retain their contours and logical source; the PDF
        // subset is not registered as a browser font.
        var glyphProjection = Assert.Single(svg.Descendants().Where(element =>
            element.Name.LocalName == "g"
            && element.Attribute("aria-label")?.Value == "Document font outlines"));
        Assert.Equal("img", glyphProjection.Attribute("role")?.Value);
        Assert.Contains(glyphProjection.Descendants(), element => element.Name.LocalName == "path"
            && element.Attribute("d")?.Value.StartsWith("M", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void RenderPage_ReportsSubstitutionForUnembeddedStandardFont() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Standard font"))
            .ToBytes();

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Svg }));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.font-substitution");
    }
}
