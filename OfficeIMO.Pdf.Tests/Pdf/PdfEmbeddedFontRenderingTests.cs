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
        // Generated subsets carry a six-letter tag before the font name.
        string familyName = System.Text.RegularExpressions.Regex.Replace(face.FamilyName, "^[A-Z]{6}\\+", string.Empty);
        Assert.StartsWith("ManagedEmbedded-", familyName, StringComparison.Ordinal);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.font-substitution");
        var svg = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(result.Bytes!));
        // A PDF subset tag contains '+', so its SVG reference must be a CSS
        // string selecting the same embedded face, after XML decoding.
        string reference = "\"" + face.FamilyName + "\"";
        Assert.Contains(svg.Descendants(), element => element.Name.LocalName == "text"
            && element.Attribute("font-family")?.Value == reference);
        Assert.Contains(svg.Descendants(), element => element.Name.LocalName == "style"
            && element.Value.IndexOf("font-family:" + reference, StringComparison.Ordinal) >= 0);
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
