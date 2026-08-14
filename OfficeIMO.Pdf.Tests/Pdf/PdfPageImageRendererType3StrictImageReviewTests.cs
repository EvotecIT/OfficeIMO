using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_FailsClosedWhenViewUsageTargetsUndeclaredOcg() {
        const string pageContent = "BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Properties << /Layer 8 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /OC /Layer BDC 0 0 500 700 re f EMC");
        string content = BuildStreamObject(4, "<<", pageContent);
        string pdfText = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /BaseState /ON /AS [<< /Event /View /Category [/View] /OCGs [8 0 R] >>] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            content,
            type3Font,
            glyph,
            "7 0 obj\n<< /Type /OCG /Name (Declared) >>\nendobj",
            "8 0 obj\n<< /Type /OCG /Name (Undeclared) /Usage << /View << /ViewState /OFF >> >> >>\nendobj",
            "trailer\n<< /Root 1 0 R >>\n%%EOF"
        });

        AssertType3FallsBackWithoutNativeShapes(Encoding.ASCII.GetBytes(pdfText));
    }

    [Theory]
    [InlineData("/SMask null")]
    [InlineData("/Mask null")]
    public void RenderPage_TreatsExplicitNullType3ImageTransparencyMaskAsAbsent(string maskEntry) {
        byte[] pdf = BuildStrictType3ImagePdf(maskEntry);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Filter 0")]
    [InlineData("/Filter []")]
    [InlineData("/Filter [/FlateDecode 9 0 R]")]
    public void RenderPage_FailsClosedForMalformedType3ImageFilter(string filterEntry) {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf(filterEntry));
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3ImageOptionalContent() {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf("/OC 99 0 R"));
    }

    [Theory]
    [InlineData("/Perceptual ri", "")]
    [InlineData("", "/Intent /Perceptual")]
    public void RenderPage_FailsClosedForAuthoredType3ImageRenderingIntent(string glyphPrefix, string imageEntry) {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf(imageEntry, glyphPrefix));
    }

    private static byte[] BuildStrictType3ImagePdf(string imageEntry, string glyphPrefix = "") {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", $"500 0 d0 {glyphPrefix} q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 {imageEntry}", "rgb");
        return BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyph,
            image);
    }
}
