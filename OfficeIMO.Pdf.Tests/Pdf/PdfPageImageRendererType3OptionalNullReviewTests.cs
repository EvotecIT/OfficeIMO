using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_TreatsExplicitNullLuminosityBackdropAsOmitted() {
        string graphicsState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /BC null /G 8 0 R >> >>\nendobj";
        string softMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /Type /Group /S /Transparency /CS /DeviceGray >> /Resources << >>", "1 g 0 0 100 100 re f");
        byte[] pdf = BuildSingleStreamPdf("/GS1 gs 0 0 100 100 re f", "<< /ExtGState << /GS1 7 0 R >> >>", graphicsState, softMask);

        OfficeDrawingEffectGroup effect = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Elements.OfType<OfficeDrawingEffectGroup>());

        Assert.NotNull(effect.SoftMask);
        Assert.Equal(OfficeSoftMaskMode.Luminosity, effect.SoftMask!.Mode);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullSoftMaskGroupColorSpaceAsOmitted() {
        string graphicsState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string softMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /Type /Group /S /Transparency /CS null >> /Resources << >>", "0 0 100 100 re f");
        byte[] pdf = BuildSingleStreamPdf("/GS1 gs 0 0 100 100 re f", "<< /ExtGState << /GS1 7 0 R >> >>", graphicsState, softMask);

        OfficeDrawingEffectGroup effect = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Elements.OfType<OfficeDrawingEffectGroup>());

        Assert.NotNull(effect.SoftMask);
        Assert.Equal(OfficeSoftMaskMode.Alpha, effect.SoftMask!.Mode);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullType3ImageMaskDecodeAsDefault() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode null", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullType3FormMatrixAsIdentity() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Fm1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Fm1 Do");
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Matrix null /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, form);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("")]
    [InlineData("/ViewState null")]
    public void RenderPage_AppliesDefaultOnViewUsageStateToType3OptionalContent(string viewState) {
        byte[] pdf = BuildViewUsageType3Pdf(viewState);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedViewUsageStateInType3OptionalContent() {
        AssertType3FallsBackWithoutNativeShapes(BuildViewUsageType3Pdf("/ViewState /Bad"));
    }

    [Theory]
    [InlineData("/F /Fl", "rgb")]
    [InlineData("/D [0 1 0 1 0 1]", "rgb")]
    [InlineData("/DP << >>", "rgb")]
    [InlineData("/IM false", "rgb")]
    [InlineData("/I true", "rgb")]
    [InlineData("/Filter /AHx", "726762>")]
    [InlineData("/Filter [/AHx]", "726762>")]
    public void RenderPage_FailsClosedForInlineOnlyNamesOnType3ImageXObjects(string imageEntry, string imageData) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 {imageEntry}", imageData);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    private static byte[] BuildViewUsageType3Pdf(string viewState) {
        const string pageContent = "BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Properties << /Layer 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /OC /Layer BDC 0 0 500 700 re f EMC");
        string content = BuildStreamObject(4, "<<", pageContent);
        string pdfText = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /BaseState /OFF /AS [<< /Event /View /Category [/View] /OCGs [7 0 R] >>] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            content,
            type3Font,
            glyph,
            $"7 0 obj\n<< /Type /OCG /Name (View default) /Usage << /View << {viewState} >> >> >>\nendobj",
            "trailer\n<< /Root 1 0 R >>\n%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdfText);
    }
}
