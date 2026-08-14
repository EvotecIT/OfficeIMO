using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_TreatsExplicitNullType3FormGroupAsAbsent() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Form 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Form Do");
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group null /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, form);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3FormGroup() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Form 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Form Do");
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group 99 0 R /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, form);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullColorSpaceOnType3ImageMaskAsAbsent() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /ColorSpace null /BitsPerComponent 1", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Mask null")]
    [InlineData("/SMask null")]
    public void RenderPage_TreatsExplicitNullNestedMasksOnType3SoftMaskImageAsAbsent(string nestedMaskEntry) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /SMask 8 0 R", "rgb");
        string softMask = BuildStreamObject(8, $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 {nestedMaskEntry}", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image, softMask);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3ImageDecode() {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf("/Decode 99 0 R"));
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3DctDecodeParameters() {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /DecodeParms 99 0 R /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3PageBlendColorSpace() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        const string graphicsState = "7 0 obj\n<< /Type /ExtGState /BM /Multiply >>\nendobj";
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            "/Group << /Type /Group /S /Transparency /CS 99 0 R >>",
            type3Font,
            glyph,
            graphicsState);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullType3PageBlendColorSpaceAsOmitted() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        const string graphicsState = "7 0 obj\n<< /Type /ExtGState /BM /Multiply >>\nendobj";
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            "/Group << /Type /Group /S /Transparency /CS null >>",
            type3Font,
            glyph,
            graphicsState);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }
}
