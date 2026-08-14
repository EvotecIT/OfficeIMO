using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("")]
    [InlineData("0 j")]
    [InlineData("2 j")]
    public void RenderPage_FailsClosedForUnsupportedType3PathJoins(string lineJoin) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 " + lineJoin + " 20 w 0 0 m 250 700 l 500 0 l S");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_PreservesSupportedRoundType3PathJoin() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 1 J 1 j 20 w 0 0 m 250 700 l 500 0 l S");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Contains(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Theory]
    [InlineData("")]
    [InlineData("0 J")]
    [InlineData("2 J")]
    public void RenderPage_FailsClosedForUnsupportedOpenType3PathCaps(string lineCap) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 1 j " + lineCap + " 20 w 0 0 m 250 700 l 500 0 l S");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_PreservesDeepPatternTileImagePaintOrder() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /Fm1 8 0 R >> >>", "/Fm1 Do");
        string form1 = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Fm2 9 0 R >> >>", "/Fm2 Do");
        string form2 = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Im1 10 0 R >> >>", "1 0 0 rg 0 0 10 10 re f q 10 0 0 10 0 0 cm /Im1 Do Q 0 0 1 rg 0 0 10 10 re f");
        string image = BuildStreamObject(10, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "\0\u00ff\0");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, form1, form2, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf)).GetPixel(24, 94);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(pixel.R < 20 && pixel.G < 20 && pixel.B > 235);
    }

    [Fact]
    public void RenderPage_PreservesDeepSoftMaskImagePaintOrder() {
        string graphicsState = "5 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 6 0 R >> >>\nendobj";
        string softMask = BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Group << /Type /Group /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Fm1 7 0 R >> >>", "/Fm1 Do");
        string form1 = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /Fm2 8 0 R >> >>", "/Fm2 Do");
        string form2 = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /Im1 9 0 R >> >>", "1 0 0 rg 0 0 20 20 re f q 20 0 0 20 0 0 cm /Im1 Do Q 0 0 0 rg 0 0 20 20 re f");
        string image = BuildStreamObject(9, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "\u00ff\u00ff\u00ff");
        byte[] pdf = BuildSingleStreamPdf("/GS1 gs 1 0 0 rg 0 0 20 20 re f", "<< /ExtGState << /GS1 5 0 R >> >>", graphicsState, softMask, form1, form2, image);

        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf)).GetPixel(10, 190);

        Assert.Equal(0, pixel.A);
    }

    [Fact]
    public void RenderPage_FailsClosedForPaintedType3BezierPath() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 0 0 m 100 700 400 700 500 0 c h f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForMagnifiedType3TilingPattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 1 1] /XStep 1 /YStep 1 /Matrix [100 0 0 100 0 0] /Resources << >>", "1 0 0 rg 0 0 1 1 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }
}
