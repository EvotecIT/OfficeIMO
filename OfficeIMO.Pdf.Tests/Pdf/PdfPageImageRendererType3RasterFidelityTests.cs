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
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 1 j 20 w 0 0 m 250 700 l 500 0 l S");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Contains(drawing.Elements, element => element is OfficeDrawingShape);
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
