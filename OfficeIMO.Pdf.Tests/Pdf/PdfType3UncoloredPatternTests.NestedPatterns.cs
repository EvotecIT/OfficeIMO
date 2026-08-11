using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Theory]
    [InlineData("/Pattern cs /P2 sc 0 0 500 700 re f")]
    [InlineData("/Pattern CS /P2 SC 60 w 30 30 440 640 re S")]
    public void RenderPage_FailsClosedForPatternNamesPassedToBasicColorOperators(string glyphContent) {
        byte[] pdf = BuildColoredType3NestedFormPatternPdf(glyphContent);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_AppliesInheritedPatternPaintInsideNestedType3Form() {
        byte[] pdf = BuildColoredType3NestedFormPatternPdf(
            "/Pattern cs /P2 scn /Fm1 Do");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor painted = raster.GetPixel(22, 96);
        Assert.Equal((byte)0, painted.R);
        Assert.Equal((byte)255, painted.G);
        Assert.Equal((byte)0, painted.B);
        Assert.True(painted.A > 0);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_PreservesInheritedPatternPhaseAcrossNestedType3FormTransform() {
        byte[] pdf = BuildColoredType3NestedFormPatternPdf(
            "/Pattern cs /P2 scn 1 0 0 1 250 0 cm /Fm1 Do");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(26, 96));
        OfficeColor painted = raster.GetPixel(31, 96);
        Assert.Equal((byte)0, painted.R);
        Assert.Equal((byte)255, painted.G);
        Assert.Equal((byte)0, painted.B);
        Assert.True(painted.A > 0);
    }

    [Fact]
    public void RenderPage_FailsClosedForDeferredMalformedHiddenPatternSelection() {
        const string glyphContent = "500 0 d0 /OC /Hidden BDC /Pattern cs 1 /P2 scn EMC 0 0 500 700 re f";
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", "BT /FType3 18 Tf 20 100 Td (A) Tj ET"),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P2 7 0 R >> /Properties << /Hidden 9 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", glyphContent),
            StreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
            "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
        };
        byte[] pdf = Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_ClearsInheritedShadingWhenNestedFormReselectsPatternColorSpace() {
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", "BT /FType3 18 Tf 20 100 Td (A) Tj ET"),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P2 7 0 R >> /XObject << /Fm1 8 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", "500 0 d0 /Pattern cs /P2 scn /Fm1 Do"),
            "7 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj",
            StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", "/Pattern cs 0 0 500 700 re f")
        };
        byte[] pdf = Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(raster.GetPixel(22, 96), raster.GetPixel(27, 96));
    }

    private static byte[] BuildColoredType3NestedFormPatternPdf(string glyphContent) {
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", "BT /FType3 18 Tf 20 100 Td (A) Tj ET"),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P2 7 0 R >> /XObject << /Fm1 8 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", "500 0 d0 " + glyphContent),
            StreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 250 700] /XStep 500 /YStep 700 /Resources << >>", "0 1 0 rg 0 0 250 700 re f"),
            StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", "q 0 0 250 700 re f Q 250 0 250 700 re f")
        };
        return Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
    }
}
