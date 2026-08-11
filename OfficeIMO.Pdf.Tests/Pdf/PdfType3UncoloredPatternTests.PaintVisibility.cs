using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Fact]
    public void RenderPage_IgnoresUnusedVisiblePatternSelectionInSavedGroupState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Pattern << /P2 9 0 R >> >>", "q /Pattern cs /P2 scn Q 0 0 500 700 re f"),
                StreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f")
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_IgnoresAuthoredShadingInsideEmptyEvenOddGroupClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Shading << /S 9 0 R >> >>", "0 0 500 700 re f q 100 100 100 100 re 100 100 100 100 re W* n /S sh Q"),
                "9 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj"
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_UsesLocalGroupSurfaceForNestedGlyphPaintAnalysis() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> /Font << /Nested 11 0 R >> >>", "/OC /Hidden BDC /Pattern cs /P2 scn EMC BT /Nested 10 Tf 20 20 Td (A) Tj ET"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                "11 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 12 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj",
                StreamObject(12, "<<", "500 0 d0 0 0 500 700 re f")
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_SkipsTransparentGroupDuringPaintChannelAnalysis() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> /ExtGState << /Zero << /ca 0 >> >> /XObject << /Inner 11 0 R >> >>", "/OC /Hidden BDC /Pattern cs /P2 scn EMC /Zero gs /Inner Do"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                StreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /ExtGState << /One << /ca 1 >> >> >>", "/One gs 0 0 500 700 re f")
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_RequiresOneStrokeRegionToIntersectEveryEffectiveClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> >>", "/OC /Hidden BDC /Pattern CS /P2 SCN EMC 50 0 100 100 re W n 10 w 10 10 m 10 90 l 120 10 m 120 90 l S"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f")
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    private static void AssertRendersInheritedRedPattern(byte[] pdf) {
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }
}
