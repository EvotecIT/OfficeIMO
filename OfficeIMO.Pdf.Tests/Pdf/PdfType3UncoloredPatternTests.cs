using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Fact]
    public void RenderPage_UsesOuterColoredTilingPatternForUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_UsesOuterUncoloredTilingPatternTintForUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/PatternRgb cs 0 1 0 /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: "/ColorSpace << /PatternRgb [ /Pattern /DeviceRGB ] >>",
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "0 g 0 0 5 5 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_RestoresOuterTilingPatternBeforeUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn q 0 0 1 rg Q BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f");

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_PropagatesOuterTilingPatternThroughFormToUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn /Fm1 Do",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            invokeThroughForm: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RenderPage_PropagatesPatternBaseColorSpaceThroughForm(bool stroke) {
        string selectColorSpace = stroke ? "/PatternRgb CS" : "/PatternRgb cs";
        string selectPattern = stroke ? "0 1 0 /P1 SCN" : "0 1 0 /P1 scn";
        string glyphContent = stroke ? "500 0 d0 60 w 30 30 440 640 re S" : "500 0 d0 0 0 500 700 re f";
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: selectColorSpace + " /Fm1 Do",
            pageColorSpaceResources: "/ColorSpace << /PatternRgb [ /Pattern /DeviceRGB ] >>",
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "0 g 0 0 5 5 re f",
            invokeThroughForm: true,
            glyphContent: glyphContent,
            formContent: selectPattern + " BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            formResourceEntries: "/Font << /FType3 5 0 R >> /Pattern << /P1 7 0 R >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor painted = raster.GetPixel(stroke ? 21 : 22, 96);
        Assert.Equal((byte)0, painted.R);
        Assert.Equal((byte)255, painted.G);
        Assert.Equal((byte)0, painted.B);
        Assert.True(painted.A > 0);
    }

    [Fact]
    public void RenderPage_RediscoversPatternSelectedInHiddenContentBeforeType3TextArray() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/OC /Hidden BDC /Pattern cs /P1 scn EMC BT /FType3 18 Tf 20 100 Td [(A)] TJ ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            pageResourceEntries: "/Properties << /Hidden 9 0 R >>",
            extraObjects: new[] { "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj" });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RenderPage_ChargesOnlyTheInheritedPatternChannelPaintedByType3Glyph(bool stroke) {
        const string fillPatternContent = "1 0 0 rg 0 0 5 5 re f";
        string strokePatternContent = "0 0 1 rg 0 0 5 5 re f" + new string(' ', 512);
        string glyphContent = stroke
            ? "500 0 d0 60 w 30 30 440 640 re S"
            : "500 0 d0 0 0 500 700 re f";
        const string pageContent = "/OC /Hidden BDC /Pattern cs /P1 scn /Pattern CS /P2 SCN EMC BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: pageContent,
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: fillPatternContent,
            glyphContent: glyphContent,
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            pageResourceEntries: "/Properties << /Hidden 9 0 R >>",
            patternResourceEntries: "/P1 7 0 R /P2 10 0 R",
            extraObjects: new[] {
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>", strokePatternContent)
            });
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxPageContentBytes = pageContent.Length + glyphContent.Length +
                    (stroke ? strokePatternContent.Length : fillPatternContent.Length)
            }
        };

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, readOptions: readOptions));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_UsesOuterStrokeTilingPatternForUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 60 w 30 30 440 640 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor paintedStroke = raster.GetPixel(21, 96);
        Assert.Equal((byte)255, paintedStroke.R);
        Assert.Equal((byte)0, paintedStroke.G);
        Assert.Equal((byte)0, paintedStroke.B);
        Assert.True(paintedStroke.A > 0);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(24, 96));
    }

    [Fact]
    public void RenderPage_IgnoresMissingOuterStrokePatternForFillOnlyGlyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /Missing SCN 1 0 0 rg BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_IgnoresMissingOuterFillPatternForStrokeOnlyGlyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /Missing scn 0 0 1 RG BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 60 w 30 30 440 640 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor painted = raster.GetPixel(21, 96);
        Assert.Equal((byte)0, painted.R);
        Assert.Equal((byte)0, painted.G);
        Assert.Equal((byte)255, painted.B);
        Assert.True(painted.A > 0);
    }

    [Fact]
    public void RenderPage_AppliesGlyphPrimitiveOpacityToOuterPatternPaint() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Half gs 0 0 500 700 re f",
            glyphResources: "<< /ExtGState << /Half 8 0 R >> >>",
            extraObjects: new[] { "8 0 obj\n<< /Type /ExtGState /ca 0.5 >>\nendobj" });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));
        OfficeColor painted = raster.GetPixel(22, 96);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal((byte)255, painted.R);
        Assert.Equal((byte)0, painted.G);
        Assert.Equal((byte)0, painted.B);
        Assert.InRange(painted.A, (byte)120, (byte)136);
    }

    [Fact]
    public void RenderPage_PreservesOuterPatternPhaseAcrossNonIdentityForm() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn /Fm1 Do",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            invokeThroughForm: true,
            formDictionaryEntries: "/Matrix [1 0 0 1 3 0]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(24, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_PreservesOuterPatternPhaseAcrossNestedType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 BT /FNested 500 Tf (B) Tj ET",
            glyphResources: "<< /Font << /FNested 9 0 R >> >>",
            extraObjects: new[] {
                "9 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 10 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << >> >>\nendobj",
                StreamObject(10, "<<", "500 0 d0 0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_DoesNotDoubleChargeOuterPatternDuringDiagnostics() {
        const string pageContent = "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        const string glyphContent = "500 0 d0 0 0 500 700 re f";
        const string patternContent = "1 0 0 rg 0 0 5 5 re f";
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: pageContent,
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: patternContent,
            glyphContent: glyphContent);
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxPageContentBytes = pageContent.Length + glyphContent.Length + patternContent.Length
            }
        };

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, readOptions: readOptions));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMissingOuterPatternOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /Missing scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_ReportsMissingOuterPatternCarriedIntoForm() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /Missing scn /Fm1 Do",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            invokeThroughForm: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_UsesOuterAxialShadingPatternOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor left = raster.GetPixel(22, 96);
        OfficeColor right = raster.GetPixel(27, 96);
        Assert.True(left.R > left.B);
        Assert.True(right.B > right.R);
    }

    [Fact]
    public void RenderPage_UsesOuterRadialShadingPatternOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor center = raster.GetPixel(25, 94);
        OfficeColor edge = raster.GetPixel(28, 94);
        Assert.True(center.R > center.B);
        Assert.True(edge.B > center.B);
    }

    [Theory]
    [InlineData(2, "20 /Bad 100 30 100")]
    [InlineData(3, "25 /Bad 106 0 25 106 8")]
    public void RenderPage_FailsClosedForMalformedAuthoredType3ShadingCoordinates(int shadingType, string coordinates) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType " + shadingType + " /ColorSpace /DeviceRGB /Coords [" + coordinates + "] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Domain [0 /Bad] /C0 [1 0 0] /C1 [0 0 1]")]
    [InlineData("/Domain [0 1] /C0 [1 0 /Bad] /C1 [0 0 1]")]
    [InlineData("/Domain [0 1] /C0 [1 0 0] /C1 [0 0 1 0]")]
    [InlineData("/Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /Range [0 1 0 1 0 /Bad]")]
    public void RenderPage_FailsClosedForMalformedAuthoredType3ShadingFunctionArrays(string functionEntries) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 " + functionEntries + " /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForOuterAxialShadingPatternOnType3Stroke() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 60 w 30 30 440 640 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_PropagatesOuterShadingPatternThroughForm() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn /Fm1 Do",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            invokeThroughForm: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).R > raster.GetPixel(22, 96).B);
        Assert.True(raster.GetPixel(27, 96).B > raster.GetPixel(27, 96).R);
    }

    [Fact]
    public void RenderPage_DiagnosesLaterUnsupportedShadingTransformThroughSameForm() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn /Fm1 Do q 1 1 0 1 0 0 cm /Pattern cs /P1 scn /Fm1 Do Q",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            invokeThroughForm: true);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_ValidatesDeferredShadingAgainstItsSelectionTransform() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/OC /Hidden BDC /Pattern cs /P1 scn EMC q 1 1 0 1 0 0 cm /Fm1 Do Q",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            invokeThroughForm: true,
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            pageResourceEntries: "/Properties << /Hidden 9 0 R >>",
            extraObjects: new[] { "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj" });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_HonorsOuterShadingPatternMatrix() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [-1 0 0 1 50 0] /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).B > raster.GetPixel(22, 96).R);
        Assert.True(raster.GetPixel(27, 96).R > raster.GetPixel(27, 96).B);
    }

    [Theory]
    [InlineData("/N 2", "/Extend [true true]")]
    [InlineData("/N 1", "/Extend [false false]")]
    public void RenderPage_FailsClosedForInexactOuterShadingSemantics(string exponent, string extend) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] " + exponent + " >> " + extend + " >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Theory]
    [InlineData("/Domain [0 2] /C0 [1 0 0] /C1 [0 0 1] /N 1")]
    [InlineData("/Domain [0 1] /Range [0 0.5 0 1 0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1")]
    [InlineData("/Domain [0 1] /C0 [-1 0 0] /C1 [1 0 0] /N 1")]
    public void RenderPage_FailsClosedForUnmodeledShadingFunctionIntervals(string functionEntries) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 " + functionEntries + " >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Theory]
    [InlineData("/ShadingType 2 /ColorSpace /DeviceCMYK /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 0 0 1] /N 1 >> /Extend [true true]")]
    [InlineData("/ShadingType 2 /ColorSpace /DeviceRGB /BBox [22 100 28 110] /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true]")]
    [InlineData("/ShadingType 3 /ColorSpace /DeviceRGB /Coords [23 106 2 27 106 2] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true]")]
    [InlineData("/ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 1e999 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true]")]
    [InlineData("/ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 8 25 106 2] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true]")]
    [InlineData("/ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 1 25.001 106.001 1.0011] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true]")]
    public void RenderPage_FailsClosedForUnmodeledOuterShadingContracts(string shadingEntries) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << " + shadingEntries + " >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnprojectedMiterLimitInStrictType3Content() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 2 M 80 w 30 30 m 250 670 l 470 30 l S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForInexactPatternFillClipIntersection() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 0 0 m 500 0 l 0 700 l 500 700 l h W n 0 700 m 500 700 l 0 0 l 500 0 l h f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3ImageWithSoftMaskMatte() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "0 0 1 rg BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /SMask 9 0 R", "red"),
                StreamObject(9, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Matte [1 0 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForOuterAxialShadingAxisOutsideGlyphBounds() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [1000000000 1000000000 1000001000 1000001000] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_FailsClosedForShadingPatternGraphicsState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /ExtGState << /ca 0.5 >> /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForAxisSwappedNonuniformRadialShading() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [0 2 3 0 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForDashedOuterShadingStroke() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 [3 2] 0 d 60 w 30 30 440 640 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForEllipticalOuterRadialShadingStroke() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [2 0 0 1 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [12.5 106 0 12.5 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 60 w 30 30 440 640 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void VisualParser_UsesSelectionTransformForRadialShadingStroke() {
        var shading = new PdfPageShadingResource(
            25D, 106D, 0D, 25D, 106D, 8D,
            OfficeColor.Red,
            OfficeColor.Blue);
        var pattern = new PdfPageShadingPatternResource(shading, Matrix2D.Identity);
        var selection = new PdfPagePatternSelection(
            "P1",
            tint: null,
            baseColorSpace: null,
            tilingPattern: null,
            shadingPattern: pattern,
            paintTransform: Matrix2D.Identity);
        bool rejected = false;

        IReadOnlyList<PdfPageVisualPrimitive> primitives = PdfPageContentVisualParser.Parse(
            "2 0 0 1 0 0 cm 60 w 30 30 440 640 re S",
            1200D,
            800D,
            graphicsStates: null,
            colorSpaces: null,
            shadings: null,
            shadingPatterns: null,
            tilingPatterns: null,
            initialStrokeWidth: 60D,
            unsupportedShadingTransformVisitor: () => rejected = true,
            requireExactType3ShadingProjection: true,
            initialStrokePattern: selection);

        PdfPageVisualPrimitive primitive = Assert.Single(primitives);
        Assert.False(rejected);
        Assert.NotNull(primitive.StrokeRadialGradient);
    }

    [Fact]
    public void VisualParser_RejectsOverflowedAxialShadingPlacement() {
        var shading = new PdfPageShadingResource(
            1D, 1D, 2D, 2D,
            OfficeColor.Red,
            OfficeColor.Blue);

        bool supported = PdfPageContentVisualParser.IsSupportedExactShadingPlacement(
            shading,
            new Matrix2D(1e308D, 0D, 0D, 1e308D, 0D, 0D),
            0D,
            0D,
            100D,
            100D,
            100D);

        Assert.False(supported);
    }

    [Fact]
    public void PatternStrokeBoundsIncludeDefaultMiterExtent() {
        OfficePathCommand[] path = {
            OfficePathCommand.MoveTo(10D, 90D),
            OfficePathCommand.LineTo(50D, 10D),
            OfficePathCommand.LineTo(90D, 90D)
        };
        Assert.True(PdfPageVisualPrimitive.TryCreatePath(
            path,
            fillColor: null,
            fillGradient: null,
            fillRadialGradient: null,
            strokeColor: null,
            strokeGradient: null,
            strokeRadialGradient: null,
            strokeWidth: 10D,
            strokeDashStyle: OfficeStrokeDashStyle.Solid,
            strokeLineCap: OfficeStrokeLineCap.Butt,
            strokeLineJoin: null,
            fillOpacity: null,
            strokeOpacity: 1D,
            fillRule: OfficeFillRule.NonZero,
            clipPath: null,
            paintOrder: 0D,
            fillTilingPattern: null,
            strokeTilingPattern: null,
            retainPathCommands: true,
            out PdfPageVisualPrimitive primitive));

        Assert.True(PdfReadPage.TryGetTilingPatternStrokeBounds(primitive, 200D, 200D, out PdfPageClipPath bounds));
        Assert.InRange(bounds.X, 0D, 0.001D);
        Assert.InRange(bounds.Width, 139.999D, 140.001D);
    }

    [Fact]
    public void NarrowPolygonClipOverlapDoesNotProveImageInvisible() {
        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(9.9999D, 80D),
            OfficePathCommand.LineTo(10.0001D, 80D),
            OfficePathCommand.LineTo(10D, 90D),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        var placement = new PdfImagePlacement(
            1, "Im1", 1, 0,
            10D, 0D, 0D, 10D, 0D, 10D,
            0D, 10D, 10D, 10D,
            clipPath: clip);

        Assert.False(PdfReadPage.IsInvisibleImagePlacement(placement, 100D, 100D, 100D));
    }

    [Fact]
    public void RenderPage_FailsClosedForShearedOuterRadialShadingPattern() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [1 1 0 1 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_FailsClosedForShearedOrdinaryRadialShadingPattern() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn 20 80 100 80 re f",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [1 1 0 1 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [70 120 0 70 120 40] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_RestoresOuterShadingPatternAcrossGraphicsState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn q 0 1 0 rg Q BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).R > raster.GetPixel(22, 96).B);
        Assert.True(raster.GetPixel(27, 96).B > raster.GetPixel(27, 96).R);
    }

    [Fact]
    public void RenderPage_RediscoversHiddenOuterShadingPatternBeforeTextArray() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/OC /Hidden BDC /Pattern cs /P1 scn EMC BT /FType3 18 Tf 20 100 Td [(A)] TJ ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            pageResourceEntries: "/Properties << /Hidden 9 0 R >>",
            extraObjects: new[] { "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj" });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).R > raster.GetPixel(22, 96).B);
        Assert.True(raster.GetPixel(27, 96).B > raster.GetPixel(27, 96).R);
    }

    [Fact]
    public void RenderPage_UsesFormLocalShadingPatternWithSameResourceName() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Fm1 Do",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            invokeThroughForm: true,
            formContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            formResourceEntries: "/Font << /FType3 5 0 R >> /Pattern << /P1 9 0 R >>",
            extraObjects: new[] {
                "9 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [1 0 0] /N 1 >> /Extend [true true] >> >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).B > raster.GetPixel(22, 96).R);
        Assert.True(raster.GetPixel(27, 96).R > raster.GetPixel(27, 96).B);
    }

    [Fact]
    public void RenderPage_UsesOuterRadialShadingPatternForImageMaskOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(25, 94).R > raster.GetPixel(25, 94).B);
        Assert.True(raster.GetPixel(28, 94).B > raster.GetPixel(25, 94).B);
    }

    [Fact]
    public void RenderPage_UsesOuterTilingPatternForImageMaskOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_CropsRepeatedPatternedImageMaskGroupsToStencilBounds() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 1000 Td (A) Tj ET BT /FType3 18 Tf 1500 1000 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            pageWidth: 2000D,
            pageHeight: 2000D,
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup[] groups = drawing.Elements.OfType<OfficeDrawingEffectGroup>().ToArray();

        Assert.Equal(2, groups.Length);
        Assert.All(groups, group => {
            Assert.InRange(group.Drawing.Width, 0.1D, 20D);
            Assert.InRange(group.Drawing.Height, 0.1D, 20D);
        });
        Assert.True(groups[1].Transform.OffsetX - groups[0].Transform.OffsetX > 1000D);
    }

    [Theory]
    [InlineData("q 500 0 0 700 10000 0 cm /Im1 Do Q")]
    [InlineData("q 0 0 0 0 re W n 500 0 0 700 0 0 cm /Im1 Do Q")]
    public void RenderPage_SkipsInvisiblePatternedImageMaskWithoutRejectingVisibleGlyphContent(string invisibleImageContent) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 0 0 250 700 re f " + invisibleImageContent,
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_KeepsBlendedPatternedImageMaskGroupsCropped() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 1000 Td (A) Tj ET BT /FType3 18 Tf 1500 1000 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Multiply gs q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /ExtGState << /Multiply << /BM /Multiply >> >> /XObject << /Im1 8 0 R >> >>",
            pageWidth: 2000D,
            pageHeight: 2000D,
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup[] groups = drawing.Elements.OfType<OfficeDrawingEffectGroup>().ToArray();

        Assert.Equal(2, groups.Length);
        Assert.All(groups, group => {
            Assert.InRange(group.Drawing.Width, 0.1D, 20D);
            Assert.InRange(group.Drawing.Height, 0.1D, 20D);
            Assert.Equal(OfficeBlendMode.Multiply, group.BlendMode);
        });
        Assert.True(groups[1].Transform.OffsetX - groups[0].Transform.OffsetX > 1000D);
    }

    [Fact]
    public void RenderPage_KeepsSoftMaskedPatternedImageMaskGroupsCropped() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 1000 Td (A) Tj ET BT /FType3 18 Tf 1500 1000 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Masked gs q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /ExtGState << /Masked << /SMask << /S /Alpha /G 9 0 R >> >> >> /XObject << /Im1 8 0 R >> >>",
            pageWidth: 2000D,
            pageHeight: 2000D,
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 2000 2000] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "1 g 0 0 2000 2000 re f")
            });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup[] groups = drawing.Elements.OfType<OfficeDrawingEffectGroup>().ToArray();

        Assert.Equal(2, groups.Length);
        Assert.All(groups, group => {
            Assert.InRange(group.Drawing.Width, 0.1D, 20D);
            Assert.InRange(group.Drawing.Height, 0.1D, 20D);
            OfficeDrawingSoftMask softMask = Assert.IsType<OfficeDrawingSoftMask>(group.SoftMask);
            Assert.InRange(softMask.Drawing.Width, 0.1D, 20D);
            Assert.InRange(softMask.Drawing.Height, 0.1D, 20D);
        });
        Assert.True(groups[1].Transform.OffsetX - groups[0].Transform.OffsetX > 1000D);
    }

    [Fact]
    public void RenderPage_FailsClosedForNonIsolatedType3SoftMaskGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Masked gs 0 0 500 700 re f",
            glyphResources: "<< /ExtGState << /Masked << /SMask << /S /Alpha /G 9 0 R >> >> >> >>",
            extraObjects: new[] {
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /CS /DeviceRGB >> /Resources << >>", "1 g 0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3LuminositySoftMaskWithoutExplicitGroupColorSpace() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Masked gs 0 0 500 700 re f",
            glyphResources: "<< /ExtGState << /Masked << /SMask << /S /Luminosity /G 9 0 R >> >> >> >>",
            extraObjects: new[] {
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << >>", "1 g 0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_AppliesIndirectType3SoftMaskGroupMatrix() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Masked gs 0 0 500 700 re f",
            glyphResources: "<< /ExtGState << /Masked << /SMask << /S /Alpha /G 9 0 R >> >> >> >>",
            extraObjects: new[] {
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Matrix 10 0 R /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "1 g 0 0 250 700 re f"),
                "10 0 obj\n[11 0 R 0 0 1 250 0]\nendobj",
                "11 0 obj\n1\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(27, 96));
    }

    [Theory]
    [InlineData("/BBox [0 0 0 700]", "/S /Alpha")]
    [InlineData("/BBox [0 0 500 700] /Matrix [1 0 0 1 0 0 7]", "/S /Alpha")]
    [InlineData("/BBox [0 0 500 700] /OC 10 0 R", "/S /Alpha")]
    [InlineData("/BBox [0 0 500 700]", "/S /Luminosity /BC [0 /Bad 0]")]
    public void RenderPage_FailsClosedForMalformedType3SoftMaskGroupSemantics(string formEntries, string maskEntries) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Masked gs 0 0 500 700 re f",
            glyphResources: "<< /ExtGState << /Masked << /SMask << " + maskEntries + " /G 9 0 R >> >> >> >>",
            extraObjects: new[] {
                StreamObject(9, "<< /Type /XObject /Subtype /Form " + formEntries + " /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "1 g 0 0 500 700 re f"),
                "10 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_DoesNotConsumeInheritedPatternForColoredType3Shading() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 20 20] /XStep 20 /YStep 20 /Resources << /Font << /FBase 9 0 R >> >>",
            patternContent: "0 g BT /FBase 8 Tf 1 0 0 1 2 12 Tm (X) Tj ET",
            glyphContent: "500 0 d0 /Shade sh",
            glyphResources: "<< /Shading << /Shade << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >> >>",
            type3PaintType: 1,
            extraObjects: new[] { "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj" });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_DoesNotConsumeInheritedPatternForColoredType3InlineImage() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /Missing scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm BI /W 1 /H 1 /BPC 8 /CS /RGB ID abc EI Q",
            patternResourceEntries: string.Empty,
            type3PaintType: 1);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForShearedImageInsideStrictType3PatternTile() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "1 0 0 rg BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << /XObject << /Im 8 0 R >> >>",
            patternContent: "q 5 1 0 5 0 0 cm /Im Do Q",
            glyphContent: "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f",
            glyphResources: "<< /Pattern << /P1 7 0 R >> >>",
            type3PaintType: 1,
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "abc")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_AlignsLocalizedPatternMaskSoftMaskInPageCoordinates() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Masked gs q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /ExtGState << /Masked << /SMask << /S /Alpha /G 9 0 R >> >> >> /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "1 g 0 350 500 350 re f")
            });

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 90));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(22, 97));
    }

    [Fact]
    public void RenderPage_UsesOuterAxialShadingPatternForImageMaskOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).R > raster.GetPixel(22, 96).B);
        Assert.True(raster.GetPixel(27, 96).B > raster.GetPixel(27, 96).R);
    }

    [Fact]
    public void RenderPage_FailsClosedForShearedOuterRadialShadingOnImageMaskGlyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [1 1 0 1 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPage_DoesNotBlameValidShadingForUnsupportedImageMaskProjection() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
    }

    [Fact]
    public void RenderPage_DoesNotDiagnoseUnsupportedShadingForInvisibleImageMask() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P2 SCN /Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Matrix [1 1 0 1 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [25 106 0 25 106 8] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 60 w 30 30 440 640 re S q 500 0 0 700 100000 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            patternResourceEntries: "/P1 7 0 R /P2 9 0 R",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x"),
                StreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 0 1 rg 0 0 5 5 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_DiagnosesUnsupportedPatternedImageMaskProjection() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPage_ClipsShearedPatternedImageMaskToRectangularGlyphClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 0 0 250 700 re W n q 0 700 -500 0 500 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
    }

    [Fact]
    public void RenderPage_FailsClosedForNonRenderablePatternedImageMaskPayload() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 2147483647 /Height 2147483647 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPages_TransformsAnnotationAppearanceBeforePatternMaskDiagnostics() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: string.Empty,
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            pageDictionaryEntries: "/Annots [9 0 R]",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x"),
                "9 0 obj\n<< /Type /Annot /Subtype /Stamp /Rect [20 20 80 80] /AP << /N 10 0 R >> >>\nendobj",
                StreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [300 300 360 360] /Resources << /Font << /FType3 5 0 R >> /Pattern << /P1 7 0 R >> >>", "/Pattern cs /P1 scn BT /FType3 18 Tf 300 320 Td (A) Tj ET")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPages_AppliesFormBoundingBoxBeforePatternMaskDiagnostics() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn /Fm1 Do",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            invokeThroughForm: true,
            glyphContent: "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 9 0 R >> >>",
            formContent: "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            formBoundingBox: "[0 0 10 10]",
            extraObjects: new[] {
                StreamObject(9, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPages_PreservesDistinctClipsForRepeatedFormDiagnostics() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "q 200 0 40 40 re W n /Fm1 Do Q q 0 0 240 200 re W n /Fm1 Do Q",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            invokeThroughForm: true,
            glyphContent: "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 9 0 R >> >>",
            formContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            formResourceEntries: "/Font << /FType3 5 0 R >> /Pattern << /P1 7 0 R >>",
            extraObjects: new[] {
                StreamObject(9, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_DiagnosesVisibleUnsupportedPatternedImageMaskWithOffsetCropBox() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 320 150 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            pageWidth: 440D,
            pageHeight: 300D,
            pageDictionaryEntries: "/CropBox [100 50 340 250]",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnsupportedOuterShadingOnImageMaskGlyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 2 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
        Assert.Empty(drawing.Images);
    }

    [Fact]
    public void RenderPage_UsesOuterTilingPatternForImageMaskInsideGlyphForm() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "/FmGlyph Do",
            glyphResources: "<< /XObject << /FmGlyph 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /XObject << /Im1 8 0 R >> >>", "q 500 0 0 700 0 0 cm /Im1 Do Q")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_AppliesImageMaskPatternOpacityOnce() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/GS1 gs /Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            pageResourceEntries: "/ExtGState << /GS1 << /ca 0.5 >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeColor painted = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf)).GetPixel(22, 96);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.InRange(painted.A, (byte)120, (byte)136);
        Assert.Equal((byte)255, painted.R);
    }

    [Fact]
    public void RenderPage_IgnoresOuterStrokePatternForFillOnlyImageMaskGlyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN 0 0 1 rg BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_AllowsOrdinaryTextInsideOuterTilingPattern() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/PatternRgb cs 0 1 0 /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: "/ColorSpace << /PatternRgb [ /Pattern /DeviceRGB ] >>",
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 20 20] /XStep 20 /YStep 20 /Resources << /Font << /FBase 9 0 R >> >>",
            patternContent: "0 g BT /FBase 8 Tf 1 0 0 1 2 12 Tm (X) Tj ET",
            extraObjects: new[] { "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj" });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.NotEmpty(drawing.Elements);
    }

    private static byte[] BuildUncoloredType3PatternPdf(
        string pageContent,
        string pageColorSpaceResources,
        string patternDictionary,
        string patternContent,
        bool patternIsStream = true,
        bool invokeThroughForm = false,
        string glyphContent = "500 0 d0 0 0 500 700 re f",
        string glyphResources = "<< >>",
        string? formContent = null,
        string formResourceEntries = "/Font << /FType3 5 0 R >>",
        string formDictionaryEntries = "",
        string? formBoundingBox = null,
        string catalogEntries = "",
        string pageResourceEntries = "",
        string pageDictionaryEntries = "",
        string patternResourceEntries = "/P1 7 0 R",
        int type3PaintType = 2,
        double pageWidth = 240D,
        double pageHeight = 200D,
        IReadOnlyList<string>? extraObjects = null) {
        string pageWidthText = pageWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string pageHeightText = pageHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string pageResources = invokeThroughForm
            ? "<< /Pattern << " + patternResourceEntries + " >> /XObject << /Fm1 8 0 R >> " + pageColorSpaceResources + " " + pageResourceEntries + " >>"
            : "<< /Font << /FType3 5 0 R >> /Pattern << " + patternResourceEntries + " >> " + pageColorSpaceResources + " " + pageResourceEntries + " >>";
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R " + catalogEntries + " >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 " + pageWidthText + " " + pageHeightText + "] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources " + pageResources + " /Contents 4 0 R " + pageDictionaryEntries + " >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType " + type3PaintType.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources " + glyphResources + " >>\nendobj",
            StreamObject(6, "<<", glyphContent),
            patternIsStream
                ? StreamObject(7, patternDictionary, patternContent)
                : "7 0 obj\n" + patternDictionary + " >>\nendobj"
        };
        if (invokeThroughForm) {
            string formBox = formBoundingBox ?? "[0 0 " + pageWidthText + " " + pageHeightText + "]";
            objects.Add(StreamObject(
                8,
                "<< /Type /XObject /Subtype /Form /BBox " + formBox + " /Resources << " + formResourceEntries + " >> " + formDictionaryEntries,
                formContent ?? "BT /FType3 18 Tf 20 100 Td (A) Tj ET"));
        }
        if (extraObjects != null) objects.AddRange(extraObjects);
        return Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
    }

    private static string StreamObject(int number, string dictionaryPrefix, string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        return number.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" +
               dictionaryPrefix + " /Length " + length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
               " >>\nstream\n" + content + "\nendstream\nendobj";
    }
}
