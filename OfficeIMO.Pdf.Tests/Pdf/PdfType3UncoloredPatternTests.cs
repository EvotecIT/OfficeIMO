using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfType3UncoloredPatternTests {
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
    public void RenderPage_FailsClosedForOuterShadingPatternOnUncoloredType3Glyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 10 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_FailsClosedForImageMaskUnderOuterPatternOnUncoloredType3Glyph() {
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
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Shapes);
        Assert.Empty(drawing.Images);
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
        bool invokeThroughForm = false,
        string glyphContent = "500 0 d0 0 0 500 700 re f",
        string glyphResources = "<< >>",
        string? formContent = null,
        string formResourceEntries = "/Font << /FType3 5 0 R >>",
        string formDictionaryEntries = "",
        string catalogEntries = "",
        string pageResourceEntries = "",
        string patternResourceEntries = "/P1 7 0 R",
        IReadOnlyList<string>? extraObjects = null) {
        string pageResources = invokeThroughForm
            ? "<< /Pattern << " + patternResourceEntries + " >> /XObject << /Fm1 8 0 R >> " + pageColorSpaceResources + " " + pageResourceEntries + " >>"
            : "<< /Font << /FType3 5 0 R >> /Pattern << " + patternResourceEntries + " >> " + pageColorSpaceResources + " " + pageResourceEntries + " >>";
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R " + catalogEntries + " >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources " + pageResources + " /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources " + glyphResources + " >>\nendobj",
            StreamObject(6, "<<", glyphContent),
            StreamObject(7, patternDictionary, patternContent)
        };
        if (invokeThroughForm) {
            objects.Add(StreamObject(
                8,
                "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << " + formResourceEntries + " >> " + formDictionaryEntries,
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
