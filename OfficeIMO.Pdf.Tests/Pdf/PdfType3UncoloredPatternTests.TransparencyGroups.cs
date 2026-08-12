using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Theory]
    [InlineData("/Matrix [1 0 0 1 0]")]
    [InlineData("/Matrix [1 0 0 1 0 1e999]")]
    public void RenderPage_FailsClosedForMalformedTransparencyGroupMatrix(string matrixEntry) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] " + matrixEntry + " /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_DiagnosticsRejectOptionalContentOnNestedType3Form() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Form Do",
            glyphResources: "<< /XObject << /Form 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /OC << /Type /OCG /Name (Conditional) >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/LW 1e999")]
    [InlineData("/D [[8 3] 0]")]
    [InlineData("/D [[3 1] 2]")]
    public void RenderPage_FailsClosedForInexactType3ExtGStateStrokeValues(string graphicsStateEntry) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Bad gs 10 w 0 0 m 500 700 l S",
            glyphResources: "<< /ExtGState << /Bad << " + graphicsStateEntry + " >> >> >>");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("[3 1] 2 d")]
    [InlineData("[8 3] 0 d")]
    public void RenderPage_FailsClosedForInexactDirectType3DashStyle(string dashOperator) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 " + dashOperator + " 20 w 0 0 m 500 700 l S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenDeferredDashIsConsumedByNestedType3Text() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 [3 1] 2 d BT /Inner 500 Tf (A) Tj ET",
            glyphResources: "<< /Font << /Inner 8 0 R >> >>",
            extraObjects: new[] {
                "8 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 9 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj",
                StreamObject(9, "<<", "500 0 d0 20 w 0 0 m 500 700 l S")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Bad g")]
    [InlineData("/Bad G")]
    [InlineData("0 0 /Bad rg")]
    [InlineData("0 0 /Bad RG")]
    [InlineData("0 0 0 /Bad k")]
    [InlineData("0 0 0 /Bad K")]
    public void RenderPage_FailsClosedForMalformedDeviceColorOperands(string colorOperator) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 " + colorOperator + " 0 0 500 700 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedStrictTransformOperands() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Bad 0 0 1 0 0 cm 0 0 500 700 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForInexactTransparencyGroupPrimitiveClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(
                    8,
                    "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>",
                    "0 0 m 500 0 l 250 300 l 500 700 l 0 700 l h W n " +
                    "0 0 m 500 0 l 250 300 l 500 700 l 0 700 l h W n " +
                    "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForInexactDirectType3PrimitiveClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 " + NonConvexSuccessiveClipContent + " 0 0 500 700 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForInexactDirectType3ImageClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 " + NonConvexSuccessiveClipContent + " q 500 0 0 700 0 0 cm /Im Do Q",
            glyphResources: "<< /XObject << /Im 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Bad 10 m")]
    [InlineData("0 0 10 /Bad re")]
    public void RenderPage_FailsClosedForMalformedStrictPathOperands(string pathOperator) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 " + pathOperator + " 0 0 500 700 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    private const string NonConvexSuccessiveClipContent =
        "0 0 m 500 0 l 250 300 l 500 700 l 0 700 l h W n " +
        "0 0 m 500 0 l 250 300 l 500 700 l 0 700 l h W n";

    [Fact]
    public void RenderPage_UsesOuterTilingPatternThroughNestedTransparencyGroups() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Outer Do",
            glyphResources: "<< /XObject << /Outer 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Inner 9 0 R >> >>", "/Inner Do"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_DoesNotDoubleChargeNestedTransparencyGroupEffects() {
        const string pageContent = "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        const string patternContent = "1 0 0 rg 0 0 5 5 re f";
        const string glyphContent = "500 0 d0 /Outer Do";
        const string outerContent = "/Inner Do";
        const string innerContent = "0 0 500 700 re f";
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: pageContent,
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: patternContent,
            glyphContent: glyphContent,
            glyphResources: "<< /XObject << /Outer 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Inner 9 0 R >> >>", outerContent),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", innerContent)
            });
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxPageContentBytes = pageContent.Length + patternContent.Length + glyphContent.Length + outerContent.Length + innerContent.Length
            }
        };

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, readOptions: readOptions));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_UsesOuterAxialShadingPatternThroughTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(raster.GetPixel(22, 96).R > raster.GetPixel(22, 96).B);
        Assert.True(raster.GetPixel(27, 96).B > raster.GetPixel(27, 96).R);
    }

    [Fact]
    public void RenderPage_UsesOuterTilingStrokePatternThroughTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "40 w 20 20 460 660 re S")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingEffectGroup patternedStroke = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.Contains(patternedStroke.Drawing.Elements, element => element is OfficeDrawingTilingPattern);
    }

    [Fact]
    public void RenderPage_IgnoresUnusedOuterStrokePatternThroughTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern CS /P1 SCN 0 0 1 rg BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeColor pixel = raster.GetPixel(24, 94);
        Assert.True(pixel.R < 20 && pixel.G < 20 && pixel.B > 235);
    }

    [Fact]
    public void RenderPage_UsesOuterTilingPatternForImageMaskInsideTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 8 0 R >> >>", "q 500 0 0 700 0 0 cm /Im1 Do Q")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_FailsClosedForColorImageInsideUncoloredTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "\0\u00ff\0"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 8 0 R >> >>", "q 500 0 0 700 0 0 cm /Im1 Do Q")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Images);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedImageMaskInsideType3TransparencyGroup() {
        byte[] jpeg = CreateType3TestJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 10 /YStep 10 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /SMask 10 0 R", marker),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 8 0 R >> >>", "q 500 0 0 700 0 0 cm /Im1 Do Q"),
                StreamObject(10, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", "x")
            });
        ReplaceType3AsciiPayload(pdf, marker, jpeg);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_IgnoresColorImageHiddenByTransparentSoftMaskInsideGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "\0\u00ff\0"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /ExtGState << /Mask 10 0 R >> /XObject << /Im1 8 0 R >> >>", "q /Mask gs 500 0 0 700 0 0 cm /Im1 Do Q 0 0 250 700 re f"),
                "10 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 11 0 R >> >>\nendobj",
                StreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", string.Empty)
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_AppliesCallerOpacityOnceToTransparencyGroupComposite() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Half gs /Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            pageResourceEntries: "/ExtGState << /Half << /ca 0.5 >> >>",
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 350 700 re f 150 0 350 700 re f")
            });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        OfficeDrawingEffectGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingEffectGroup opacityGroup = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeColor singlePaint = raster.GetPixel(21, 96);
        OfficeColor overlap = raster.GetPixel(24, 96);

        Assert.Equal(0.5D, opacityGroup.Opacity, 6);
        Assert.InRange(singlePaint.A, (byte)126, (byte)129);
        Assert.InRange(Math.Abs(overlap.A - singlePaint.A), 0, 1);
    }

    [Fact]
    public void RenderPage_SkipsFullyTransparentUnsupportedTransparencyGroupContent() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Zero gs /Group Do /Normal gs 0 0 250 700 re f",
            glyphResources: "<< /ExtGState << /Zero << /ca 0 >> /Normal << /ca 1 >> >> /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Shading << /S 9 0 R >> >>", "/S sh"),
                "9 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_SkipsPatternPaintHiddenByBlackLuminositySoftMask() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /ExtGState << /Mask << /SMask << /S /Luminosity /G 10 0 R >> >> >> /Pattern << /P2 9 0 R >> >>", "q /Pattern cs /P2 scn /Mask gs 0 0 500 700 re f Q 0 0 250 700 re f"),
                StreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                StreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 g 0 0 500 700 re f")
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_UsesVisibleLocalSurfaceForSmallTransparencyGroupOnLargePage() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            pageWidth: 2000D,
            pageHeight: 2000D,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());

        Assert.InRange(group.Drawing.Width, 8.9D, 9.1D);
        Assert.InRange(group.Drawing.Height, 12.5D, 12.7D);
    }

    [Fact]
    public void RenderPage_ClipsShearedTransparencyGroupToExactFormBounds() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Matrix [1 0.5 0 1 0 0] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "-500 -500 1500 1700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(28, 90));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(28, 98));
    }

    [Fact]
    public void RenderPage_SkipsInvisibleTransparencyGroupWithoutRejectingVisibleGlyphContent() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Matrix [1 0 0 1 10000 0] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_SkipsDegenerateTransparencyGroupWithoutRejectingVisibleGlyphContent() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 0 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_FailsClosedForZeroAdvanceOrdinaryTextInsideTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /F1 9 0 R >> >>", "BT /F1 100 Tf 100 100 Td (A) Tj ET 0 0 250 700 re f"),
                "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FirstChar 65 /LastChar 65 /Widths [0] >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForInexactSuccessiveClipIntersectionAroundTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 0 0 m 500 0 l 500 300 l 250 300 l 250 700 l 0 700 l h W n 0 0 m 250 0 l 250 400 l 500 400 l 500 700 l 0 700 l h W n /Group Do Q",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_SkipsTransparencyGroupOutsideActiveClipWithoutRejectingGlyph() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 0 0 250 700 re W n /Group Do Q 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [300 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "300 0 200 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_PreservesConcaveClipAroundShearedTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 0 0 m 500 0 l 500 300 l 250 300 l 250 700 l 0 700 l h W n /Group Do Q",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Matrix [1 0.2 0 1 0 0] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(21, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 90));
    }

    [Fact]
    public void RenderPage_SkipsClippedColorImageInsideUncoloredTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "\0\u00ff\0"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 8 0 R >> >>", "0 0 250 700 re f q 0 0 250 700 re W n 100 0 0 100 500 0 cm /Im1 Do Q")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_SkipsColorImageWhollyInsideEvenOddClipHole() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "\0\u00ff\0"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 8 0 R >> >>", "q 0 0 500 700 re 150 150 200 400 re W* n q 100 0 0 100 200 250 cm /Im1 Do Q Q 0 0 150 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(21, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_LocalizesOuterAxialShadingPatternForOffsetTransparencyGroupDiagnostics() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [22.5 107 27 107] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false,
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [100 100 400 600] /Matrix [1 0 0 1 25 50] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "100 100 300 500 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Contains(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPage_DoesNotDecodeInvisibleImageInsideTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 9 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 2147483647 /Height 2147483647 /ColorSpace /DeviceRGB /BitsPerComponent 8", "x"),
                StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /XObject << /Im1 8 0 R >> >>", "0 0 250 700 re f q 1 0 0 1 10000 0 cm /Im1 Do Q")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_FailsClosedForAuthoredPatternOperatorInsideUncoloredTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Pattern << /P1 7 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
        Assert.Empty(drawing.Shapes);
    }

    [Fact]
    public void RenderPage_IgnoresOrdinaryGroupTextOutsideTheEffectiveClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /F1 9 0 R >> >>", "BT /F1 100 Tf 1000 1000 Td (X) Tj ET 0 0 250 700 re f"),
                "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj"
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_IgnoresAuthoredPatternPaintHiddenByTransparentSoftMask() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /ExtGState << /Mask 9 0 R >> /Pattern << /P2 10 0 R >> >>", "q /Pattern cs /P2 scn /Mask gs 0 0 500 700 re f Q 0 0 250 700 re f"),
                "9 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 11 0 R >> >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                StreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", string.Empty)
            });

        AssertRendersInheritedRedPattern(pdf);
    }

    [Fact]
    public void RenderPage_IgnoresAuthoredPatternSelectionInsideHiddenSavedGroupState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P1 7 0 R >> >>", "/OC /Hidden BDC q /Pattern cs /P1 scn Q EMC 0 0 500 700 re f"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_DoesNotConsumeDeferredFillPatternForEmptyFormInvocation() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> /XObject << /Empty 11 0 R >> >>", "q /OC /Hidden BDC /Pattern cs /P2 scn EMC /Empty Do Q"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                StreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", string.Empty)
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_DoesNotConsumeDeferredFillPatternForEmptyPathPaint() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> >>", "q /OC /Hidden BDC /Pattern cs /P2 scn EMC f Q"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_DoesNotConsumeDeferredFillPatternForClippedImageMask() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> /XObject << /Mask 11 0 R >> >>", "q /OC /Hidden BDC /Pattern cs /P2 scn EMC 1 0 0 1 1000 1000 cm /Mask Do Q"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                StreamObject(11, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_DoesNotConsumeDeferredFillPatternForPaintOutsideFormBounds() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> /XObject << /Outside 11 0 R >> >>", "q /OC /Hidden BDC /Pattern cs /P2 scn EMC /Outside Do Q"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f"),
                StreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >>", "20 20 10 10 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_DoesNotConsumeDeferredStrokePatternForTransparentFormStroke() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /ExtGState << /ZeroStroke << /CA 0 >> >> /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> >>", "/OC /Hidden BDC /Pattern CS /P2 SCN EMC /ZeroStroke gs 40 w 0 0 500 700 re B"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Theory]
    [InlineData("/ZeroFill gs 0 0 500 700 re f")]
    [InlineData("/ZeroFill gs BI /W 1 /H 1 /IM true /BPC 1 ID x EI")]
    public void RenderPage_DoesNotConsumeDeferredFillPatternForTransparentPaint(string transparentPaint) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do 0 0 250 700 re f",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /ExtGState << /ZeroFill << /ca 0 >> >> /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> >>", "/OC /Hidden BDC /Pattern cs /P2 scn EMC " + transparentPaint),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_FailsClosedWhenHiddenPatternSelectionEscapesIntoVisibleGroupPaint() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            catalogEntries: "/OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Properties << /Hidden 9 0 R >> /Pattern << /P2 10 0 R >> >>", "/OC /Hidden BDC /Pattern cs /P2 scn EMC 0 0 500 700 re f"),
                "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj",
                StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPage_FailsClosedForAuthoredShadingInsideUncoloredTransparencyGroup() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Shading << /S 9 0 R >> >>", "/S sh"),
                "9 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
    }

    [Fact]
    public void RenderPage_IgnoresAuthoredShadingInsideEmptyGroupClip() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Shading << /S 9 0 R >> >>", "0 0 500 700 re f q 10000 10000 10 10 re W n /S sh Q"),
                "9 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj"
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.NotEmpty(group.Drawing.Elements);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
    }

    [Fact]
    public void RenderPage_DoesNotDecodeInvisibleTransparencyGroupStream() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 /Group Do",
            glyphResources: "<< /XObject << /Group 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [100000 100000 100500 100700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", new string(' ', 128) + "0 0 500 700 re f")
            });
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = 64 }
        });

        OfficeDrawing drawing = document.Pages[0].ToDrawing();

        Assert.Empty(drawing.Elements);
    }

    [Fact]
    public void RenderPage_DoesNotValidateUnsupportedInvisibleType3Image() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 1 0 0 1 100000 100000 cm /Im Do Q",
            glyphResources: "<< /XObject << /Im 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /Unsupported", "abc")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    private static byte[] CreateType3TestJpeg(int width, int height) => new byte[] {
        0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08,
        (byte)(height >> 8), (byte)height,
        (byte)(width >> 8), (byte)width,
        0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00,
        0xFF, 0xD9
    };

    private static void ReplaceType3AsciiPayload(byte[] pdf, string marker, byte[] replacement) {
        byte[] markerBytes = Encoding.ASCII.GetBytes(marker);
        Assert.Equal(markerBytes.Length, replacement.Length);
        for (int offset = 0; offset <= pdf.Length - markerBytes.Length; offset++) {
            bool matches = true;
            for (int index = 0; index < markerBytes.Length; index++) {
                if (pdf[offset + index] == markerBytes[index]) continue;
                matches = false;
                break;
            }
            if (!matches) continue;
            Buffer.BlockCopy(replacement, 0, pdf, offset, replacement.Length);
            return;
        }
        throw new InvalidOperationException("Marker payload was not found.");
    }
}
