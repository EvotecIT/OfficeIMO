using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
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
}
