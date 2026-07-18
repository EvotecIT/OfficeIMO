using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_ProjectsGeneratedPdfContentToSharedDrawingAndImages() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(p => p.Text("Managed raster page"))
            .Image(PdfPngTestImages.CreateRgbPng(255, 0, 0), 24, 18, alternativeText: "Proof pixel")
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        Assert.True(drawing.Width > 0D);
        Assert.True(drawing.Height > 0D);
        string drawingText = string.Concat(drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text));
        Assert.Contains("Managed raster page", drawingText, StringComparison.Ordinal);
        Assert.Contains(drawing.Images, image => image.ContentType == "image/png" && image.Bytes.Length > 8);
        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("Managed", svgText, StringComparison.Ordinal);
        Assert.Contains("raster", svgText, StringComparison.Ordinal);
        Assert.Contains("page", svgText, StringComparison.Ordinal);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_ProjectsDeviceCmykImageXObjectAsPng() {
        byte[] cmykPixels = {
            255, 0, 0, 0,
            0, 255, 0, 0
        };
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(CompressWithDeflate(cmykPixels));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        AssertPngSignature(image.Bytes);
        Assert.Equal(new byte[] { 0, 0, 255, 255, 255, 0, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ProjectsDeviceCmykImageXObjectSoftMaskAsPngAlpha() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 0, 255, 0 }),
            CompressWithDeflate(new byte[] { 128 }));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 255, 0, 128 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesImageXObjectSoftMaskDecodeArrayAsPngAlpha() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 255, 0, 0 }),
            CompressWithDeflate(new byte[] { 0 }),
            colorSpace: "/DeviceRGB",
            softMaskExtraEntries: " /Decode [1 0]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesAscii85FlateSoftMaskAsPngAlpha() {
        byte[] softMaskCompressed = CompressWithDeflate(new byte[] { 128 });
        byte[] softMaskEncoded = Encoding.ASCII.GetBytes(EncodeAscii85(softMaskCompressed));
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 0 }),
            softMaskEncoded,
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            softMaskFilterEntry: "/Filter [/ASCII85Decode /FlateDecode]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 0, 128 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_TreatsSoftMaskNoneAsUnmaskedImage() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 255, 0, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            extraImageEntries: " /SMask /None");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_RendersBaseJpegImageXObjectWithUnresolvedSoftMask() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CreateMinimalJpeg(1, 1),
            CompressWithDeflate(new byte[] { 128 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            imageFilterEntry: "/Filter /DCTDecode");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/jpeg", image.ContentType);
        Assert.Equal(CreateMinimalJpeg(1, 1), image.Bytes);
    }

    [Fact]
    public void RenderPageAsPng_DecodesJpegImageXObjectThroughManagedRasterEngine() {
        var raster = new OfficeRasterImage(1, 1, OfficeColor.FromRgb(12, 34, 56));
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions { Quality = 100 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            imageFilterEntry: "/Filter /DCTDecode");

        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        AssertPngSignature(png);
        Assert.DoesNotContain(
            PdfPageImageRenderer.RenderPages(pdf).Single().CapabilityDiagnostics,
            diagnostic => diagnostic.Code == "render.resource.image-codec-optional");
    }

    [Fact]
    public void RenderPages_UsesOptionalSharedCodecWhenManagedJpegDecodeFails() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CreateMinimalJpeg(1, 1),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            imageFilterEntry: "/Filter /DCTDecode");
        var codec = new TestRasterImageCodec();

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, options: new PdfPageRenderOptions {
            ImageCodec = codec,
            ContinueOnError = false
        }));

        Assert.True(result.Succeeded);
        Assert.True(codec.WasCalled);
        Assert.NotEmpty(result.Bytes!);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.image-codec-optional");
    }

    [Fact]
    public void RenderPages_ReportsOptionalCodecDiagnosticForJpeg2000ImageXObject() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CreateMinimalJpeg(1, 1),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            imageFilterEntry: "/Filter /JPXDecode");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.True(result.Succeeded);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.image-codec-optional");
    }

    [Fact]
    public void RenderPage_AppliesImageXObjectExtGStateOpacity() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 255, 0 }),
            imageWidth: 1,
            extraResourceEntries: " /ExtGState << /GS1 6 0 R >>",
            contentStream: """
                /GS1 gs
                q
                20 0 0 20 40 80 cm
                /Im1 Do
                Q
                """,
            extraObjects: new[] { "6 0 obj\n<< /Type /ExtGState /ca 0.5 >>\nendobj" });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] renderedPng = PdfPageImageRenderer.RenderPageAsPng(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal(0.5D, image.Opacity, 3);
        Assert.True(OfficePngReader.TryDecode(renderedPng, out OfficeRasterImage? rendered));
        Assert.NotNull(rendered);
        OfficeColor blended = rendered!.GetPixel(50, 110);
        Assert.Equal(255, blended.R);
        Assert.InRange(blended.G, 126, 128);
        Assert.InRange(blended.B, 126, 128);
    }

    [Fact]
    public void RenderPage_ProjectsExtGStateBlendModeThroughSharedEffectGroup() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            0.94 0.5 0.13 rg
            40 80 80 60 re f
            /GS1 gs
            0.25 0.5 1 rg
            40 80 80 60 re f
            """,
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /BM /Multiply >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup effect = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(60, 90);

        Assert.Equal(OfficeBlendMode.Multiply, effect.BlendMode);
        Assert.InRange(pixel.R, (byte)59, (byte)61);
        Assert.InRange(pixel.G, (byte)63, (byte)65);
        Assert.InRange(pixel.B, (byte)32, (byte)34);
        Assert.Equal((byte)255, pixel.A);
    }

    [Fact]
    public void RenderPage_SelectsFirstSupportedBlendModeFromFallbackArray() {
        byte[] pdf = BuildSingleStreamPdf(
            "/GS1 gs\n0 0 1 rg\n20 20 40 40 re f",
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /BM [/UnknownVendorMode /Screen] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup effect = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Equal(OfficeBlendMode.Screen, effect.BlendMode);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedBlendModeId);
    }

    [Fact]
    public void RenderPage_ProjectsExtGStateFormSoftMaskThroughSharedDrawing() {
        string maskContent = "1 g\n0 0 120 200 re f";
        byte[] pdf = BuildSingleStreamPdf(
            """
            /GS1 gs
            1 0 0 rg
            0 0 240 200 re f
            """,
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /SMask 6 0 R >>\nendobj",
            "6 0 obj\n<< /S /Alpha /G 7 0 R >>\nendobj",
            BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency /CS /DeviceGray >> /Resources << >>", maskContent));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingEffectGroup effect = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.NotNull(effect.SoftMask);
        Assert.Equal(OfficeSoftMaskMode.Alpha, effect.SoftMask!.Mode);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(60, 100));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(180, 100));
    }

    [Fact]
    public void RenderPage_IgnoresBackdropColorForAlphaSoftMask() {
        byte[] pdf = BuildSingleStreamPdf(
            "/GS1 gs\n1 0 0 rg\n0 0 240 200 re f",
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /SMask 6 0 R >>\nendobj",
            "6 0 obj\n<< /S /Alpha /BC [1] /G 7 0 R >>\nendobj",
            BuildStreamObject(
                7,
                "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency /CS /DeviceGray >> /Resources << >>",
                "1 g\n0 0 120 200 re f"));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(60, 100));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(180, 100));
    }

    [Fact]
    public void RenderPage_ProjectsColoredTilingPatternThroughSharedVectorPattern() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /Pattern cs
            /P1 scn
            20 40 100 80 re f
            """,
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Matrix [1 0 0 1 0 0] /Resources << >>",
                "1 0 0 rg\n0 0 10 10 re f"));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 155));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(35, 155));
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.TilingPatternId);
    }

    [Fact]
    public void RenderPage_ProjectsBasicUncoloredTilingPatternWithPaintTint() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs\n0 1 0 /P1 scn\n20 40 100 80 re f",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Resources << >>",
                "0 g\n0 0 10 10 re f"));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), raster.GetPixel(25, 155));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(35, 155));
    }

    [Fact]
    public void RenderPage_TintsTextInsideUncoloredTilingPattern() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs\n0 1 0 /P1 scn\n20 40 100 80 re f",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 20 20] /XStep 30 /YStep 30 /Resources << /Font << /F1 6 0 R >> >>",
                "0 g\nBT /F1 8 Tf 1 0 0 1 2 12 Tm (X) Tj ET"),
            "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeDrawingTilingPattern pattern = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingTilingPattern>());
        OfficeDrawingText text = Assert.Single(pattern.Tile.Elements.OfType<OfficeDrawingText>());

        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), text.Color);
    }

    [Fact]
    public void RenderPage_ClipsTilingPatternFillThatOverlapsPageEdge() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs\n/P1 scn\n-10 40 100 80 re f",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Resources << >>",
                "1 0 0 rg\n0 0 10 10 re f"));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(5, 155));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(15, 155));
    }

    [Fact]
    public void RenderPage_UsesDeclaredCmykBaseColorForUncoloredTilingPattern() {
        byte[] pdf = BuildSingleStreamPdf(
            "/PatternCmyk cs\n0 1 1 0 /P1 scn\n20 40 100 80 re f",
            "<< /ColorSpace << /PatternCmyk [ /Pattern /DeviceCMYK ] >> /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Resources << >>",
                "0 g\n0 0 10 10 re f"));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 155));
    }

    [Fact]
    public void RenderPage_ClearsPreviousShadingWhenTilingPatternIsSelected() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs\n/S1 scn\n/P1 scn\n20 40 100 80 re f",
            "<< /Pattern << /S1 5 0 R /P1 6 0 R >> >>",
            "5 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 120 80] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj",
            BuildStreamObject(
                6,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Resources << >>",
                "1 0 0 rg\n0 0 10 10 re f"));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 155));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(35, 155));
    }

    [Fact]
    public void RenderPage_ProjectsStrokeTilingPatternThroughVectorMask() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern CS\n/P1 SCN\n8 w\n20 40 100 80 re S\n20 20 m\n120 20 l\nS",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Resources << >>",
                "1 0 0 rg\n0 0 10 10 re f"));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(21, 155));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(30, 155));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(21, 178));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(30, 178));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(21, 170));
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedTilingPatternId);
    }

    [Fact]
    public void RenderPage_ReportsMalformedTilingPatternInsteadOfClaimingSupport() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs\n/P1 scn\n20 40 100 80 re f",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 0 /YStep 20 /Resources << >>",
                "1 0 0 rg\n0 0 10 10 re f"));

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedTilingPatternId);
    }

    [Fact]
    public void RenderPage_ReportsNonInvertibleTilingPatternMatrixInsteadOfThrowing() {
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs\n/P1 scn\n20 40 100 80 re f",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(
                5,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 20 /YStep 20 /Matrix [0 0 0 0 0 0] /Resources << >>",
                "1 0 0 rg\n0 0 10 10 re f"));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingTilingPattern);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedTilingPatternId);
    }

    [Fact]
    public void RenderPage_ReportsSoftMaskWithoutTransparencyGroup() {
        byte[] pdf = BuildSingleStreamPdf(
            "/GS1 gs\n1 0 0 rg\n0 0 120 200 re f",
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /SMask 6 0 R >>\nendobj",
            "6 0 obj\n<< /S /Alpha /G 7 0 R >>\nendobj",
            BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 120 200] /Resources << >>", "1 g\n0 0 120 200 re f"));

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingEffectGroup);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedSoftMaskId);
    }

    [Fact]
    public void RenderPage_InheritsOpacityIntoFormInlineImage() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /GS1 gs
            /Fm1 Do
            """,
            "<< /XObject << /Fm1 5 0 R >> /ExtGState << /GS1 6 0 R >> >>",
            BuildStreamObject(5, "<< /Type /XObject /Subtype /Form /Resources << >>", """
                q
                20 0 0 20 40 80 cm
                BI
                /W 1
                /H 1
                /CS /RGB
                /BPC 8
                ID
                abc
                EI
                Q
                """),
            "6 0 obj\n<< /Type /ExtGState /ca 0.4 >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal(0.4D, image.Opacity, 3);
    }

    [Fact]
    public void RenderPage_AppliesTextExtGStateOpacity() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /GS1 gs
            BT
            /F1 20 Tf
            1 0 0 rg
            40 100 Td
            (Transparent) Tj
            ET
            """,
            "<< /Font << /F1 5 0 R >> /ExtGState << /GS1 6 0 R >> >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /ExtGState /ca 0.25 >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        string svg = Encoding.UTF8.GetString(PdfPageImageRenderer.RenderPageAsSvg(pdf));

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Transparent", text.Text);
        Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 64), text.Color);
        Assert.Contains("fill-opacity", svg, StringComparison.Ordinal);
        Assert.Contains("Transparent", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void RenderPage_PreservesPdfPainterOrderAcrossTextAndImages() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            extraResourceEntries: " /Font << /F1 6 0 R >>",
            contentStream: """
                BT /F1 12 Tf 20 150 Td (Before) Tj ET
                q
                20 0 0 20 40 80 cm
                /Im1 Do
                Q
                BT /F1 12 Tf 20 120 Td (After) Tj ET
                """,
            extraObjects: new[] { "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj" });

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingElement[] painted = drawing.Elements
            .Where(element => element is OfficeDrawingText || element is OfficeDrawingImage)
            .ToArray();
        Assert.Collection(
            painted,
            element => Assert.Equal("Before", Assert.IsType<OfficeDrawingText>(element).Text),
            element => Assert.IsType<OfficeDrawingImage>(element),
            element => Assert.Equal("After", Assert.IsType<OfficeDrawingText>(element).Text));
    }

    [Fact]
    public void RenderPage_RendersHairlineStrokeAsDrawableWidth() {
        byte[] pdf = BuildSingleStreamPdf("""
            0 w
            1 0 0 RG
            20 20 m
            120 20 l
            S
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape line = Assert.Single(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line);
        Assert.True(line.Shape.StrokeWidth > 0D);
    }

    [Fact]
    public void RenderPage_ProjectsPdfBaseFontIntoDrawingTextFont() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            BT
            /F1 18 Tf
            40 100 Td
            (Courier face) Tj
            ET
            """,
            "<< /Font << /F1 5 0 R >> >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier-BoldOblique >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Courier face", text.Text);
        Assert.Equal("Courier New", text.Font.FamilyName);
        Assert.True(text.Font.IsBold);
        Assert.True(text.Font.IsItalic);
    }

    [Fact]
    public void TextParser_PreservesDistinctEmbeddedSubsetFontIdentities() {
        var first = new PdfFontResource("F1", "ABCDEF+Arial", "WinAnsiEncoding", hasToUnicode: false, embeddedTrueTypeFont: new byte[] { 1, 2, 3 });
        var second = new PdfFontResource("F2", "GHIJKL+Arial", "WinAnsiEncoding", hasToUnicode: false, embeddedTrueTypeFont: new byte[] { 4, 5, 6 });
        var fonts = new Dictionary<string, PdfFontResource>(StringComparer.Ordinal) {
            ["F1"] = first,
            ["F2"] = second
        };

        List<PdfTextSpan> spans = TextContentParser.Parse(
            "BT /F1 12 Tf (First) Tj /F2 12 Tf (Second) Tj ET",
            (_, bytes) => Encoding.ASCII.GetString(bytes),
            (_, bytes) => bytes.Length * 500D,
            baseFontForResource: resource => fonts[resource].BaseFont,
            drawingFontFamilyForResource: resource => fonts[resource].DrawingFontFamily);

        Assert.Collection(
            spans,
            span => {
                Assert.Equal("ABCDEF+Arial", span.BaseFont);
                Assert.Equal(first.DrawingFontFamily, span.DrawingFontFamily);
            },
            span => {
                Assert.Equal("GHIJKL+Arial", span.BaseFont);
                Assert.Equal(second.DrawingFontFamily, span.DrawingFontFamily);
            });
        Assert.NotEqual(spans[0].DrawingFontFamily, spans[1].DrawingFontFamily);
    }

    [Fact]
    public void RenderPage_AppliesDeviceGrayImageXObjectDecodeArray() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255 }),
            colorSpace: "/DeviceGray",
            imageWidth: 2,
            extraImageEntries: " /Decode [1 0]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(0, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesDeviceRgbImageXObjectDecodeArray() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 128, 255 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            extraImageEntries: " /Decode [1 0 0 1 0 1]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 128, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ProjectsAscii85FlateImageXObjectAsPng() {
        byte[] compressed = CompressWithDeflate(new byte[] { 255, 0, 0, 0, 255, 0 });
        byte[] encoded = Encoding.ASCII.GetBytes(EncodeAscii85(compressed));
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            encoded,
            colorSpace: "/DeviceRGB",
            imageWidth: 2,
            imageFilterEntry: "/Filter [/ASCII85Decode /FlateDecode]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ResolvesNamedDeviceRgbImageXObjectColorSpace() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 128, 255 }),
            colorSpace: "/CsRgb",
            imageWidth: 1,
            extraResourceEntries: " /ColorSpace << /CsRgb /DeviceRGB >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 128, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_DecodesMismatchedPngPredictorBeforeWrappingImageXObject() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 1, 10, 10, 10, 10, 10, 10 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 2,
            extraImageEntries: " /DecodeParms << /Predictor 12 /Colors 1 /BitsPerComponent 8 /Columns 6 >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 10, 20, 30, 40, 50, 60 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesDeviceRgbImageXObjectColorKeyMaskAsPngAlpha() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 255, 0, 0, 0, 255, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 2,
            extraImageEntries: " /Mask [0 10 250 255 0 10]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 255, 0, 255, 0, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesDeviceCmykImageXObjectDecodeArray() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 0, 0, 0 }),
            imageWidth: 1,
            extraImageEntries: " /Decode [1 0 0 1 0 1 0 1]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesDeviceCmykImageXObjectColorKeyMaskAsPngAlpha() {
        byte[] cmykPixels = {
            255, 0, 0, 0,
            0, 255, 0, 0
        };
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(cmykPixels),
            imageWidth: 2,
            extraImageEntries: " /Mask [255 255 0 0 0 0 0 0]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 255, 0, 255, 0, 255, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ProjectsInlineDeviceRgbImageAsPng() {
        byte[] pdf = BuildInlineDeviceRgbImagePdf();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
        Assert.Equal(40D, image.Projection.X, 1);
        Assert.Equal(100D, image.Projection.Y, 1);
        Assert.Equal(20D, image.Projection.Width, 1);
        Assert.Equal(20D, image.Projection.Height, 1);
    }

    [Fact]
    public void RenderPage_ResolvesNamedColorSpaceForFilteredInlineImage() {
        byte[] pdf = BuildInlineNamedDeviceRgbImagePdf();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
        Assert.Equal(40D, image.Projection.X, 1);
        Assert.Equal(100D, image.Projection.Y, 1);
        Assert.Equal(20D, image.Projection.Width, 1);
        Assert.Equal(20D, image.Projection.Height, 1);
    }

    [Fact]
    public void RenderPage_PreservesRotatedImageXObjectProjection() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 255, 0, 0, 0, 0, 255 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 2,
            contentStream: """
                q
                0 20 -20 0 100 80 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.Equal(80D, image.Projection.X, 1);
        Assert.Equal(100D, image.Projection.Y, 1);
        Assert.Equal(20D, image.Projection.Width, 1);
        Assert.Equal(20D, image.Projection.Height, 1);
        Assert.Equal(-90D, image.Projection.RotationDegrees, 1);
        Assert.Equal(90D, image.Projection.RotationCenterX, 1);
        Assert.Equal(110D, image.Projection.RotationCenterY, 1);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_AppliesRectangleClipToRotatedImageXObject() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 255, 0, 0, 0, 0, 255 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 2,
            contentStream: """
                40 40 40 40 re
                W
                n
                q
                0 80 -80 0 120 40 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal(OfficeClipPathKind.Rectangle, group.ClipPath.Kind);
        Assert.Equal(40D, group.ClipPath.Width, 1);
        Assert.Equal(40D, group.ClipPath.Height, 1);
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
    }

    [Fact]
    public void RenderPage_PreservesMirroredImageXObjectProjection() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 255, 0, 0, 0, 0, 255 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 2,
            contentStream: """
                q
                -20 0 0 20 100 80 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.Equal(80D, image.Projection.X, 1);
        Assert.Equal(100D, image.Projection.Y, 1);
        Assert.Equal(20D, image.Projection.Width, 1);
        Assert.Equal(20D, image.Projection.Height, 1);
        Assert.True(image.Projection.FlipHorizontal);
        Assert.False(image.Projection.FlipVertical);
        Assert.Equal(0D, image.Projection.RotationDegrees, 1);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_ProjectsImageMaskXObjectAsPngAlpha() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0xA0 }),
            colorSpace: string.Empty,
            bitsPerComponent: 1,
            imageWidth: 4,
            extraImageEntries: " /ImageMask true");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(6, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 0, 0, 0, 255, 0, 0, 0, 0, 0, 0, 0, 255, 0, 0, 0, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_AppliesImageMaskXObjectDecodeArray() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0x80 }),
            colorSpace: string.Empty,
            bitsPerComponent: 1,
            imageWidth: 2,
            extraImageEntries: " /ImageMask true /Decode [1 0]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 255 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ProjectsImageMaskXObjectWithCurrentFillColor() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0x80 }),
            colorSpace: string.Empty,
            bitsPerComponent: 1,
            imageWidth: 2,
            extraImageEntries: " /ImageMask true",
            contentStream: """
                1 0 0 rg
                q
                20 0 0 20 40 80 cm
                /Im1 Do
                Q
                0 0 1 rg
                q
                20 0 0 20 80 80 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(2, drawing.Images.Count);
        Assert.Equal(new byte[] { 0, 255, 0, 0, 255, 255, 0, 0, 0 }, PdfPngTestImages.DecodeStoredPngIdat(drawing.Images[0].Bytes));
        Assert.Equal(new byte[] { 0, 0, 0, 255, 255, 0, 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(drawing.Images[1].Bytes));
    }

    [Fact]
    public void RenderPage_ProjectsIndexedImageXObjectPaletteAsPng() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0x19 }),
            colorSpace: "[/Indexed /DeviceRGB 2 <FF000000FF000000FF>]",
            bitsPerComponent: 2,
            imageWidth: 4);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0, 255, 0, 0, 0, 255, 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ResolvesNamedIndexedImageXObjectColorSpace() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0x19 }),
            colorSpace: "/CsIndexed",
            bitsPerComponent: 2,
            imageWidth: 4,
            extraResourceEntries: " /ColorSpace << /CsIndexed [/Indexed /DeviceRGB 2 <FF000000FF000000FF>] >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg }));

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(2, PdfPngTestImages.ReadPngColorType(image.Bytes));
        Assert.Equal(new byte[] { 0, 255, 0, 0, 0, 255, 0, 0, 0, 255, 0, 255, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.colorspace-unsupported");
    }

    [Fact]
    public void RenderPage_ReportsOnlySelectedUnsupportedContentColorSpaces() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /CsSpot cs
            0.5 scn
            40 80 70 40 re
            f
            """,
            "<< /ColorSpace << /CsUnused [/DeviceN [/Cyan] /DeviceCMYK 5 0 R] /CsSpot [/Separation /Spot /DeviceRGB 5 0 R] >> >>",
            "5 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 0 0] /N 1 >>\nendobj");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg }));

        PdfRenderCapabilityDiagnostic diagnostic = Assert.Single(
            result.CapabilityDiagnostics,
            item => item.Code == "render.resource.colorspace-unsupported");
        Assert.Equal("CsSpot", diagnostic.Subject);
    }

    [Fact]
    public void RenderPage_AppliesIndexedImageXObjectDecodeArray() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0x1C }),
            colorSpace: "[/Indexed /DeviceRGB 2 <FF000000FF000000FF>]",
            bitsPerComponent: 2,
            imageWidth: 3,
            extraImageEntries: " /Decode [2 0]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        var image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(new byte[] { 0, 0, 0, 255, 0, 255, 0, 255, 0, 0 }, PdfPngTestImages.DecodeStoredPngIdat(image.Bytes));
    }

    [Fact]
    public void RenderPage_ProjectsFilledAndStrokedRectanglesFromGeneratedPdf() {
        byte[] pdf = PdfDocument.Create()
            .Rectangle(120, 40, strokeColor: PdfColor.FromRgb(0, 64, 128), strokeWidth: 2, fillColor: PdfColor.FromRgb(204, 238, 255))
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(204, 238, 255));
        Assert.Equal(OfficeColor.FromRgb(204, 238, 255), shape.Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgb(0, 64, 128), shape.Shape.StrokeColor);
        Assert.Equal(2D, shape.Shape.StrokeWidth);
    }

    [Fact]
    public void RenderPage_ProjectsLegacyUppercaseFillOperator() {
        byte[] pdf = BuildSingleStreamPdf("""
            0 0.5 0 rg
            40 80 100 40 re
            F
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(0, 128, 0));
        Assert.Equal(40D, shape.X, 1);
        Assert.Equal(80D, shape.Y, 1);
        Assert.Equal(100D, shape.Shape.Width, 1);
        Assert.Equal(40D, shape.Shape.Height, 1);

        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("<rect", svgText, StringComparison.Ordinal);
        Assert.Contains("#008000", svgText, StringComparison.OrdinalIgnoreCase);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_ProjectsCmykFillAndStrokeOperators() {
        byte[] pdf = BuildSingleStreamPdf("""
            1 0 0 0 k
            0 1 0 0 K
            3 w
            40 80 120 40 re
            B
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(0, 255, 255));
        Assert.Equal(OfficeColor.FromRgb(0, 255, 255), shape.Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 255), shape.Shape.StrokeColor);
        Assert.Equal(3D, shape.Shape.StrokeWidth);
    }

    [Fact]
    public void RenderPage_ProjectsColorSpaceSetColorOperators() {
        byte[] pdf = BuildSingleStreamPdf("""
            /DeviceRGB cs
            0.2 0.4 0.6 sc
            /DeviceGray CS
            0.25 SC
            4 w
            40 80 80 40 re
            B
            /DeviceCMYK cs
            0 1 1 0 scn
            140 80 40 40 re
            f
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape rgbShape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(51, 102, 153));
        Assert.Equal(OfficeColor.FromRgb(64, 64, 64), rgbShape.Shape.StrokeColor);
        Assert.Equal(4D, rgbShape.Shape.StrokeWidth);

        OfficeDrawingShape cmykShape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.Red);
        Assert.Null(cmykShape.Shape.StrokeColor);
    }

    [Fact]
    public void RenderPage_ProjectsResourceColorSpaceAliases() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /CsRgb cs
            0.1 0.2 0.3 sc
            /CsGray CS
            0.5 SC
            2 w
            40 80 80 40 re
            B
            """,
            "<< /ColorSpace << /CsRgb /DeviceRGB /CsGray [/DeviceGray] >> >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(26, 51, 76));
        Assert.Equal(OfficeColor.FromRgb(128, 128, 128), shape.Shape.StrokeColor);
        Assert.Equal(2D, shape.Shape.StrokeWidth);
    }

    [Fact]
    public void RenderPage_ProjectsCalibratedAndIccColorSpacesThroughManagedConversion() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /CsCal cs
            0.1 0.2 0.3 scn
            40 80 70 40 re
            f
            /CsIcc cs
            0.8 0.1 0.2 scn
            130 80 70 40 re
            f
            """,
            "<< /ColorSpace << /CsCal [/CalRGB << /WhitePoint [0.9642 1 0.8249] /Gamma [2.2 1.8 1.4] /Matrix [0.7 0.2 0.1 0.1 0.8 0.1 0.2 0.1 0.7] >>] /CsIcc [/ICCBased 5 0 R] >> >>",
            "5 0 obj\n<< /N 3 /Length 0 >>\nstream\n\nendstream\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg }));

        OfficeColor calibrated = OfficeColorSpaceConverter.FromCalibratedRgb(
            0.1D, 0.2D, 0.3D,
            0.9642D, 1D, 0.8249D,
            new[] { 2.2D, 1.8D, 1.4D },
            new[] { 0.7D, 0.2D, 0.1D, 0.1D, 0.8D, 0.1D, 0.2D, 0.1D, 0.7D });
        Assert.Contains(drawing.Shapes, item => item.Shape.FillColor == calibrated);
        Assert.NotEqual(OfficeColorSpaceConverter.FromCalibratedRgb(0.1D, 0.2D, 0.3D, 0.9505D, 1D, 1.089D), calibrated);
        Assert.Contains(drawing.Shapes, item => item.Shape.FillColor == OfficeColor.FromRgb(204, 26, 51));
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == "render.resource.colorspace-unsupported");
    }

    [Fact]
    public void RenderPage_ProjectsLabColorSpaceToSrgb() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /CsLab cs
            53.24 80.09 67.2 scn
            40 80 100 40 re
            f
            """,
            "<< /ColorSpace << /CsLab [/Lab << /WhitePoint [0.9642 1 0.8249] /Range [-128 127 -128 127] >>] >> >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeColor color = Assert.Single(drawing.Shapes).Shape.FillColor!.Value;
        Assert.InRange(color.R, 245, 255);
        Assert.InRange(color.G, 0, 15);
        Assert.InRange(color.B, 0, 15);
    }

    [Fact]
    public void RenderPage_ProjectsAxialShadingResourceAsLinearGradient() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 80 120 40 re
            W
            n
            /Sh1 sh
            """,
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.FillGradient != null);
        Assert.Equal(20D, shape.X);
        Assert.Equal(80D, shape.Y);
        Assert.Equal(120D, shape.Shape.Width);
        Assert.Equal(40D, shape.Shape.Height);
        Assert.NotNull(shape.Shape.ClipPath);
        OfficeLinearGradient gradient = shape.Shape.FillGradient!;
        Assert.Equal(0D, gradient.StartX);
        Assert.Equal(1D, gradient.EndX);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[1].Color);

        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor leftPixel = raster!.GetPixel(28, 100);
        OfficeColor rightPixel = raster.GetPixel(132, 100);
        Assert.True(leftPixel.R > leftPixel.B);
        Assert.True(rightPixel.B > rightPixel.R);
    }

    [Fact]
    public void RenderPage_PreservesStitchedShadingFunctionsAsMultipleGradientStops() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 80 120 40 re
            W
            n
            /Sh1 sh
            """,
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function << /FunctionType 3 /Domain [0 1] /Functions [6 0 R 7 0 R] /Bounds [0.5] /Encode [0 1 0 1] >> /Extend [true true] >>\nendobj",
            "6 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 1 0] /N 1 >>\nendobj",
            "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeLinearGradient gradient = Assert.Single(drawing.Shapes).Shape.FillGradient!;
        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(0.5D, gradient.Stops[1].Offset, 3);
        Assert.Equal(OfficeColor.Lime, gradient.Stops[1].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[2].Color);
    }

    [Fact]
    public void RenderPage_UsesCmykDefaultsForOmittedShadingFunctionColors() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 80 120 40 re
            W
            n
            /Sh1 sh
            """,
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceCMYK /Coords [20 80 140 80] /Function << /FunctionType 2 /Domain [0 1] /N 1 >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.FillGradient != null);
        OfficeLinearGradient gradient = shape.Shape.FillGradient!;
        Assert.Equal(OfficeColor.White, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Black, gradient.Stops[1].Color);
    }

    [Fact]
    public void RenderPage_ProjectsShadingPatternFillAsLinearGradient() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /Pattern cs
            /P1 scn
            20 80 120 40 re
            f
            """,
            "<< /Pattern << /P1 5 0 R >> >>",
            "5 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.FillGradient != null);
        Assert.Equal(20D, shape.X);
        Assert.Equal(80D, shape.Y);
        Assert.Equal(120D, shape.Shape.Width);
        Assert.Equal(40D, shape.Shape.Height);
        Assert.Null(shape.Shape.FillColor);
        OfficeLinearGradient gradient = shape.Shape.FillGradient!;
        Assert.Equal(0D, gradient.StartX);
        Assert.Equal(1D, gradient.EndX);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[1].Color);

        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor leftPixel = raster!.GetPixel(28, 100);
        OfficeColor rightPixel = raster.GetPixel(132, 100);
        Assert.True(leftPixel.R > leftPixel.B);
        Assert.True(rightPixel.B > rightPixel.R);
    }

    [Fact]
    public void RenderPage_UsesBezierControlsForShadingPatternBounds() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /Pattern cs
            /P1 scn
            20 40 m
            20 120 140 120 140 40 c
            140 30 l
            20 30 l
            h
            f
            """,
            "<< /Pattern << /P1 5 0 R >> >>",
            "5 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 40 140 40] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.FillGradient != null);
        Assert.Equal(20D, shape.X);
        Assert.Equal(80D, shape.Y);
        Assert.Equal(120D, shape.Shape.Width);
        Assert.Equal(90D, shape.Shape.Height);
        Assert.NotNull(shape.Shape.FillGradient);
    }

    [Fact]
    public void RenderPage_ProjectsShadingPatternStrokeAsLinearGradient() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /Pattern CS
            /P1 SCN
            8 w
            20 80 120 40 re
            S
            """,
            "<< /Pattern << /P1 5 0 R >> >>",
            "5 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 80 140 80] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.StrokeGradient != null);
        Assert.Null(shape.Shape.StrokeColor);
        Assert.Equal(8D, shape.Shape.StrokeWidth);
        OfficeLinearGradient gradient = shape.Shape.StrokeGradient!;
        Assert.Equal(0D, gradient.StartX);
        Assert.Equal(1D, gradient.EndX);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[1].Color);

        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
        Assert.Contains("stroke=\"url(#", svgText, StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor leftPixel = raster!.GetPixel(28, 80);
        OfficeColor rightPixel = raster.GetPixel(132, 80);
        Assert.True(leftPixel.R > leftPixel.B);
        Assert.True(rightPixel.B > rightPixel.R);
    }

    [Fact]
    public void RenderPage_ProjectsShadingPatternStrokeLineAsLinearGradient() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /Pattern CS
            /P1 SCN
            8 w
            20 100 m
            140 100 l
            S
            """,
            "<< /Pattern << /P1 5 0 R >> >>",
            "5 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 140 100] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.Kind == OfficeShapeKind.Line && item.Shape.StrokeGradient != null);
        Assert.Null(shape.Shape.StrokeColor);
        Assert.Equal(8D, shape.Shape.StrokeWidth);
        OfficeLinearGradient gradient = shape.Shape.StrokeGradient!;
        Assert.Equal(0D, gradient.StartX);
        Assert.Equal(1D, gradient.EndX);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[1].Color);

        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
        Assert.Contains("stroke=\"url(#", svgText, StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor leftPixel = raster!.GetPixel(28, 100);
        OfficeColor rightPixel = raster.GetPixel(132, 100);
        Assert.True(leftPixel.R > leftPixel.B);
        Assert.True(rightPixel.B > rightPixel.R);
    }

    [Fact]
    public void RenderPage_ProjectsRadialShadingResourceAsRadialGradient() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 40 120 120 re
            W
            n
            /Sh1 sh
            """,
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 3 /ColorSpace /DeviceRGB /Coords [80 100 0 80 100 60] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.FillRadialGradient != null);
        Assert.Equal(20D, shape.X);
        Assert.Equal(40D, shape.Y);
        Assert.Equal(120D, shape.Shape.Width);
        Assert.Equal(120D, shape.Shape.Height);
        OfficeRadialGradient gradient = shape.Shape.FillRadialGradient!;
        Assert.Equal(0.5D, gradient.StartX);
        Assert.Equal(0.5D, gradient.StartY);
        Assert.Equal(0D, gradient.StartRadius);
        Assert.Equal(0.5D, gradient.EndX);
        Assert.Equal(0.5D, gradient.EndY);
        Assert.Equal(0.5D, gradient.EndRadius);
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[1].Color);

        string svgText = Encoding.UTF8.GetString(svg);
        Assert.Contains("<radialGradient", svgText, StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor centerPixel = raster!.GetPixel(80, 100);
        OfficeColor edgePixel = raster.GetPixel(138, 100);
        Assert.True(centerPixel.R > centerPixel.B);
        Assert.True(edgePixel.B > edgePixel.R);
    }

    [Fact]
    public void RenderPage_PreservesOffCanvasRadialShadingCenters() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 40 120 120 re
            W
            n
            /Sh1 sh
            """,
            "<< /Shading << /Sh1 5 0 R >> >>",
            "5 0 obj\n<< /ShadingType 3 /ColorSpace /DeviceRGB /Coords [-20 100 0 -20 100 60] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item => item.Shape.FillRadialGradient != null);
        OfficeRadialGradient gradient = shape.Shape.FillRadialGradient!;
        Assert.True(gradient.StartX < 0D);
        Assert.True(gradient.EndX < 0D);
        Assert.Equal(0.5D, gradient.StartY);
        Assert.Equal(0.5D, gradient.EndY);
    }

    [Fact]
    public void RenderPage_ProjectsTextFillColorFromContentStream() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            BT
            /F1 18 Tf
            /CsRgb cs
            0.8 0.1 0.2 sc
            40 120 Td
            (Ruby text) Tj
            ET
            """,
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> /ColorSpace << /CsRgb /DeviceRGB >> >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Ruby text", text.Text);
        Assert.Equal(OfficeColor.FromRgb(204, 26, 51), text.Color);
    }

    [Fact]
    public void RenderPage_DoesNotPaintInvisibleTextRenderingMode() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            BT
            /F1 18 Tf
            3 Tr
            40 130 Td
            (Hidden OCR) Tj
            0 Tr
            0 -30 Td
            (Visible label) Tj
            ET
            """,
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Visible label", text.Text);
    }

    [Fact]
    public void RenderPage_ProjectsStrokeTextRenderingModeWithStrokeColorAndOpacity() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /GS1 gs
            BT
            /F1 18 Tf
            1 0 0 rg
            0 0 1 RG
            1 Tr
            40 130 Td
            (Stroke text) Tj
            ET
            """,
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /CA 0.5 /ca 0.25 >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        string svg = Encoding.UTF8.GetString(PdfPageImageRenderer.RenderPageAsSvg(pdf));

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Stroke text", text.Text);
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), text.Color);
        Assert.Contains("#0000FF", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("fill-opacity", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void RenderPage_DoesNotPaintClipOnlyTextRenderingMode() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            q
            BT
            /F1 18 Tf
            7 Tr
            40 130 Td
            (Clip only) Tj
            ET
            Q
            BT
            /F1 18 Tf
            0 Tr
            40 100 Td
            (Painted label) Tj
            ET
            """,
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Painted label", text.Text);
    }

    [Fact]
    public void RenderPage_ProjectsRotatedTextMatrix() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            BT
            /F1 18 Tf
            0 1 -1 0 120 60 Tm
            (Rotated) Tj
            ET
            """,
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Rotated", text.Text);
        Assert.Equal(-90D, text.RotationDegrees, 1);
        Assert.Equal(text.X, text.RotationCenterX, 1);
    }

    [Fact]
    public void RenderPage_AppliesPageRotationToDrawingProjection() {
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            """
            0 0.5 0 rg
            20 40 30 50 re
            f
            """,
            "<< >>",
            "/Rotate 90");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        Assert.Equal(200D, drawing.Width);
        Assert.Equal(240D, drawing.Height);
        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(0, 128, 0));
        Assert.Equal(110D, shape.X, 1);
        Assert.Equal(190D, shape.Y, 1);
        Assert.Equal(50D, shape.Shape.Width, 1);
        Assert.Equal(30D, shape.Shape.Height, 1);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_ProjectsCloseFillAndStrokePathOperator() {
        byte[] pdf = BuildSingleStreamPdf("""
            0 0 1 rg
            0 G
            2 w
            40 80 m
            160 80 l
            160 120 l
            40 120 l
            b
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(0, 0, 255));
        Assert.Equal(OfficeColor.FromRgb(0, 0, 255), shape.Shape.FillColor);
        Assert.Equal(OfficeColor.Black, shape.Shape.StrokeColor);
        Assert.Equal(2D, shape.Shape.StrokeWidth);
    }

    [Fact]
    public void RenderPage_ProjectsVectorShapesFromFormXObjects() {
        const string formContent = """
            0 0.5 0 rg
            0 0 100 40 re
            f
            """;
        string formObject = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Resources << >>",
            formContent);
        byte[] pdf = BuildSingleStreamPdf(
            """
            q
            1 0 0 1 50 70 cm
            /Fm1 Do
            Q
            """,
            "<< /XObject << /Fm1 5 0 R >> >>",
            formObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(0, 128, 0));
        Assert.Equal(50D, shape.X, 1);
        Assert.Equal(90D, shape.Y, 1);
        Assert.Equal(100D, shape.Shape.Width, 1);
        Assert.Equal(40D, shape.Shape.Height, 1);
    }

    [Fact]
    public void RenderPage_ClipsFormXObjectVectorContentToBoundingBox() {
        const string formContent = """
            1 0 0 rg
            0 0 100 40 re
            f
            """;
        string formObject = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Resources << >>",
            formContent);
        byte[] pdf = BuildSingleStreamPdf(
            """
            q
            1 0 0 1 50 70 cm
            /Fm1 Do
            Q
            """,
            "<< /XObject << /Fm1 5 0 R >> >>",
            formObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.Red);
        Assert.Equal(100D, shape.Shape.Width, 1);
        Assert.NotNull(shape.Shape.ClipPath);
        Assert.Equal(OfficeClipPathKind.Rectangle, shape.Shape.ClipPath!.Kind);
        Assert.Equal(40D, shape.Shape.ClipPath.Width, 1);
        Assert.Equal(40D, shape.Shape.ClipPath.Height, 1);
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor inside = raster!.GetPixel(70, 100);
        OfficeColor outside = raster.GetPixel(120, 100);
        Assert.True(inside.R > 200 && inside.G < 80 && inside.B < 80);
        Assert.Equal(OfficeColor.White, outside);
    }

    [Fact]
    public void RenderPage_ClipsFormXObjectImageContentToBoundingBox() {
        const string formContent = """
            q
            100 0 0 40 0 0 cm
            BI
            /W 1
            /H 1
            /CS /RGB
            /BPC 8
            ID
            abc
            EI
            Q
            """;
        string formObject = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Resources << >>",
            formContent);
        byte[] pdf = BuildSingleStreamPdf(
            """
            q
            1 0 0 1 50 70 cm
            /Fm1 Do
            Q
            """,
            "<< /XObject << /Fm1 5 0 R >> >>",
            formObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(image.Projection.HasCrop);
        Assert.Equal(50D, image.Projection.X, 1);
        Assert.Equal(90D, image.Projection.Y, 1);
        Assert.Equal(40D, image.Projection.Width, 1);
        Assert.Equal(40D, image.Projection.Height, 1);
        Assert.Equal(0D, image.Projection.SourceLeft);
        Assert.Equal(0D, image.Projection.SourceTop);
        Assert.Equal(0.4D, image.Projection.SourceWidth, 3);
        Assert.Equal(1D, image.Projection.SourceHeight);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor inside = raster!.GetPixel(70, 100);
        OfficeColor outside = raster.GetPixel(120, 100);
        Assert.NotEqual(OfficeColor.White, inside);
        Assert.Equal(OfficeColor.White, outside);
    }

    [Fact]
    public void RenderPage_ClipsFormXObjectTextContentToBoundingBox() {
        const string formContent = """
            BT
            /F1 24 Tf
            0 12 Td
            (MMMMMMMM) Tj
            ET
            """;
        string formObject = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Resources << /Font << /F1 6 0 R >> >>",
            formContent);
        string fontObject = "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf(
            """
            q
            1 0 0 1 50 70 cm
            /Fm1 Do
            Q
            """,
            "<< /XObject << /Fm1 5 0 R >> >>",
            formObject,
            fontObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal(50D, group.X, 1);
        Assert.Equal(90D, group.Y, 1);
        Assert.Equal(OfficeClipPathKind.Rectangle, group.ClipPath.Kind);
        Assert.Equal(40D, group.ClipPath.Width, 1);
        Assert.Equal(40D, group.ClipPath.Height, 1);
        Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        bool hasInkInsideClip = false;
        for (int y = 90; y <= 130 && !hasInkInsideClip; y++) {
            for (int x = 50; x <= 90; x++) {
                if (raster!.GetPixel(x, y) != OfficeColor.White) {
                    hasInkInsideClip = true;
                    break;
                }
            }
        }

        OfficeColor outside = raster.GetPixel(105, 105);
        Assert.True(hasInkInsideClip);
        Assert.Equal(OfficeColor.White, outside);
    }

    [Fact]
    public void RenderPage_AppliesPathClipToTextContent() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 60 m
            140 60 l
            80 160 l
            h
            W
            n
            BT
            /F1 48 Tf
            20 90 Td
            (WWWWWW) Tj
            ET
            """,
            "<< /Font << /F1 5 0 R >> >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Equal(OfficeFillRule.NonZero, group.ClipPath.FillRule);
        Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        bool hasInkInsideClip = false;
        for (int y = 75; y <= 125 && !hasInkInsideClip; y++) {
            for (int x = 65; x <= 115; x++) {
                if (raster!.GetPixel(x, y) != OfficeColor.White) {
                    hasInkInsideClip = true;
                    break;
                }
            }
        }

        OfficeColor outside = raster!.GetPixel(45, 75);
        Assert.True(hasInkInsideClip);
        Assert.Equal(OfficeColor.White, outside);
    }

    [Fact]
    public void RenderPage_ProjectsGeneratedExtGStateOpacity() {
        var rectangle = OfficeShape.Rectangle(120, 40);
        rectangle.FillColor = OfficeColor.FromRgb(204, 238, 255);
        rectangle.StrokeColor = OfficeColor.FromRgb(0, 64, 128);
        rectangle.StrokeWidth = 2D;
        rectangle.FillOpacity = 0.35D;
        rectangle.StrokeOpacity = 0.75D;

        byte[] pdf = PdfDocument.Create()
            .Shape(rectangle)
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(204, 238, 255));
        Assert.Equal(0.35D, shape.Shape.FillOpacity);
        Assert.Equal(0.75D, shape.Shape.StrokeOpacity);
    }

    [Fact]
    public void RenderPage_AppliesExtGStateStrokeStyle() {
        byte[] pdf = BuildSingleStreamPdf(
            """
            /GS1 gs
            0 0 1 RG
            20 100 m
            140 100 l
            S
            """,
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /LW 6 /LC 1 /LJ 2 /D [[8 3] 0] /CA 0.6 >>\nendobj");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Line &&
            item.Shape.StrokeColor == OfficeColor.Blue);
        Assert.Equal(6D, shape.Shape.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, shape.Shape.StrokeDashStyle);
        Assert.Equal(OfficeStrokeLineCap.Round, shape.Shape.StrokeLineCap);
        Assert.Equal(OfficeStrokeLineJoin.Bevel, shape.Shape.StrokeLineJoin);
        Assert.Equal(0.6D, shape.Shape.StrokeOpacity);
    }

    [Fact]
    public void RenderPage_PreservesGeneratedLineEndpoints() {
        byte[] pdf = PdfDocument.Create()
            .Line(
                90,
                0,
                0,
                30,
                strokeColor: PdfColor.FromRgb(16, 96, 48),
                strokeWidth: 3,
                strokeDashStyle: OfficeStrokeDashStyle.Dash,
                strokeLineCap: OfficeStrokeLineCap.Square,
                strokeLineJoin: OfficeStrokeLineJoin.Bevel)
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Line &&
            item.Shape.StrokeColor == OfficeColor.FromRgb(16, 96, 48));
        Assert.Equal(3D, shape.Shape.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, shape.Shape.StrokeDashStyle);
        Assert.Equal(OfficeStrokeLineCap.Square, shape.Shape.StrokeLineCap);
        Assert.Equal(OfficeStrokeLineJoin.Bevel, shape.Shape.StrokeLineJoin);
        Assert.Equal(2, shape.Shape.Points.Count);
        Assert.True(shape.Shape.Points[0].X > shape.Shape.Points[1].X);
        Assert.True(shape.Shape.Points[0].Y < shape.Shape.Points[1].Y);
    }

    [Fact]
    public void RenderPage_ProjectsGeneratedBezierPaths() {
        var ellipse = OfficeShape.Ellipse(80, 40);
        ellipse.FillColor = OfficeColor.FromRgb(245, 250, 255);
        ellipse.StrokeColor = OfficeColor.FromRgb(15, 98, 160);
        ellipse.StrokeWidth = 2D;

        byte[] pdf = PdfDocument.Create()
            .Shape(ellipse)
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Path &&
            item.Shape.FillColor == OfficeColor.FromRgb(245, 250, 255));
        Assert.Equal(OfficeColor.FromRgb(15, 98, 160), shape.Shape.StrokeColor);
        Assert.Equal(2D, shape.Shape.StrokeWidth);
        Assert.Contains(shape.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
    }

    [Fact]
    public void RenderPage_PreservesPdfPathFillRules() {
        const string sameWindingNestedPath = """
            1 0 0 rg
            20 20 m
            140 20 l
            140 140 l
            20 140 l
            h
            60 60 m
            100 60 l
            100 100 l
            60 100 l
            h
            """;
        byte[] nonZeroPdf = BuildSingleStreamPdf(sameWindingNestedPath + "\nf");
        byte[] evenOddPdf = BuildSingleStreamPdf(sameWindingNestedPath + "\nf*");

        OfficeDrawing nonZeroDrawing = PdfPageImageRenderer.RenderPage(nonZeroPdf);
        OfficeDrawing evenOddDrawing = PdfPageImageRenderer.RenderPage(evenOddPdf);
        byte[] evenOddSvg = PdfPageImageRenderer.RenderPageAsSvg(evenOddPdf);
        byte[] nonZeroPng = PdfPageImageRenderer.RenderPageAsPng(nonZeroPdf);
        byte[] evenOddPng = PdfPageImageRenderer.RenderPageAsPng(evenOddPdf);

        OfficeDrawingShape nonZeroShape = Assert.Single(nonZeroDrawing.Shapes);
        OfficeDrawingShape evenOddShape = Assert.Single(evenOddDrawing.Shapes);
        Assert.Equal(OfficeFillRule.NonZero, nonZeroShape.Shape.FillRule);
        Assert.Equal(OfficeFillRule.EvenOdd, evenOddShape.Shape.FillRule);
        Assert.Contains("fill-rule=\"evenodd\"", Encoding.UTF8.GetString(evenOddSvg), StringComparison.Ordinal);

        Assert.True(OfficePngReader.TryDecode(nonZeroPng, out OfficeRasterImage? nonZeroRaster));
        Assert.True(OfficePngReader.TryDecode(evenOddPng, out OfficeRasterImage? evenOddRaster));
        OfficeColor nonZeroCenter = nonZeroRaster!.GetPixel(80, 120);
        OfficeColor evenOddCenter = evenOddRaster!.GetPixel(80, 120);
        Assert.True(nonZeroCenter.R > 240 && nonZeroCenter.G < 20 && nonZeroCenter.B < 20);
        Assert.True(evenOddCenter.R > 240 && evenOddCenter.G > 240 && evenOddCenter.B > 240);
    }

    [Fact]
    public void RenderPage_ClosesCurrentSubpathAfterAbandonedMove() {
        byte[] pdf = BuildSingleStreamPdf("""
            1 0 0 rg
            1000 1000 m
            20 20 m
            80 20 l
            80 80 l
            20 80 l
            h
            f
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes);
        Assert.Equal(20D, shape.X, 1);
        Assert.Equal(120D, shape.Y, 1);
        Assert.Equal(60D, shape.Shape.Width, 1);
        Assert.Equal(60D, shape.Shape.Height, 1);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor center = raster!.GetPixel(50, 150);
        Assert.True(center.R > 240 && center.G < 20 && center.B < 20);
    }

    [Fact]
    public void RenderPage_ProjectsGeneratedShapeClipPath() {
        var rectangle = OfficeShape.Rectangle(90, 40);
        rectangle.FillColor = OfficeColor.WhiteSmoke;
        rectangle.StrokeColor = OfficeColor.SteelBlue;
        rectangle.StrokeWidth = 1.5D;
        rectangle.ClipPath = OfficeClipPath.Rectangle(45, 20);

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(rectangle)
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.WhiteSmoke);
        Assert.NotNull(shape.Shape.ClipPath);
        Assert.Equal(OfficeClipPathKind.Rectangle, shape.Shape.ClipPath!.Kind);
        Assert.Equal(45D, shape.Shape.ClipPath.Width);
        Assert.Equal(20D, shape.Shape.ClipPath.Height);
    }

    [Fact]
    public void RenderPage_ProjectsGeneratedImageClipPathAsCroppedProjection() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Image(PdfPngTestImages.CreateRgbPng(2, 2), 24, 24, clipPath: OfficeClipPath.Rectangle(12, 12))
            .ToBytes();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(image.Projection.HasCrop);
        Assert.Equal(12D, image.Projection.Width);
        Assert.Equal(12D, image.Projection.Height);
        Assert.Equal(0D, image.Projection.SourceLeft);
        Assert.Equal(0D, image.Projection.SourceTop);
        Assert.Equal(0.5D, image.Projection.SourceWidth);
        Assert.Equal(0.5D, image.Projection.SourceHeight);
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
    }

    [Fact]
    public void RenderPage_CropsImageXObjectAtPageEdge() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            contentStream: """
                q
                80 0 0 80 200 60 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(image.Projection.HasCrop);
        Assert.Equal(200D, image.Projection.X);
        Assert.Equal(60D, image.Projection.Y);
        Assert.Equal(40D, image.Projection.Width);
        Assert.Equal(80D, image.Projection.Height);
        Assert.Equal(0D, image.Projection.SourceLeft);
        Assert.Equal(0.5D, image.Projection.SourceWidth);
    }

    [Fact]
    public void RenderPage_IntersectsSuccessiveRectangleClipsForImageXObject() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            contentStream: """
                0 0 100 100 re
                W
                n
                40 0 100 100 re
                W
                n
                q
                100 0 0 100 0 0 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(image.Projection.HasCrop);
        Assert.Equal(40D, image.Projection.X);
        Assert.Equal(100D, image.Projection.Y);
        Assert.Equal(60D, image.Projection.Width);
        Assert.Equal(100D, image.Projection.Height);
        Assert.Equal(0.4D, image.Projection.SourceLeft);
        Assert.Equal(0.6D, image.Projection.SourceWidth);
    }

    [Fact]
    public void RenderPage_CarriesCallerClipIntoFormXObjectVectorContent() {
        string formObject = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << >>",
            """
            1 0 0 rg
            0 0 100 100 re
            f
            """);
        byte[] pdf = BuildSingleStreamPdf(
            """
            20 20 40 40 re
            W
            n
            q
            /Fm1 Do
            Q
            """,
            "<< /XObject << /Fm1 5 0 R >> >>",
            formObject);

        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor clippedInside = raster!.GetPixel(30, 150);
        OfficeColor clippedOutside = raster.GetPixel(10, 110);
        Assert.True(clippedInside.R > 240 && clippedInside.G < 20 && clippedInside.B < 20);
        Assert.True(clippedOutside.R > 240 && clippedOutside.G > 240 && clippedOutside.B > 240);
    }

    [Fact]
    public void RenderPage_AppliesPathClipToImageXObject() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            contentStream: """
                80 40 m
                140 100 l
                80 160 l
                20 100 l
                h
                W
                n
                q
                120 0 0 120 20 40 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Equal(OfficeFillRule.NonZero, group.ClipPath.FillRule);
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor center = raster!.GetPixel(80, 100);
        OfficeColor clippedCorner = raster.GetPixel(25, 45);
        Assert.True(center.G > 240 && center.R < 20 && center.B < 20);
        Assert.True(clippedCorner.R > 240 && clippedCorner.G > 240 && clippedCorner.B > 240);
    }

    [Fact]
    public void RenderPage_AppliesPathClipToVectorShapeWhenBoundsDiffer() {
        byte[] pdf = BuildSingleStreamPdf("""
            80 160 m
            140 100 l
            80 40 l
            20 100 l
            h
            W
            n
            1 0 0 rg
            40 40 100 120 re
            f
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);

        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor center = raster!.GetPixel(80, 100);
        OfficeColor clippedCorner = raster.GetPixel(45, 45);
        Assert.True(center.R > 240 && center.G < 20 && center.B < 20);
        Assert.True(clippedCorner.R > 240 && clippedCorner.G > 240 && clippedCorner.B > 240);
    }

    [Fact]
    public void RenderPage_DoesNotWidenUnsupportedConcavePathClipIntersections() {
        byte[] pdf = BuildSingleStreamPdf("""
            40 40 m
            120 40 l
            120 160 l
            40 160 l
            h
            W
            n
            20 20 m
            140 20 l
            140 60 l
            60 60 l
            60 140 l
            20 140 l
            h
            W
            n
            1 0 0 rg
            0 0 160 200 re
            f
            """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        Assert.Empty(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor previouslyWidenedPixel = raster!.GetPixel(100, 100);
        Assert.True(previouslyWidenedPixel.R > 240 && previouslyWidenedPixel.G > 240 && previouslyWidenedPixel.B > 240);
    }

    [Fact]
    public void RenderPage_AppliesBezierPathClipToImageXObject() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(
            CompressWithDeflate(new byte[] { 0, 255, 0 }),
            colorSpace: "/DeviceRGB",
            imageWidth: 1,
            contentStream: """
                80 40 m
                113.137 40 140 66.863 140 100 c
                140 133.137 113.137 160 80 160 c
                46.863 160 20 133.137 20 100 c
                20 66.863 46.863 40 80 40 c
                h
                W
                n
                q
                120 0 0 120 20 40 cm
                /Im1 Do
                Q
                """);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.Equal(OfficeFillRule.NonZero, group.ClipPath.FillRule);
        Assert.Contains(group.ClipPath.Commands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
        Assert.Contains("<clipPath", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? raster));
        OfficeColor center = raster!.GetPixel(80, 100);
        OfficeColor clippedCorner = raster.GetPixel(25, 45);
        Assert.True(center.G > 240 && center.R < 20 && center.B < 20);
        Assert.True(clippedCorner.R > 240 && clippedCorner.G > 240 && clippedCorner.B > 240);
    }

    [Fact]
    public void RenderPage_ProjectsAnnotationNormalAppearanceStream() {
        string annotationObject = "5 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [50 70 150 110] /F 4 /AP << /N 6 0 R >> >>\nendobj";
        string appearanceObject = BuildStreamObject(
            6,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Resources << >>",
            """
            1 0 0 rg
            0 0 100 40 re
            f
            """);
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "",
            "<< >>",
            "/Annots [5 0 R]",
            annotationObject,
            appearanceObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.Red);
        Assert.Equal(50D, shape.X, 1);
        Assert.Equal(90D, shape.Y, 1);
        Assert.Equal(100D, shape.Shape.Width, 1);
        Assert.Equal(40D, shape.Shape.Height, 1);
    }

    [Fact]
    public void RenderPage_ProjectsAnnotationNormalAppearanceText() {
        string annotationObject = "5 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [50 70 170 110] /F 4 /AP << /N 6 0 R >> >>\nendobj";
        string appearanceObject = BuildStreamObject(
            6,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 120 40] /Resources << /Font << /F1 8 0 R >> >>",
            """
            BT
            /F1 14 Tf
            0 0 1 rg
            10 16 Td
            (Widget Label) Tj
            ET
            """);
        string fontObject = "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Times-BoldItalic >>\nendobj";
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "",
            "<< >>",
            "/Annots [5 0 R]",
            annotationObject,
            appearanceObject,
            fontObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Widget Label", text.Text);
        Assert.Equal(OfficeColor.Blue, text.Color);
        Assert.Equal("Times New Roman", text.Font.FamilyName);
        Assert.True(text.Font.IsBold);
        Assert.True(text.Font.IsItalic);
        Assert.Equal(60D, text.X, 1);
        Assert.Equal(100D, text.Y, 1);
        Assert.Contains("Widget Label", Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_ProjectsAnnotationNormalAppearanceInlineImage() {
        string annotationObject = "5 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [50 70 150 110] /F 4 /AP << /N 6 0 R >> >>\nendobj";
        string appearanceObject = BuildStreamObject(
            6,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Resources << >>",
            """
            q
            20 0 0 20 10 10 cm
            BI
            /W 1
            /H 1
            /CS /RGB
            /BPC 8
            ID
            abc
            EI
            Q
            """);
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "",
            "<< >>",
            "/Annots [5 0 R]",
            annotationObject,
            appearanceObject);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[] png = PdfPageImageRenderer.RenderPageAsPng(pdf);

        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(60D, image.Projection.X, 1);
        Assert.Equal(100D, image.Projection.Y, 1);
        Assert.Equal(20D, image.Projection.Width, 1);
        Assert.Equal(20D, image.Projection.Height, 1);
        AssertPngSignature(png);
    }

    [Fact]
    public void RenderPage_ProjectsNamedWidgetNormalAppearanceState() {
        string annotationObject = "5 0 obj\n<< /Type /Annot /Subtype /Widget /Rect [50 70 150 110] /F 4 /AS /Yes /AP << /N << /Off 6 0 R /Yes 7 0 R >> >> >>\nendobj";
        string offAppearance = BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Resources << >>", "");
        string yesAppearance = BuildStreamObject(
            7,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Resources << >>",
            """
            0 0.6 0 rg
            0 0 100 40 re
            f
            """);
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "",
            "<< >>",
            "/Annots [5 0 R]",
            annotationObject,
            offAppearance,
            yesAppearance);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeDrawingShape shape = Assert.Single(drawing.Shapes, item =>
            item.Shape.Kind == OfficeShapeKind.Rectangle &&
            item.Shape.FillColor == OfficeColor.FromRgb(0, 153, 0));
        Assert.Equal(50D, shape.X, 1);
        Assert.Equal(90D, shape.Y, 1);
    }

    [Fact]
    public void RenderPage_SkipsHiddenOptionalContentLayer() {
        byte[] pdf = BuildOptionalContentPdf();

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(drawing.Shapes, item => item.Shape.FillColor == OfficeColor.Red);
        Assert.Contains(drawing.Shapes, item => item.Shape.FillColor == OfficeColor.FromRgb(0, 153, 0));
        string drawingText = string.Concat(drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text));
        Assert.DoesNotContain("Hidden", drawingText, StringComparison.Ordinal);
        Assert.Contains("Visible", drawingText, StringComparison.Ordinal);
        Assert.Empty(drawing.Images);
    }

    [Fact]
    public void RenderPage_RejectsInvalidInputs() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("Page one")).ToBytes();

        Assert.Throws<ArgumentNullException>(() => PdfPageImageRenderer.RenderPage((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImageRenderer.RenderPage((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImageRenderer.RenderPage((string)null!));
        Assert.Throws<ArgumentException>(() => PdfPageImageRenderer.RenderPage(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImageRenderer.RenderPage(pdf, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImageRenderer.RenderPage(pdf, 2));
        Assert.Throws<ArgumentException>(() => PdfPageImageRenderer.RenderPage(new WriteOnlyStream()));
    }

    private static void AssertPngSignature(byte[] bytes) {
        Assert.True(bytes.Length > 8);
        Assert.Equal(137, bytes[0]);
        Assert.Equal(80, bytes[1]);
        Assert.Equal(78, bytes[2]);
        Assert.Equal(71, bytes[3]);
    }

    private static byte[] BuildSingleStreamPdf(string streamContent) =>
        BuildSingleStreamPdf(streamContent, "<< >>");

    private static byte[] BuildSingleStreamPdf(string streamContent, string resources, params string[] extraObjects) =>
        BuildSingleStreamPdfWithPageEntries(streamContent, resources, string.Empty, extraObjects);

    private static byte[] BuildSingleStreamPdfWithBinaryImageXObject(
        byte[] imageStream,
        byte[]? softMaskStream = null,
        string colorSpace = "/DeviceCMYK",
        int bitsPerComponent = 8,
        int? imageWidth = null,
        string extraImageEntries = "",
        string imageFilterEntry = "/Filter /FlateDecode",
        string? contentStream = null,
        string extraResourceEntries = "",
        string softMaskExtraEntries = "",
        string softMaskFilterEntry = "/Filter /FlateDecode",
        string[]? extraObjects = null) {
        byte[] contentStreamBytes = Encoding.ASCII.GetBytes((contentStream ?? """
            q
            20 0 0 20 40 80 cm
            /Im1 Do
            Q
            """).TrimEnd('\r', '\n'));

        using var pdf = new MemoryStream();
        WriteAscii(pdf, "%PDF-1.4\n");
        WriteAscii(pdf, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(pdf, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(pdf, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >>" + extraResourceEntries + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(pdf, "4 0 obj\n<< /Length " + contentStreamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        pdf.Write(contentStreamBytes, 0, contentStreamBytes.Length);
        WriteAscii(pdf, "\nendstream\nendobj\n");
        string softMaskReference = softMaskStream is null ? string.Empty : " /SMask 6 0 R";
        int resolvedImageWidth = imageWidth ?? (softMaskStream is null ? 2 : 1);
        string colorSpaceEntry = string.IsNullOrWhiteSpace(colorSpace) ? string.Empty : " /ColorSpace " + colorSpace;
        WriteAscii(pdf, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width " + resolvedImageWidth.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Height 1" + colorSpaceEntry + " /BitsPerComponent " + bitsPerComponent.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + imageFilterEntry + softMaskReference + extraImageEntries + " /Length " + imageStream.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        pdf.Write(imageStream, 0, imageStream.Length);
        WriteAscii(pdf, "\nendstream\nendobj\n");
        if (softMaskStream is not null) {
            WriteAscii(pdf, "6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 " + softMaskFilterEntry + softMaskExtraEntries + " /Length " + softMaskStream.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
            pdf.Write(softMaskStream, 0, softMaskStream.Length);
            WriteAscii(pdf, "\nendstream\nendobj\n");
        }

        if (extraObjects is not null) {
            for (int i = 0; i < extraObjects.Length; i++) {
                WriteAscii(pdf, extraObjects[i].TrimEnd('\r', '\n'));
                WriteAscii(pdf, "\n");
            }
        }

        WriteAscii(pdf, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return pdf.ToArray();
    }

    private static byte[] BuildSingleStreamPdfWithPageEntries(string streamContent, string resources, string pageEntries, params string[] extraObjects) {
        streamContent = streamContent.TrimEnd('\r', '\n');
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);
        string pageExtra = string.IsNullOrWhiteSpace(pageEntries) ? string.Empty : " " + pageEntries.Trim();

        string[] objects = new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources " + resources + " /Contents 4 0 R" + pageExtra + " >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent,
            "endstream",
            "endobj"
        };
        string pdf = string.Join("\n", objects.Concat(extraObjects).Concat(new[] {
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        })) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOptionalContentPdf() {
        string streamContent = """
            /OC /Hidden BDC
            1 0 0 rg
            20 80 60 40 re
            f
            BT /F1 16 Tf 1 0 0 1 20 150 Tm (Hidden) Tj ET
            q
            20 0 0 20 150 80 cm
            /Im1 Do
            Q
            EMC
            0 0.6 0 rg
            90 80 40 40 re
            f
            BT /F1 16 Tf 1 0 0 1 20 40 Tm (Visible) Tj ET
            """.TrimEnd('\r', '\n');
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        byte[] imageStream = CompressWithDeflate(new byte[] { 0, 255, 0 });

        using var pdf = new MemoryStream();
        WriteAscii(pdf, "%PDF-1.4\n");
        WriteAscii(pdf, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /OFF [7 0 R] >> >> >>\nendobj\n");
        WriteAscii(pdf, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(pdf, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Properties << /Hidden 7 0 R >> /Font << /F1 6 0 R >> /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(pdf, "4 0 obj\n<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        pdf.Write(streamBytes, 0, streamBytes.Length);
        WriteAscii(pdf, "\nendstream\nendobj\n");
        WriteAscii(pdf, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length " + imageStream.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        pdf.Write(imageStream, 0, imageStream.Length);
        WriteAscii(pdf, "\nendstream\nendobj\n");
        WriteAscii(pdf, "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
        WriteAscii(pdf, "7 0 obj\n<< /Type /OCG /Name (Hidden layer) >>\nendobj\n");
        WriteAscii(pdf, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return pdf.ToArray();
    }

    private static byte[] BuildInlineDeviceRgbImagePdf() {
        using var content = new MemoryStream();
        WriteAscii(content, "q\n20 0 0 20 40 80 cm\nBI\n/W 1\n/H 1\n/CS /RGB\n/BPC 8\nID\n");
        content.WriteByte(0);
        content.WriteByte(255);
        content.WriteByte(0);
        WriteAscii(content, "\nEI\nQ");
        byte[] contentBytes = content.ToArray();

        using var pdf = new MemoryStream();
        WriteAscii(pdf, "%PDF-1.4\n");
        WriteAscii(pdf, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(pdf, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(pdf, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(pdf, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        pdf.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(pdf, "\nendstream\nendobj\n");
        WriteAscii(pdf, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return pdf.ToArray();
    }

    private static byte[] BuildInlineNamedDeviceRgbImagePdf() {
        byte[] encoded = Encoding.ASCII.GetBytes(EncodeAscii85(CompressWithDeflate(new byte[] { 0, 255, 0 })));
        using var content = new MemoryStream();
        WriteAscii(content, "q\n20 0 0 20 40 80 cm\nBI\n/W 1\n/H 1\n/CS /CsRgb\n/BPC 8\n/F [/A85 /Fl]\nID\n");
        content.Write(encoded, 0, encoded.Length);
        WriteAscii(content, "\nEI\nQ");
        byte[] contentBytes = content.ToArray();

        using var pdf = new MemoryStream();
        WriteAscii(pdf, "%PDF-1.4\n");
        WriteAscii(pdf, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(pdf, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(pdf, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CsRgb /DeviceRGB >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(pdf, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        pdf.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(pdf, "\nendstream\nendobj\n");
        WriteAscii(pdf, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return pdf.ToArray();
    }

    private static byte[] CompressWithDeflate(byte[] input) {
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(input, 0, input.Length);
        }

        return output.ToArray();
    }

    private static byte[] CreateMinimalJpeg(int width, int height) {
        return new byte[] {
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            (byte)(height >> 8), (byte)(height & 0xFF),
            (byte)(width >> 8), (byte)(width & 0xFF),
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        };
    }

    private static string EncodeAscii85(byte[] input) {
        var builder = new StringBuilder();
        int index = 0;
        while (index + 4 <= input.Length) {
            uint value =
                ((uint)input[index] << 24) |
                ((uint)input[index + 1] << 16) |
                ((uint)input[index + 2] << 8) |
                input[index + 3];
            if (value == 0) {
                builder.Append('z');
            } else {
                AppendAscii85Tuple(builder, value, 5);
            }

            index += 4;
        }

        int remaining = input.Length - index;
        if (remaining > 0) {
            uint value = 0;
            for (int i = 0; i < remaining; i++) {
                value |= (uint)input[index + i] << (24 - i * 8);
            }

            AppendAscii85Tuple(builder, value, remaining + 1);
        }

        builder.Append("~>");
        return builder.ToString();
    }

    private static void AppendAscii85Tuple(StringBuilder builder, uint value, int count) {
        var tuple = new char[5];
        for (int i = 4; i >= 0; i--) {
            tuple[i] = (char)(value % 85 + 33);
            value /= 85;
        }

        for (int i = 0; i < count; i++) {
            builder.Append(tuple[i]);
        }
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string BuildStreamObject(int objectNumber, string dictionaryPrefix, string streamContent) {
        streamContent = streamContent.TrimEnd('\r', '\n');
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);
        return string.Join("\n", new[] {
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj",
            dictionaryPrefix + " /Length " + streamLength.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj"
        });
    }

    private sealed class WriteOnlyStream : Stream {
        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => 0;

        public override long Position {
            get => 0;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin) {
            throw new NotSupportedException();
        }

        public override void SetLength(long value) {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count) {
        }
    }

    private sealed class TestRasterImageCodec : IOfficeRasterImageCodec {
        public bool WasCalled { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            WasCalled = true;
            Assert.Equal("image/jpeg", contentType);
            image = new OfficeRasterImage(1, 1, OfficeColor.FromRgb(255, 0, 0));
            return true;
        }
    }
}
