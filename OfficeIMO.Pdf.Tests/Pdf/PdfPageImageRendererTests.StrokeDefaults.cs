using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(1, OfficeStrokeLineJoin.Round)]
    [InlineData(2, OfficeStrokeLineJoin.Bevel)]
    public void RenderPage_StrokeDefaults_RestoresImplicitJoinAfterAuthoredJoin(int join, OfficeStrokeLineJoin expected) {
        const string path = "10 60 m 30 100 l 50 60 l S";
        byte[] pdf = BuildSingleStreamPdf(FormattableString.Invariant(
            $"0 J 10 w 1 0 0 RG {path} q {join} j 1 0 0 1 70 0 cm {path} Q q 1 0 0 1 140 0 cm {path} Q"));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Collection(drawing.Shapes,
            first => Assert.Equal(OfficeStrokeLineJoin.Miter, first.Shape.StrokeLineJoin),
            authored => Assert.Equal(expected, authored.Shape.StrokeLineJoin),
            restored => Assert.Equal(OfficeStrokeLineJoin.Miter, restored.Shape.StrokeLineJoin));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(30, 92));
        Assert.Equal(OfficeColor.White, raster.GetPixel(100, 92));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(170, 92));
    }

    [Theory]
    [InlineData("", OfficeStrokeLineJoin.Miter)]
    [InlineData("2 j", OfficeStrokeLineJoin.Bevel)]
    public void RenderPage_StrokeDefaults_PreservesFormCallerJoins(string callerJoin, OfficeStrokeLineJoin expected) {
        const string stroke = "0 J 20 w 1 0 0 RG 40 60 m 100 140 l 160 60 l S";
        const string resources = "<< /XObject << /Fm1 5 0 R >> >>";
        const string dictionary = "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << >>";
        byte[] inherited = BuildSingleStreamPdf(callerJoin + " /Fm1 Do", resources,
            BuildStreamObject(5, dictionary, stroke));
        byte[] local = BuildSingleStreamPdf("/Fm1 Do", resources,
            BuildStreamObject(5, dictionary, FormattableString.Invariant($"{(int)expected} j ") + stroke));
        OfficeRasterImage inheritedRaster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(inherited), background: OfficeColor.White);
        OfficeRasterImage localRaster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(local), background: OfficeColor.White);

        Assert.Equal(expected == OfficeStrokeLineJoin.Miter ? OfficeColor.Red : OfficeColor.White,
            inheritedRaster.GetPixel(100, 46));
        Assert.Equal(OfficePngWriter.Encode(localRaster), OfficePngWriter.Encode(inheritedRaster));
    }

    [Fact]
    public void RenderPage_StrokeDefaults_UsesPdfInitialMiterLimitForPatternStrokeMask() {
        byte[] pdf = BuildSingleStreamPdf("0 J 0 j /Pattern CS /P1 SCN 20 w 80 20 m 100 120 l 120 20 l S",
            "<< /Pattern << /P1 5 0 R >> >>",
            BuildStreamObject(5,
                "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 20 20] /XStep 20 /YStep 20 /Resources << >>",
                "1 0 0 rg 0 0 20 20 re f"));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(pdf), background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(100, 40));
        Assert.Equal(OfficeColor.White, raster.GetPixel(100, 20));
    }

    [Fact]
    public void RenderPage_StrokeDefaults_UsesPdfInitialMiterLimitForAcutePath() {
        byte[] pdf = BuildSingleStreamPdf("0 J 0 j 20 w 1 0 0 RG 80 20 m 100 120 l 120 20 l S");
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(100, 40));
        Assert.Equal(OfficeColor.White, raster.GetPixel(100, 20));
        Assert.Equal(10D, Assert.Single(drawing.Shapes).Shape.StrokeMiterLimit);
    }

    [Fact]
    public void RenderPage_StrokeDefaults_UsesPdfDefaultMiterJoinWhenOperatorOmitted() {
        const string stroke = "0 J 20 w 1 0 0 RG 40 60 m 100 140 l 160 60 l S";
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(stroke));
        OfficeRasterImage implicitJoin = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        OfficeRasterImage explicitJoin = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf("0 j " + stroke)), background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, implicitJoin.GetPixel(100, 46));
        Assert.Equal(OfficePngWriter.Encode(explicitJoin), OfficePngWriter.Encode(implicitJoin));
        Assert.Equal(OfficeStrokeLineJoin.Miter, Assert.Single(drawing.Shapes).Shape.StrokeLineJoin);
    }

    [Fact]
    public void RenderPage_UsesPdfDefaultButtCapWhenJOmitted() {
        const string stroke = "20 w 1 0 0 RG 40 100 m 120 100 l S";
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(stroke));
        OfficeRasterImage implicitCap = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        OfficeRasterImage explicitCap = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf("0 J " + stroke)), background: OfficeColor.White);

        Assert.Equal(OfficeColor.White, implicitCap.GetPixel(35, 100));
        Assert.Equal(OfficeColor.White, implicitCap.GetPixel(125, 100));
        Assert.Equal(OfficeColor.Red, implicitCap.GetPixel(40, 100));
        Assert.Equal(OfficePngWriter.Encode(explicitCap), OfficePngWriter.Encode(implicitCap));
        Assert.Equal(OfficeStrokeLineCap.Butt, Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>()).Shape.StrokeLineCap);
    }

    [Theory]
    [InlineData(1, OfficeStrokeLineCap.Round)]
    [InlineData(2, OfficeStrokeLineCap.Square)]
    public void RenderPage_StrokeDefaults_RestoresImplicitCapAfterAuthoredCap(int cap, OfficeStrokeLineCap expected) {
        byte[] pdf = BuildSingleStreamPdf(FormattableString.Invariant($"""
            20 w 1 0 0 RG
            40 150 m 120 150 l S
            q {cap} J 40 100 m 120 100 l S Q
            40 50 m 120 50 l S
            """));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Collection(drawing.Shapes,
            first => Assert.Equal(OfficeStrokeLineCap.Butt, first.Shape.StrokeLineCap),
            authored => Assert.Equal(expected, authored.Shape.StrokeLineCap),
            restored => Assert.Equal(OfficeStrokeLineCap.Butt, restored.Shape.StrokeLineCap));
        Assert.Equal(OfficeColor.White, raster.GetPixel(35, 50));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 100));
        Assert.Equal(OfficeColor.White, raster.GetPixel(35, 150));
        Assert.Equal(expected == OfficeStrokeLineCap.Square ? OfficeColor.Red : OfficeColor.White,
            raster.GetPixel(31, 91));
    }

    [Theory]
    [InlineData("", OfficeStrokeLineCap.Butt, false)]
    [InlineData("1 J", OfficeStrokeLineCap.Round, false)]
    [InlineData("2 J", OfficeStrokeLineCap.Square, false)]
    [InlineData("", OfficeStrokeLineCap.Butt, true)]
    [InlineData("1 J", OfficeStrokeLineCap.Round, true)]
    [InlineData("2 J", OfficeStrokeLineCap.Square, true)]
    public void RenderPage_StrokeDefaults_PreservesFormCallerCaps(string callerCap, OfficeStrokeLineCap expected, bool transparency) {
        const string stroke = "20 w 1 0 0 RG 40 100 m 120 100 l S";
        const string resources = "<< /XObject << /Fm1 5 0 R >> >>";
        string dictionary = "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources << >>"
            + (transparency ? " /Group << /Type /Group /S /Transparency /I true /CS /DeviceRGB >>" : "");
        byte[] inherited = BuildSingleStreamPdf(callerCap + " /Fm1 Do", resources,
            BuildStreamObject(5, dictionary, stroke));
        byte[] local = BuildSingleStreamPdf("/Fm1 Do", resources,
            BuildStreamObject(5, dictionary, FormattableString.Invariant($"{(int)expected} J ") + stroke));
        OfficeRasterImage inheritedRaster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(inherited), background: OfficeColor.White);
        OfficeRasterImage localRaster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(local), background: OfficeColor.White);

        Assert.Equal(expected == OfficeStrokeLineCap.Butt ? OfficeColor.White : OfficeColor.Red,
            inheritedRaster.GetPixel(35, 100));
        Assert.Equal(OfficeColor.Red, inheritedRaster.GetPixel(50, 100));
        Assert.Equal(OfficePngWriter.Encode(localRaster), OfficePngWriter.Encode(inheritedRaster));
    }

    [Fact]
    public void RenderPage_StrokeDefaults_ExtGStateWithoutCapRetainsAuthoredCap() {
        byte[] pdf = BuildSingleStreamPdf("1 J /GS1 gs 1 0 0 RG 40 100 m 120 100 l S",
            "<< /ExtGState << /GS1 5 0 R >> >>",
            "5 0 obj\n<< /Type /ExtGState /LW 20 >>\nendobj");
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Equal(OfficeStrokeLineCap.Round, Assert.Single(drawing.Shapes).Shape.StrokeLineCap);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 100));
        Assert.Equal(OfficeColor.White, raster.GetPixel(31, 91));
    }

    [Fact]
    public void RenderPage_StrokeDefaults_UsesButtCapForTilingPatternStrokeMask() {
        const string stroke = "/Pattern CS /P1 SCN 20 w 40 100 m 120 100 l S";
        const string resources = "<< /Pattern << /P1 5 0 R >> >>";
        string pattern = BuildStreamObject(5,
            "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 20 20] /XStep 20 /YStep 20 /Resources << >>",
            "1 0 0 rg 0 0 20 20 re f");
        OfficeRasterImage implicitCap = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(stroke, resources, pattern)), background: OfficeColor.White);
        OfficeRasterImage explicitCap = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf("0 J " + stroke, resources, pattern)), background: OfficeColor.White);

        Assert.Equal(OfficeColor.White, implicitCap.GetPixel(35, 100));
        Assert.Equal(OfficeColor.White, implicitCap.GetPixel(125, 100));
        Assert.Equal(OfficeColor.Red, implicitCap.GetPixel(50, 100));
        Assert.Equal(OfficePngWriter.Encode(explicitCap), OfficePngWriter.Encode(implicitCap));
    }
}
