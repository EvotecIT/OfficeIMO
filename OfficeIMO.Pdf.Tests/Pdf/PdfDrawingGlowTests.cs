using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingGlowTests {
    [Theory]
    [InlineData("pdf", (byte)255)]
    [InlineData("pdf", (byte)128)]
    [InlineData("raster", (byte)255)]
    [InlineData("raster", (byte)128)]
    [InlineData("svg", (byte)255)]
    [InlineData("svg", (byte)128)]
    public void FullOpacityGlowFadesOutwardAcrossRenderers(string renderer, byte alpha) {
        OfficeShape shape = OfficeShape.Rectangle(48D, 28D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        shape.Glow = new OfficeGlow(OfficeColor.FromRgba(255, 0, 255, alpha), 1D, 6D);
        var drawing = new OfficeDrawing(120D, 90D).AddShape(shape, 32D, 24D);

        OfficeRasterImage raster = RenderGlow(drawing, renderer);

        OfficeColor inner = raster.GetPixel(31, 38);
        OfficeColor outer = raster.GetPixel(27, 38);
        Assert.Equal(255, inner.R);
        Assert.Equal(255, inner.B);
        int expectedGreen = 255 - alpha;
        Assert.InRange((int)inner.G, System.Math.Max(0, expectedGreen - 3), expectedGreen + 3);
        Assert.InRange((int)outer.G, inner.G + 20, 254);
        Assert.Equal(OfficeColor.White, raster.GetPixel(24, 38));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(48, 38));
    }

    [Theory]
    [InlineData("pdf")]
    [InlineData("raster")]
    [InlineData("svg")]
    public void RectangleGlowHasRoundNativeCornerBounds(string renderer) {
        OfficeShape shape = OfficeShape.Rectangle(48D, 28D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
        shape.Glow = new OfficeGlow(OfficeColor.FromRgb(255, 0, 255), 1D, 6D);
        var drawing = new OfficeDrawing(120D, 90D).AddShape(shape, 32D, 24D);

        OfficeRasterImage raster = RenderGlow(drawing, renderer);

        // Outside the radius-six corner disk, but inside a square/miter expansion.
        Assert.Equal(OfficeColor.White, raster.GetPixel(26, 18));
        OfficeColor corner = raster.GetPixel(29, 21);
        Assert.True(corner.R == 255 && corner.B == 255 && corner.G < 250);
        Assert.Equal(OfficeStrokeLineJoin.Miter, shape.StrokeLineJoin);
    }

    [Theory]
    [InlineData("pdf")]
    [InlineData("raster")]
    [InlineData("svg")]
    public void AcuteJoinAndOpenCapGlowStayWithinRoundRadius(string renderer) {
        OfficeShape angle = OfficeShape.Path(40D, 40D,
            OfficePathCommand.MoveTo(10D, 30D), OfficePathCommand.LineTo(20D, 5D), OfficePathCommand.LineTo(30D, 30D));
        angle.StrokeColor = OfficeColor.Blue;
        angle.StrokeWidth = 2D;
        angle.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
        angle.Glow = new OfficeGlow(OfficeColor.FromRgb(255, 0, 255), 1D, 6D);
        OfficeShape line = OfficeShape.Line(0D, 0D, 20D, 0D);
        line.StrokeColor = OfficeColor.Blue;
        line.StrokeWidth = 2D;
        line.StrokeLineCap = OfficeStrokeLineCap.Square;
        line.Glow = angle.Glow.Clone();
        var drawing = new OfficeDrawing(140D, 100D).AddShape(angle, 30D, 25D).AddShape(line, 90D, 60D);

        OfficeRasterImage raster = RenderGlow(drawing, renderer);

        // The apex is (50,30); the halo radius, including half the source stroke, is seven.
        Assert.Equal(OfficeColor.White, raster.GetPixel(50, 20));
        Assert.True(raster.GetPixel(50, 24).G < 250, "The acute vertex must retain its round halo.");
        Assert.Equal(OfficeColor.White, raster.GetPixel(84, 54));
        Assert.True(raster.GetPixel(86, 60).G < 250, "The open endpoint must retain its round cap.");
    }

    [Theory]
    [InlineData((byte)255)]
    [InlineData((byte)128)]
    public void DrawingGlowRetainsVectorHaloAlphaAndPaintOrder(byte alpha) {
        OfficeShape shape = OfficeShape.Rectangle(48D, 28D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        shape.StrokeOpacity = 0D;
        shape.Glow = new OfficeGlow(OfficeColor.FromRgba(255, 0, 255, alpha), 0.5D, 6D);
        shape.Shadow = new OfficeShadow(OfficeColor.Black, 1D, 10D, 0D);
        OfficeShape following = OfficeShape.Rectangle(12D, 12D);
        following.FillColor = OfficeColor.Green;
        following.StrokeWidth = 0D;
        var drawing = new OfficeDrawing(120D, 90D)
            .AddShape(shape, 32D, 24D)
            .AddShape(following, 96D, 64D);

        byte[] bytes = Write(drawing);
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Empty(PdfDocument.Load(bytes).Reader.Images());
        OfficeRasterImage raster = ReadRaster(bytes);
        OfficeColor halo = raster.GetPixel(31, 38);
        Assert.Equal(255, halo.R);
        Assert.Equal(255, halo.B);
        int expectedGreen = (int)System.Math.Round(255D - 0.5D * alpha);
        Assert.InRange((int)halo.G, expectedGreen - 3, expectedGreen + 3);
        OfficeColor outerHalo = raster.GetPixel(27, 38);
        Assert.True(outerHalo.G > halo.G && outerHalo.G < 255, "The halo must fade away from the contour.");
        Assert.Equal(OfficeColor.White, raster.GetPixel(24, 38));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(48, 38));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(83, 38));
        Assert.Equal(OfficeColor.Green, raster.GetPixel(100, 68));
        Assert.Equal(0D, shape.StrokeOpacity);
        Assert.Equal(alpha, shape.Glow!.Color.A);
    }

    [Fact]
    public void GradientStrokeDoesNotPaintTheGlowTwice() {
        OfficeShape solid = OfficeShape.Rectangle(40D, 20D);
        solid.FillColor = OfficeColor.Blue;
        solid.StrokeColor = OfficeColor.Red;
        solid.StrokeWidth = 2D;
        solid.Glow = new OfficeGlow(OfficeColor.FromRgb(255, 0, 255), 0.5D, 6D);
        OfficeShape gradient = solid.Clone();
        gradient.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Green);
        var drawing = new OfficeDrawing(100D, 100D)
            .AddShape(solid, 30D, 20D)
            .AddShape(gradient, 30D, 65D);

        OfficeRasterImage raster = ReadRaster(Write(drawing));
        OfficeColor solidHalo = raster.GetPixel(27, 30);
        OfficeColor gradientHalo = raster.GetPixel(27, 75);
        Assert.True(solidHalo.R == 255 && solidHalo.B == 255 && solidHalo.G < 240);
        Assert.Equal(solidHalo, gradientHalo);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(50, 75));
        Assert.NotNull(gradient.StrokeGradient);
        Assert.NotNull(gradient.Glow);
    }

    [Fact]
    public void OpenPathGlowUsesLocalTransformAndClipWithoutInheritingDashes() {
        OfficeShape shape = OfficeShape.Path(30D, 20D,
            OfficePathCommand.MoveTo(5D, 10D), OfficePathCommand.LineTo(25D, 10D));
        shape.StrokeColor = OfficeColor.Blue;
        shape.StrokeWidth = 2D;
        shape.SetStrokeDashArray(new[] { 2D, 8D }, 1D);
        shape.Transform = OfficeTransform.Translate(10D, 5D);
        shape.ClipPath = OfficeClipPath.Rectangle(20D, 20D);
        shape.Glow = new OfficeGlow(OfficeColor.FromRgb(255, 0, 255), 0.5D, 6D);
        var drawing = new OfficeDrawing(100D, 80D).AddShape(shape, 20D, 20D);

        OfficeRasterImage raster = ReadRaster(Write(drawing));
        for (int x = 36; x < 49; x++) {
            OfficeColor halo = raster.GetPixel(x, 31);
            Assert.True(halo.R == 255 && halo.B == 255 && halo.G < 250, "Glow must form a continuous contour around a dashed path.");
        }
        Assert.Equal(OfficeColor.White, raster.GetPixel(51, 31));
        Assert.Equal(OfficeColor.White, raster.GetPixel(40, 24));
        Assert.Equal(new[] { 2D, 8D }, shape.StrokeDashArray);
        Assert.Equal(1D, shape.StrokeDashOffset);
    }

    [Theory]
    [InlineData("radius")]
    [InlineData("opacity")]
    [InlineData("alpha")]
    [InlineData("clip")]
    public void DisabledOrEmptyClippedGlowDoesNotPaintAHalo(string disabled) {
        OfficeShape shape = OfficeShape.Rectangle(40D, 20D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        shape.Glow = new OfficeGlow(OfficeColor.FromRgba(255, 0, 255, disabled == "alpha" ? (byte)0 : (byte)255),
            disabled == "opacity" ? 0D : 0.5D, disabled == "radius" ? 0D : 6D);
        if (disabled == "clip") shape.ClipPath = OfficeClipPath.Empty();

        OfficeShape background = OfficeShape.Rectangle(100D, 80D);
        background.FillColor = OfficeColor.White;
        background.StrokeWidth = 0D;
        OfficeRasterImage raster = ReadRaster(Write(new OfficeDrawing(100D, 80D)
            .AddShape(background, 0D, 0D).AddShape(shape, 30D, 25D)));
        Assert.Equal(OfficeColor.White, raster.GetPixel(29, 35));
        Assert.Equal(disabled == "clip" ? OfficeColor.White : OfficeColor.Blue, raster.GetPixel(50, 35));
    }

    private static OfficeRasterImage ReadRaster(byte[] bytes) =>
        OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes), background: OfficeColor.White);

    private static OfficeRasterImage RenderGlow(OfficeDrawing drawing, string renderer) {
        if (renderer == "pdf") return ReadRaster(Write(drawing));
        if (renderer == "raster") return OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 1D, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg), out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        return OfficeDrawingRasterRenderer.Render(imported!, background: OfficeColor.White);
    }

    private static byte[] Write(OfficeDrawing drawing) => PdfDocument.Create(new PdfCreateOptions {
        PageWidth = drawing.Width, PageHeight = drawing.Height,
        MarginLeft = 0D, MarginTop = 0D, MarginRight = 0D, MarginBottom = 0D,
        CompressContentStreams = false
    }).Drawing(drawing).ToBytes();
}
