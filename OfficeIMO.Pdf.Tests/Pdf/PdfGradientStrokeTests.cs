using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGradientStrokeTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ZeroAxisLineBoundsRemainValidForNativeGradientStrokes(bool vertical) {
        OfficeShape shape = vertical ? OfficeShape.Line(0, 0, 0, 60) : OfficeShape.Line(0, 0, 60, 0);
        shape.StrokeGradient = vertical ? OfficeLinearGradient.Vertical(OfficeColor.Red, OfficeColor.Blue)
            : OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeColor = null;
        shape.StrokeWidth = 8;
        byte[] bytes = Export(shape);
        Assert.Contains("/ShadingType 2", Encoding.ASCII.GetString(bytes));
        Assert.Equal(0D, vertical ? shape.Width : shape.Height);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void VerticalAndOffCenterRadialFieldsRetainTheirVisualDirection(bool radial) {
        OfficeShape shape = VerticalPath();
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(.375, .15, 0, .375, .15, .7,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        else shape.StrokeGradient = OfficeLinearGradient.Vertical(OfficeColor.Red, OfficeColor.Blue);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(Export(shape)));
        OfficeColor top = image.GetPixel(30, 20), bottom = image.GetPixel(30, 90);
        Assert.True(top.R > top.B + 100);
        Assert.True(bottom.B > bottom.R + 20);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeaderAndFooterUseTheSameNativeGradientStrokeProjection(bool footer) {
        OfficeShape shape = VerticalPath();
        shape.StrokeGradient = OfficeLinearGradient.Vertical(OfficeColor.Red, OfficeColor.Blue);
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 200, PageHeight = 400, MarginTop = 140, MarginBottom = 140,
            MarginLeft = 20, MarginRight = 20
        });
        byte[] bytes = document.Compose(compose => compose.Page(page => {
            if (footer) page.Footer(content => content.Shape(shape));
            else page.Header(content => content.Shape(shape));
            page.Content(content => content.Spacer(20));
        })).ToBytes();
        string syntax = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/ShadingType 2", syntax);
        Assert.DoesNotContain("/Subtype /Image", syntax);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GradientStopAlphaMultipliesStrokeOpacityWithoutRasterFallback(bool fade) {
        OfficeShape shape = OfficeShape.Path(80, 40, OfficePathCommand.MoveTo(10, 20), OfficePathCommand.LineTo(70, 20));
        shape.StrokeColor = null;
        shape.StrokeWidth = 12;
        shape.StrokeOpacity = .5;
        shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.FromRgba(255, 0, 0, fade ? (byte)0 : (byte)128),
            OfficeColor.FromRgba(255, 0, 0, fade ? (byte)255 : (byte)128));
        byte[] bytes = Export(shape);
        string syntax = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/SMask", syntax);
        Assert.DoesNotContain("/Subtype /Image", syntax);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        if (fade) {
            Assert.InRange(image.GetPixel(20, 20).A, (byte)25, (byte)40);
            Assert.InRange(image.GetPixel(60, 20).A, (byte)90, (byte)105);
        } else Assert.InRange(image.GetPixel(40, 20).A, (byte)62, (byte)66);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RadialFillAlphaMaskTracksThePaintTransform(bool transformed) {
        OfficeShape shape = OfficeShape.Rectangle(80, 40);
        shape.StrokeWidth = 0;
        shape.FillOpacity = .5;
        shape.FillRadialGradient = new OfficeRadialGradient(.25, .25, 0, .25, .25, .5,
            new OfficeGradientStop(0, OfficeColor.FromRgba(255, 0, 0, 0)),
            new OfficeGradientStop(1, OfficeColor.Red));
        if (transformed) shape.Transform = OfficeTransform.Scale(1.1, .8);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(Export(shape)));
        Assert.InRange(image.GetPixel(transformed ? 22 : 20, transformed ? 8 : 10).A, (byte)0, (byte)12);
        Assert.InRange(image.GetPixel(transformed ? 66 : 60, transformed ? 8 : 10).A, (byte)115, (byte)128);
    }

    private static OfficeShape VerticalPath() {
        OfficeShape shape = OfficeShape.Path(80, 120, OfficePathCommand.MoveTo(30, 10), OfficePathCommand.LineTo(30, 100));
        shape.FillColor = null;
        shape.StrokeColor = null;
        shape.StrokeWidth = 16;
        return shape;
    }

    private static byte[] Export(OfficeShape shape) => PdfDocument.Create(new PdfOptions {
        PageWidth = 120, PageHeight = 160, MarginTop = 0, MarginBottom = 0, MarginLeft = 0, MarginRight = 0
    }).Compose(c => c.Page(p => p.Content(content => content.Shape(shape)))).ToBytes();

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void GradientOnlyStrokeProducesNativeShadingAndPreservesSource(bool transformed, bool radial) {
        var shape=OfficeShape.Path(80,40,OfficePathCommand.MoveTo(10,30),OfficePathCommand.LineTo(40,10),OfficePathCommand.LineTo(70,30));
        var gradient=OfficeLinearGradient.Horizontal(OfficeColor.Red,OfficeColor.Blue);
        var radialGradient = new OfficeRadialGradient(.5, .5, 0, .5, .5, .5,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        shape.FillColor=null; shape.StrokeColor=null; shape.StrokeWidth=4;
        if (radial) shape.StrokeRadialGradient = radialGradient;
        else shape.StrokeGradient = gradient;
        shape.StrokeOpacity=.5; shape.StrokeLineCap=OfficeStrokeLineCap.Round;
        shape.SetStrokeDashArray(new[]{6D,3D},2D);
        if(transformed)shape.Transform=OfficeTransform.Scale(1.2,.8);
        byte[] bytes=PdfDocument.Create(new PdfOptions{PageWidth=120,PageHeight=80,MarginTop=0,MarginLeft=0,MarginRight=0,MarginBottom=0})
            .Compose(c=>c.Page(p=>p.Content(content=>content.Shape(shape)))).ToBytes();
        string syntax=Encoding.ASCII.GetString(bytes);
        Assert.Contains(radial ? "/ShadingType 3" : "/ShadingType 2",syntax);
        Assert.Contains("/ca 0.5",syntax);
        Assert.DoesNotContain("/Subtype /Image",syntax);
        if (radial) Assert.Same(radialGradient, shape.StrokeRadialGradient);
        else Assert.Same(gradient,shape.StrokeGradient);
        Assert.Null(shape.FillColor); Assert.Null(shape.StrokeColor);
    }
}
