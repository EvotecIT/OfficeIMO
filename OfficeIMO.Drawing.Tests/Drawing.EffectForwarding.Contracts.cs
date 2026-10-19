using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingEffectForwardingContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TranslatedImageSamplesFractionalEdgeOnce(bool cropped) {
        var source = new OfficeRasterImage(2, 1, OfficeColor.Red);
        source.SetPixel(1, 0, OfficeColor.Blue);
        byte[] png = OfficePngWriter.Encode(source);
        OfficeImageSourceCrop crop = cropped ? new OfficeImageSourceCrop(0D, 0D, 0.5D, 0D) : default;
        var localProjection = new OfficeImageProjection(new OfficeImagePlacement(0, 0, 7.5D, 4D), crop);
        var layer = new OfficeDrawing(7.5D, 4D).AddImageWithInterpolation(png, "image/png", localProjection, interpolate: false);
        var actual = new OfficeDrawing(40, 8).AddEffectDrawing(layer, OfficeTransform.Translate(30.5D, 2D));
        var control = new OfficeDrawing(40, 8).AddImageWithInterpolation(png, "image/png",
            localProjection.Translate(30.5D, 2D), interpolate: false);

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(actual);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(control).GetPixels(), rendered.GetPixels());
        Assert.Equal(cropped ? OfficeColor.Red : OfficeColor.Blue, rendered.GetPixel(37, 3));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ForwardedTextKeepsAuthoredTrackingAndOpticalSize(bool opticalSize) {
        var source = new OfficeDrawing(100, 30);
        if (opticalSize) {
            source.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 700F };
            source.Fonts.Add("Probe", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf")));
        } else {
            byte[] tracking = ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { -1D, 1D },
                new[] { new short[] { -120, -40 }, new short[] { -80, 40 } });
            source.Fonts.Add("Probe", ManagedTextShapingTestAssets.CreateTrackingFont(tracking));
        }
        source.AddText("AB", 2, 2, 90, 25, new OfficeFontInfo("Probe", 15), OfficeColor.Black);
        // A real nearest image makes the surrounding uniform transform eligible
        // for forwarding; ordinary scaled rendering remains the independent control.
        source.AddImageWithInterpolation(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue)), "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(95, 0, 1, 1)), interpolate: false);
        var forwarded = new OfficeDrawing(200, 60).AddEffectDrawing(source, OfficeTransform.Scale(2, 2));

        Assert.Equal(OfficeDrawingRasterRenderer.Render(source, 2D).GetPixels(),
            OfficeDrawingRasterRenderer.Render(forwarded).GetPixels());
    }

    [Theory]
    [InlineData(0.5D, false)]
    [InlineData(0.5D, true)]
    [InlineData(1D, false)]
    [InlineData(1D, true)]
    public void InheritedClipCoversOverlappingForwardedPaintOnce(double left, bool polygon) {
        var source = new OfficeDrawing(2, 1);
        AddFill(source, OfficeColor.Red);
        AddFill(source, OfficeColor.Blue);
        source.AddImageWithInterpolation(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black)), "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(1, 0, 1, 1)), interpolate: false);
        var enlarged = new OfficeDrawing(4, 2).AddEffectDrawing(source, OfficeTransform.Scale(2, 2));
        OfficeClipPath clip = polygon ? OfficeClipPath.Path(new[] {
            OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(4 - left, 0),
            OfficePathCommand.LineTo(4 - left, 2), OfficePathCommand.LineTo(0, 2), OfficePathCommand.Close()
        }) : OfficeClipPath.Rectangle(4 - left, 2);
        var drawing = new OfficeDrawing(4, 2).AddClippedDrawing(enlarged, left, 0, clip, -left, 0);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { PdfRasterConventions = true });

        // Composite the source first, then apply the inherited fractional clip once.
        OfficeRasterImage flattened = OfficeDrawingRasterRenderer.Render(source);
        var expected = new OfficeRasterImage(4, 2);
        var canvas = new OfficeRasterCanvas(expected) { AntialiasedClipping = true };
        using (canvas.PushClipRectangle(left, 0, 4 - left, 2)) {
            canvas.DrawAffineImage(flattened, OfficeTransform.Scale(2, 2), 1D, interpolate: false);
        }
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
        if (left == 0.5D) Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), actual.GetPixel(0, 0));
    }

    private static void AddFill(OfficeDrawing source, OfficeColor color) {
        OfficeShape rectangle = OfficeShape.Rectangle(2, 1);
        rectangle.FillColor = color;
        rectangle.StrokeColor = null;
        source.AddShape(rectangle, 0, 0);
    }
}
