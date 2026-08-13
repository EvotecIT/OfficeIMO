using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasRotatedImageClipTests {
    private static readonly byte[] Png = PdfPngTestImages.CreateRgbPng(1, 1);

    [Fact]
    public void CanvasClip_KeepsDirectDrawingImagesWhenOnlyTheRotatedFootprintIntersects() {
        var drawing = new OfficeDrawing(60D, 50D)
            .AddImage(
                Png,
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(10D, 20D, 40D, 10D), rotationDegrees: 45D));

        byte[] bytes = RenderClippedDrawing(drawing);

        Assert.Contains("/Im1 Do", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasClip_KeepsImagesWhenOnlyTheRotatedDrawingFootprintIntersects() {
        var drawing = new OfficeDrawing(60D, 50D)
            .AddImage(
                Png,
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(10D, 20D, 40D, 10D)));

        byte[] bytes = RenderClippedDrawing(drawing, rotationAngle: 45D);

        Assert.Contains("/Im1 Do", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    private static byte[] RenderClippedDrawing(OfficeDrawing drawing, double rotationAngle = 0D) =>
        PdfDocument.Create(new PdfOptions {
                PageWidth = 120D,
                PageHeight = 120D,
                MarginLeft = 0D,
                MarginRight = 0D,
                MarginTop = 0D,
                MarginBottom = 0D,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(0D, 18D, 120D, 2D, clipped =>
                clipped.Drawing(drawing, 10D, 10D, 60D, 50D, rotationAngle: rotationAngle)))
            .ToBytes();
}
