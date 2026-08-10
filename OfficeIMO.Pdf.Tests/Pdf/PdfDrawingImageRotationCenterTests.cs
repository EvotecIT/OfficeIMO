using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingImageRotationCenterTests {
    private static readonly byte[] Png = PdfPngTestImages.CreateRgbPng(1, 1);

    [Fact]
    public void DrawingImageUsesCustomRotationCenterInPdfTransform() {
        string centered = RenderImageTransform(rotationCenterX: 60D, rotationCenterY: 60D);
        string custom = RenderImageTransform(rotationCenterX: 50D, rotationCenterY: 52D);

        Assert.NotEqual(centered, custom);
    }

    [Fact]
    public void DrawingImageClockwiseRotationIsConvertedToPdfCoordinates() {
        string matrix = RenderImageTransform(rotationCenterX: 60D, rotationCenterY: 60D, rotationDegrees: 90D);
        string[] values = matrix.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Equal(0D, double.Parse(values[0], CultureInfo.InvariantCulture), 6);
        Assert.Equal(-40D, double.Parse(values[1], CultureInfo.InvariantCulture), 6);
        Assert.Equal(40D, double.Parse(values[2], CultureInfo.InvariantCulture), 6);
        Assert.Equal(0D, double.Parse(values[3], CultureInfo.InvariantCulture), 6);
    }

    private static string RenderImageTransform(double rotationCenterX, double rotationCenterY, double rotationDegrees = 30D) {
        var projection = new OfficeImageProjection(
            new OfficeImagePlacement(40D, 40D, 40D, 40D),
            rotationDegrees: rotationDegrees,
            rotationCenterX: rotationCenterX,
            rotationCenterY: rotationCenterY);
        var drawing = new OfficeDrawing(120D, 120D)
            .AddImage(Png, "image/png", projection, "Rotated marker");
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(drawing)
            .ToBytes();
        string content = Encoding.ASCII.GetString(pdf);
        Match match = Regex.Match(
            content,
            @"(?<matrix>-?\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+cm)\s*/Im\d+\s+Do",
            RegexOptions.CultureInvariant);

        Assert.True(match.Success, "Expected an image transform matrix in the PDF content stream.");
        return match.Groups["matrix"].Value;
    }
}
