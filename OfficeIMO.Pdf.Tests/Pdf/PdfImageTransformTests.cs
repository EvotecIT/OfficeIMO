using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageTransformTests {
    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(37)]
    public void ProportionalTransformPreservesRotationAndUnselectedPlacement(double rotation) {
        byte[] source = PdfDocument.Create().Paragraph(text => text.Text("Keep this text")).ToBytes();
        byte[] image = PdfPngTestImages.CreateRgbPng(255, 0, 0);
        source = PdfStamper.StampImage(source, image, new PdfImageStampOptions {
            PageNumbers = new[] { 1 }, X = 100, Y = 180, Width = 80, Height = 40, RotationDegrees = rotation
        });
        PdfDocument document = PdfDocument.Load(source).Images.Add(new PdfPageRegion(1, 300, 300, 30, 30), image).Document;
        PdfImagePlacement before = document.Images.Placements()[0];
        PdfImageEditResult result = document.Images.Transform(before, 10, -20, 1.5);
        PdfImagePlacement after = result.Document.Images.Placements().Single(placement => Math.Abs(placement.X - 300) > 0.01);
        Assert.InRange(Math.Abs(before.A * 1.5 - after.A), 0, 0.002);
        Assert.InRange(Math.Abs(before.B * 1.5 - after.B), 0, 0.002);
        Assert.InRange(Math.Abs(before.C * 1.5 - after.C), 0, 0.002);
        Assert.InRange(Math.Abs(before.D * 1.5 - after.D), 0, 0.002);
        Assert.InRange(Math.Abs(before.E + (before.A + before.C) / 2 + 10 - after.E - (after.A + after.C) / 2), 0, 0.002);
        Assert.InRange(Math.Abs(before.F + (before.B + before.D) / 2 - 20 - after.F - (after.B + after.D) / 2), 0, 0.002);
        Assert.Contains("Keep this text", result.Document.Read().Text);
        Assert.Contains(result.Document.Images.Placements(), placement => Math.Abs(placement.X - 300) < 0.01 && Math.Abs(placement.Width - 30) < 0.01);
        Assert.Throws<InvalidOperationException>(() => result.Document.Images.Transform(before, 0, 0, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Images.Transform(before, 0, 0, 0));
    }
}
