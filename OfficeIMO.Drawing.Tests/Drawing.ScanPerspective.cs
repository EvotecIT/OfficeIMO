using OfficeIMO.Drawing;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ScanPerspectiveTests {
    [Fact]
    public void IdentityRetainsPixelsAndOwnsItsBuffer() {
        var source = new OfficeRasterImage(40, 30, OfficeColor.White);
        source.SetPixel(12, 14, OfficeColor.Black);
        var result = OfficeScanProcessor.CorrectPerspective(source, new OfficeScanPerspectiveOptions());
        Assert.Equal(source.GetPixels(), result.Image.GetPixels());
        result.Image.SetPixel(12, 14, OfficeColor.White);
        Assert.Equal(OfficeColor.Black, source.GetPixel(12, 14));
    }

    [Fact]
    public void PerspectiveCornersAndInteriorRoundTripToOriginalPixels() {
        var options = new OfficeScanPerspectiveOptions {
            TopLeft = new OfficePoint(.2, .1),
            TopRight = new OfficePoint(.8, .15),
            BottomRight = new OfficePoint(.95, .9),
            BottomLeft = new OfficePoint(.05, .85)
        };
        var source = new OfficeRasterImage(200, 300, OfficeColor.White);
        var result = OfficeScanProcessor.CorrectPerspective(source, options);
        OfficePoint[] sourceCorners = { options.TopLeft, options.TopRight, options.BottomRight, options.BottomLeft };
        OfficePoint[] outputCorners = { new(0, 0), new(result.Image.Width, 0), new(result.Image.Width, result.Image.Height), new(0, result.Image.Height) };
        for (int index = 0; index < 4; index++) {
            var mapped = result.Mapping.MapProcessedToSource(outputCorners[index]);
            Assert.Equal(sourceCorners[index].X * 200, mapped.X, 7);
            Assert.Equal(sourceCorners[index].Y * 300, mapped.Y, 7);
        }
        for (int y = 0; y <= 10; y++) for (int x = 0; x <= 10; x++) {
                var point = new OfficePoint(x * result.Image.Width / 10D, y * result.Image.Height / 10D);
                var mapped = result.Mapping.MapSourceToProcessed(result.Mapping.MapProcessedToSource(point));
                Assert.Equal(point.X, mapped.X, 7); Assert.Equal(point.Y, mapped.Y, 7);
            }
    }

    [Fact]
    public void CropRetainsSelectedPixelsAndReportsOffset() {
        var source = new OfficeRasterImage(100, 100, OfficeColor.White);
        for (int y = 25; y < 75; y++) for (int x = 25; x < 75; x++) source.SetPixel(x, y, OfficeColor.Black);
        var result = OfficeScanProcessor.CorrectPerspective(source, new OfficeScanPerspectiveOptions {
            TopLeft = new(.25, .25),
            TopRight = new(.75, .25),
            BottomRight = new(.75, .75),
            BottomLeft = new(.25, .75)
        });
        Assert.Equal(50, result.Image.Width); Assert.Equal(50, result.Image.Height);
        Assert.Equal(OfficeColor.Black, result.Image.GetPixel(0, 0)); Assert.Equal(OfficeColor.Black, result.Image.GetPixel(49, 49));
        var corner = result.Mapping.MapProcessedToSource(new(0, 0)); Assert.Equal(25, corner.X); Assert.Equal(25, corner.Y);
    }

    [Fact]
    public void RejectsCrossedCornersBudgetsAndCancellation() {
        var source = new OfficeRasterImage(40, 30, OfficeColor.White);
        Assert.Throws<ArgumentException>(() => OfficeScanProcessor.CorrectPerspective(source, new() { TopRight = new(0, 1), BottomLeft = new(1, 0) }));
        Assert.Throws<ArgumentException>(() => OfficeScanProcessor.CorrectPerspective(source, new() { TopRight = new(0, 0) }));
        Assert.Throws<OfficeScanProcessingLimitException>(() => OfficeScanProcessor.CorrectPerspective(source, new() { MaximumPixels = 100 }));
        Assert.Throws<OfficeScanProcessingLimitException>(() => OfficeScanProcessor.CorrectPerspective(source, new() { MaximumWorkingBytes = 100 }));
        Assert.ThrowsAny<OperationCanceledException>(() => OfficeScanProcessor.CorrectPerspective(source, new(), new CancellationToken(true)));
    }

    [Fact]
    public void ManualStraighteningReportsReversibleGeometryAndLevels() {
        var source = new OfficeRasterImage(100, 80, OfficeColor.FromRgb(128, 128, 128));
        var options = new OfficeScanProcessingOptions {
            Deskew = false,
            NormalizeBackground = false,
            StraightenDegrees = 5,
            BlackPoint = 64,
            WhitePoint = 192,
            Gamma = 2
        };
        var result = OfficeScanProcessor.Process(source, options);
        var point = new OfficePoint(20, 30);
        var roundTrip = result.Report.ProcessedToSource.TransformPoint(result.Report.SourceToProcessed.TransformPoint(point));
        Assert.Equal(point.X, roundTrip.X, 8); Assert.Equal(point.Y, roundTrip.Y, 8);
        Assert.InRange(result.Image.GetPixel(result.Image.Width / 2, result.Image.Height / 2).R, (byte)179, (byte)181);
        Assert.Equal((byte)128, source.GetPixel(20, 30).R);
        Assert.Contains(result.Report.Steps, step => step.Operation == "straighten" && step.Applied);
        Assert.Contains(result.Report.Steps, step => step.Operation == "levels" && step.Applied);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeScanProcessor.Process(source, new() { BlackPoint = 200, WhitePoint = 100 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeScanProcessor.Process(source, new() { Gamma = double.NaN }));
    }
}