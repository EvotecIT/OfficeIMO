using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingScanProcessingTests {
    [Theory]
    [InlineData("phototest", false)]
    [InlineData("eurotext", false)]
    [InlineData("phototest-shadow", true)]
    [InlineData("eurotext-shadow", true)]
    public void DeskewFindsIndependentlyRotatedScannedTextAndPreservesSource(string name, bool normalizeBackground) {
        byte[] bytes = System.IO.File.ReadAllBytes(System.IO.Path.Combine(AppContext.BaseDirectory, "ScanQuality", name + "-skew.png"));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image));
        byte[] original = image!.GetPixels();
        var result = OfficeScanProcessor.Process(image, new OfficeScanProcessingOptions { NormalizeBackground = normalizeBackground });
        Assert.InRange(result.Report.DetectedSkewDegrees, 2.2D, 3.8D);
        Assert.InRange(result.Report.AppliedDeskewDegrees, -3.8D, -2.2D);
        Assert.Equal(original, image.GetPixels());
        OfficePoint originalPoint = new OfficePoint(image.Width * 0.35D, image.Height * 0.6D);
        OfficePoint roundTrip = result.Report.ProcessedToSource.TransformPoint(result.Report.SourceToProcessed.TransformPoint(originalPoint));
        Assert.Equal(originalPoint.X, roundTrip.X, 8); Assert.Equal(originalPoint.Y, roundTrip.Y, 8);
    }
    [Fact]
    public void QuarterTurnPreservesSourceAndMapsPixelEdgesBackExactly() {
        var source = new OfficeRasterImage(3, 2, OfficeColor.White);
        source.SetPixel(0, 0, OfficeColor.Red);
        source.SetPixel(2, 1, OfficeColor.Blue);
        byte[] original = source.GetPixels();
        var result = OfficeScanProcessor.Process(source, new OfficeScanProcessingOptions {
            ClockwiseQuarterTurns = 1, Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.PreserveColor
        });
        Assert.Equal(original, source.GetPixels());
        Assert.Equal(2, result.Image.Width);
        Assert.Equal(3, result.Image.Height);
        Assert.Equal(OfficeColor.Red, result.Image.GetPixel(1, 0));
        Assert.Equal(OfficeColor.Blue, result.Image.GetPixel(0, 2));
        OfficePoint mapped = result.Report.SourceToProcessed.TransformPoint(new OfficePoint(0, 0));
        Assert.Equal(2D, mapped.X, 8); Assert.Equal(0D, mapped.Y, 8);
        OfficePoint restored = result.Report.ProcessedToSource.TransformPoint(mapped);
        Assert.Equal(0D, restored.X, 8); Assert.Equal(0D, restored.Y, 8);
        result.Image.SetPixel(1, 0, OfficeColor.Black);
        Assert.Equal(original, source.GetPixels());
    }

    [Fact]
    public void BackgroundAndBilevelRetainTextOnShadedPaperAndReportDecisions() {
        var source = new OfficeRasterImage(100, 40, OfficeColor.FromRgb(180, 180, 180));
        for (int y = 15; y < 20; y++) for (int x = 25; x < 75; x++) source.SetPixel(x, y, OfficeColor.FromRgb(40, 40, 40));
        byte[] original = source.GetPixels();
        var result = OfficeScanProcessor.Process(source, new OfficeScanProcessingOptions {
            Deskew = false, ColorMode = OfficeScanColorMode.Bilevel, BackgroundRadius = 8
        });
        Assert.Equal(OfficeColor.White, result.Image.GetPixel(50, 4));
        Assert.Equal(OfficeColor.Black, result.Image.GetPixel(50, 17));
        Assert.Equal(original, source.GetPixels());
        Assert.False(result.Report.IsProbablyBlank);
        Assert.True(result.Report.ForegroundFraction > 0D);
        Assert.Contains(result.Report.Steps, step => step.Operation == "background" && step.Applied);
        Assert.Contains(result.Report.Steps, step => step.Operation == "color" && step.Applied);
    }

    [Fact]
    public void DownsamplingRetainsAspectRatioGeometryAndBilevelContract() {
        var source = new OfficeRasterImage(120, 80, OfficeColor.White);
        for (int y = 0; y < 80; y++) for (int x = 0; x < 60; x++) source.SetPixel(x, y, OfficeColor.Black);
        var result = OfficeScanProcessor.Process(source, new OfficeScanProcessingOptions {
            Deskew = false, NormalizeBackground = false, MaximumDimension = 60, ColorMode = OfficeScanColorMode.Bilevel
        });
        Assert.Equal(60, result.Image.Width); Assert.Equal(40, result.Image.Height);
        Assert.Equal(OfficeColor.Black, result.Image.GetPixel(10, 20));
        Assert.Equal(OfficeColor.White, result.Image.GetPixel(50, 20));
        Assert.All(result.Image.GetPixels(), value => Assert.True(value == 0 || value == 255));
        OfficePoint sourcePoint = result.Report.ProcessedToSource.TransformPoint(new OfficePoint(30, 20));
        Assert.Equal(60D, sourcePoint.X, 8); Assert.Equal(40D, sourcePoint.Y, 8);
    }

    [Fact]
    public void BlankAndLowConfidenceInputsAreRetainedWithoutInventingDeskew() {
        var source = new OfficeRasterImage(100, 80, OfficeColor.White);
        var result = OfficeScanProcessor.Process(source);
        Assert.True(result.Report.IsProbablyBlank);
        Assert.Equal(0D, result.Report.AppliedDeskewDegrees);
        Assert.Equal(source.GetPixels(), result.Image.GetPixels());
        Assert.Contains(result.Report.Steps, step => step.Operation == "deskew" && !step.Applied);
    }

    [Fact]
    public void LimitsAndCancellationLeaveTheOriginalUntouched() {
        var source = new OfficeRasterImage(100, 80, OfficeColor.White);
        byte[] original = source.GetPixels();
        Assert.Throws<OfficeScanProcessingLimitException>(() => OfficeScanProcessor.Process(source,
            new OfficeScanProcessingOptions { MaximumPixels = 100 }));
        Assert.Throws<OfficeScanProcessingLimitException>(() => OfficeScanProcessor.Process(source,
            new OfficeScanProcessingOptions { MaximumWorkingBytes = 100 }));
        Assert.ThrowsAny<OperationCanceledException>(() => OfficeScanProcessor.Process(source, cancellationToken: new CancellationToken(true)));
        Assert.ThrowsAny<OperationCanceledException>(() => OfficeRasterResampler.Resize(source, 50, 40,
            OfficeRasterResamplingMode.Area, OfficeRasterResamplingColorSpace.EncodedSrgb, new CancellationToken(true)));
        Assert.Equal(original, source.GetPixels());
    }

    [Fact]
    public void DeskewAnalysisWorkLimitRejectsBeforeChangingTheSource() {
        byte[] bytes = System.IO.File.ReadAllBytes(System.IO.Path.Combine(AppContext.BaseDirectory, "ScanQuality", "phototest-skew.png"));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? source));
        byte[] original = source!.GetPixels();
        Assert.Throws<OfficeScanProcessingLimitException>(() => OfficeScanProcessor.Process(source,
            new OfficeScanProcessingOptions { MaximumAnalysisOperations = 1 }));
        Assert.Equal(original, source.GetPixels());
    }
}
