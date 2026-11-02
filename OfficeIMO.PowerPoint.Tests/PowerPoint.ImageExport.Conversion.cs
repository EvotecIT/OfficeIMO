using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointImageConversionTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg, OfficeImageFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
    public async Task ImageOptions_SelectEncodingAndReturnTheSameEvidenceSynchronouslyAndAsynchronously(
        OfficeImageExportFormat encoding, OfficeImageFormat expected) {
        using var deck = PowerPointDocument.Create();
        var document = deck.AddSlide();
        document.AddTextBox("Image conversion");
        var options = new PowerPointImageExportOptions { Format = encoding, Scale = .15D };
        byte[] bytes = document.ToImage(options, out var report);
        var converted = await document.ToImageAsync(options);
        Assert.Equal(expected, OfficeImageReader.Identify(bytes).Format);
        Assert.Equal(bytes, converted.Value);
        Assert.Equal("PowerPoint", report.Source);
        Assert.Equal(encoding.ToString(), report.Target);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Null(report.OutputPath);
        Assert.Equal(report.Diagnostics.Select(d => (d.Code, d.Severity, d.LossKind, d.Location, d.Format)),
            converted.Report.Diagnostics.Select(d => (d.Code, d.Severity, d.LossKind, d.Location, d.Format)));
        Assert.Equal(.15D, options.Scale);
    }

    [Fact]
    public async Task ImageSaving_PreservesCallerStreamAndHonorsFileConflictPolicy() {
        using var deck = PowerPointDocument.Create();
        var document = deck.AddSlide();
        document.AddTextBox("Saved page");
        var options = new PowerPointImageExportOptions { Format = OfficeImageExportFormat.Png, Scale = .1D };
        using var stream = new MemoryStream();
        var report = await document.SaveAsImageAsync(stream, options);
        Assert.True(stream.CanWrite);
        Assert.Equal(OfficeImageFormat.Png, OfficeImageReader.Identify(stream.ToArray()).Format);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Null(report.OutputPath);
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-PowerPointImages-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "page.png");
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            Assert.Throws<IOException>(() => document.SaveAsImage(path, options));
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(path));
            options.FileConflictPolicy = OfficeImageExportFileConflictPolicy.Replace;
            var saved = document.SaveAsImage(path, options);
            Assert.Equal(Path.GetFullPath(path), saved.OutputPath);
            Assert.Equal(OfficeImageFormat.Png, OfficeImageReader.Identify(File.ReadAllBytes(path)).Format);
        } finally {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public async Task ImageBatchSaving_ReportsSelectedSlidesAndHonorsConflicts() {
        using var document = PowerPointDocument.Create();
        document.SlideSize.SetSize(OfficeLength.Points(100), OfficeLength.Points(60));
        document.AddSlide().BackgroundColor = "112233";
        document.AddSlide().Hidden = true;
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-SlideBatch-" + Guid.NewGuid().ToString("N"));
        var options = new PowerPointImageExportOptions {
            SlideNumbers = new[] { 2 }, IncludeHiddenSlides = true, Format = OfficeImageExportFormat.Webp
        };
        try {
            var report = await document.SaveAsImagesAsync(folder, options);
            Assert.Equal(1, report.ItemsWritten);
            Assert.Equal(Path.GetFullPath(folder), report.OutputPath);
            string path = Assert.Single(Directory.GetFiles(folder));
            Assert.EndsWith("Slide 2.webp", path, StringComparison.Ordinal);
            byte[] first = File.ReadAllBytes(path);
            Assert.Equal(OfficeImageFormat.Webp, OfficeImageReader.Identify(first).Format);
            Assert.Throws<IOException>(() => document.SaveAsImages(folder, options));
            Assert.Equal(first, File.ReadAllBytes(path));
            options.FileConflictPolicy = OfficeImageExportFileConflictPolicy.Replace;
            Assert.Equal(1, document.SaveAsImages(folder, options).ItemsWritten);
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public async Task ImageSaving_PreCanceledOperationsLeaveStreamsUntouched() {
        using var deck = PowerPointDocument.Create();
        var document = deck.AddSlide();
        document.AddTextBox("Canceled");
        using var stream = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToImage(cancellationToken: cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() => document.SaveAsImage(stream, cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.SaveAsImageAsync(stream,
            cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.ToImageAsync(cancellationToken: cancellation.Token));
        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
    [Fact]
    public void FailOnLossRejectsAnOmittedGroupBeforeWriting() {
        using var deck = PowerPointDocument.Create();
        var slide = deck.AddSlide();
        var first = slide.AddRectangle(OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20));
        var second = slide.AddRectangle(OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20));
        slide.GroupShapes(new PowerPointShape[] { first, second }, "Omitted group");
        using var output = new MemoryStream();
        var options = new PowerPointImageExportOptions {
            Format = OfficeImageExportFormat.Svg, MaxGroupShapeDepth = 0, FailOnLoss = true
        };
        var exception = Assert.Throws<OfficeConversionException>(() => slide.SaveAsImage(output, options));
        var report = Assert.IsType<OfficeConversionReport>(exception.Report);
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "group-depth-limit");
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
    }
}
