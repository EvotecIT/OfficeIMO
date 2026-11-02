using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordImageConversionTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg, OfficeImageFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
    public async Task ImageOptions_SelectEncodingAndReturnTheSameEvidenceSynchronouslyAndAsynchronously(
        OfficeImageExportFormat encoding, OfficeImageFormat expected) {
        using var document = WordDocument.Create();
        document.AddParagraph("Image conversion");
        var options = new WordImageExportOptions { Format = encoding, Scale = .15D };
        byte[] bytes = document.ToImage(options, out var report);
        var converted = await document.ToImageAsync(options);
        Assert.Equal(expected, OfficeImageReader.Identify(bytes).Format);
        Assert.Equal(bytes, converted.Value);
        Assert.Equal("Word", report.Source);
        Assert.Equal(encoding.ToString(), report.Target);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Null(report.OutputPath);
        Assert.Equal(report.Diagnostics.Select(d => (d.Code, d.Severity, d.LossKind, d.Location, d.Format)),
            converted.Report.Diagnostics.Select(d => (d.Code, d.Severity, d.LossKind, d.Location, d.Format)));
        Assert.Equal(.15D, options.Scale);
    }

    [Fact]
    public async Task ImageSaving_PreservesCallerStreamAndHonorsFileConflictPolicy() {
        using var document = WordDocument.Create();
        document.AddParagraph("Saved page");
        var options = new WordImageExportOptions { Format = OfficeImageExportFormat.Png, Scale = .1D };
        using var stream = new MemoryStream();
        var report = await document.SaveAsImageAsync(stream, options);
        Assert.True(stream.CanWrite);
        Assert.Equal(OfficeImageFormat.Png, OfficeImageReader.Identify(stream.ToArray()).Format);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Null(report.OutputPath);
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-WordImages-" + Guid.NewGuid().ToString("N"));
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
    public async Task ImageBatchSaving_ReportsOnlyTheSelectedPages() {
        using var document = WordDocument.Create();
        document.AddParagraph("First");
        document.AddPageBreak();
        document.AddParagraph("Second");
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-WordPageBatch-" + Guid.NewGuid().ToString("N"));
        try {
            var report = await document.SaveAsImagesAsync(folder,
                new WordImageExportOptions { PageIndex = 1, PageCount = 1, Format = OfficeImageExportFormat.Webp, Scale = .1D });
            Assert.Equal(1, report.ItemsWritten);
            Assert.Equal(Path.GetFullPath(folder), report.OutputPath);
            string path = Assert.Single(Directory.GetFiles(folder));
            Assert.EndsWith(".webp", path, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(OfficeImageFormat.Webp, OfficeImageReader.Identify(File.ReadAllBytes(path)).Format);
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public async Task ImageSaving_PreCanceledOperationsLeaveStreamsUntouched() {
        using var document = WordDocument.Create();
        document.AddParagraph("Canceled");
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
}
