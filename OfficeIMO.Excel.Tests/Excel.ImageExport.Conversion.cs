using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelImageConversionTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png, OfficeImageFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg, OfficeImageFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
    public async Task ImageOptions_SelectEncodingAndReturnTheSameEvidenceSynchronouslyAndAsynchronously(
        OfficeImageExportFormat encoding, OfficeImageFormat expected) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Data");
        var document = sheet.Range("A1:B2");
        sheet.CellValue(1, 1, "Image conversion");
        var options = new ExcelImageExportOptions { Format = encoding, Scale = .15D };
        byte[] bytes = document.ToImage(options, out var report);
        var converted = await document.ToImageAsync(options);
        Assert.Equal(expected, OfficeImageReader.Identify(bytes).Format);
        Assert.Equal(bytes, converted.Value);
        Assert.Equal("Excel", report.Source);
        Assert.Equal(encoding.ToString(), report.Target);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Null(report.OutputPath);
        Assert.Equal(report.Diagnostics.Select(d => (d.Code, d.Severity, d.LossKind, d.Location, d.Format)),
            converted.Report.Diagnostics.Select(d => (d.Code, d.Severity, d.LossKind, d.Location, d.Format)));
        Assert.Equal(.15D, options.Scale);
    }

    [Fact]
    public async Task ImageSaving_PreservesCallerStreamAndHonorsFileConflictPolicy() {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Data");
        var document = sheet.Range("A1:B2");
        sheet.CellValue(1, 1, "Saved page");
        var options = new ExcelImageExportOptions { Format = OfficeImageExportFormat.Png, Scale = .1D };
        using var stream = new MemoryStream();
        var report = await document.SaveAsImageAsync(stream, options);
        Assert.True(stream.CanWrite);
        Assert.Equal(OfficeImageFormat.Png, OfficeImageReader.Identify(stream.ToArray()).Format);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Null(report.OutputPath);
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelImages-" + Guid.NewGuid().ToString("N"));
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
    public async Task WorkbookSaving_SelectsSheetsAndHonorsConflicts() {
        using var document = ExcelDocument.Create();
        document.AddWorksheet("First").CellValue(1, 1, "First");
        document.AddWorksheet("Hidden").CellValue(1, 1, "Hidden");
        document.Sheets[1].SetHidden(true);
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-SheetBatch-" + Guid.NewGuid().ToString("N"));
        var options = new ExcelImageExportOptions {
            SheetNames = new[] { "Hidden" }, Format = OfficeImageExportFormat.Webp
        };
        try {
            var report = await document.SaveAsImagesAsync(folder, options);
            Assert.Equal(1, report.ItemsWritten);
            Assert.Equal(Path.GetFullPath(folder), report.OutputPath);
            string path = Assert.Single(Directory.GetFiles(folder));
            Assert.EndsWith("Hidden.webp", path, StringComparison.Ordinal);
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
    public async Task WorksheetSaving_RetainsManualPageBreakSplits() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Data");
        for (int row = 1; row <= 4; row++) sheet.CellValue(row, 1, "Row " + row);
        sheet.AddManualRowPageBreak(2, save: false);
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-PrintImages-" + Guid.NewGuid().ToString("N"));
        var options = new ExcelImageExportOptions {
            Format = OfficeImageExportFormat.Svg, Range = "A1:A4", SplitByManualPageBreaks = true
        };
        try {
            var report = await sheet.SaveAsImagesAsync(folder, options);
            Assert.Equal(2, report.ItemsWritten);
            Assert.Equal(2, Directory.GetFiles(folder, "*.svg").Length);
            Assert.All(Directory.GetFiles(folder), path => Assert.Contains("<svg", File.ReadAllText(path), StringComparison.Ordinal));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public async Task ImageSaving_PreCanceledOperationsLeaveStreamsUntouched() {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Data");
        var document = sheet.Range("A1:B2");
        sheet.CellValue(1, 1, "Canceled");
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
