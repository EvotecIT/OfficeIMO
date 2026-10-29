using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterStreamFileTests {
    [Fact]
    public async Task StreamFileConversionPreservesSourcePositionAndPublishesTheCompletedPdf() {
        using var scope = new ConversionDirectory();
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("# Captured text\n\nStream publication evidence."));
        source.Position = 3;
        string destination = Path.Combine(scope.Root, "converted.pdf");
        OfficeConversionReport report = await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, destination);
        Assert.Equal(3, source.Position);
        Assert.True(source.CanRead);
        Assert.Equal(destination, report.OutputPath);
        using var pdf = PdfDocument.Load(destination);
        Assert.Contains("Stream publication evidence.", pdf.Read().Text);
        Assert.Empty(Directory.GetFiles(scope.Root, ".*.tmp"));
    }

    [Fact]
    public async Task StreamFileConversionRejectsLossAndCancellationBeforeCreatingTheDestination() {
        using var scope = new ConversionDirectory();
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("# Capture\n\n![Unavailable](https://example.invalid/image.png)"));
        string destination = Path.Combine(scope.Root, "new", "converted.pdf");
        await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown,
            destination, new OfficeConversionOptions { FailOnLoss = true }));
        Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown,
            destination, cancellationToken: cancellation.Token));
        Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
    }

    [Fact]
    public async Task StreamFilePublicationRetainsOccupiedOutputAndAppliesTheSharedRenamePolicy() {
        using var scope = new ConversionDirectory();
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("A verified replacement candidate."));
        string destination = Path.Combine(scope.Root, "converted.pdf");
        File.WriteAllText(destination, "Retain this output");
        await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Text, destination));
        Assert.Equal("Retain this output", File.ReadAllText(destination));
        OfficeConversionReport report = await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Text, destination,
            new OfficeConversionOptions { FileConflictPolicy = OfficePublicationPolicy.Rename });
        Assert.NotEqual(destination, report.OutputPath);
        Assert.Equal("Retain this output", File.ReadAllText(destination));
        using var pdf = PdfDocument.Load(report.OutputPath!);
        Assert.Contains("A verified replacement candidate.", pdf.Read().Text);
        Assert.Empty(Directory.GetFiles(scope.Root, ".*.tmp"));
    }

    [Fact]
    public async Task FileStreamDestinationDoesNotReopenOrReplaceItsSourceAndFileTextExportKeepsTheOutputStreamOpen() {
        using var scope = new ConversionDirectory();
        string sourcePath = Path.Combine(scope.Root, "source.docx");
        using (var document = WordDocument.Create()) { document.AddParagraph("Streamed Word export"); document.Save(sourcePath); }
        byte[] before = File.ReadAllBytes(sourcePath);
        using (var source = File.OpenRead(sourcePath))
            await Assert.ThrowsAsync<ArgumentException>(() => OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Word, sourcePath));
        Assert.Equal(before, File.ReadAllBytes(sourcePath));
        using var target = new MemoryStream();
        OfficeConversionReport report = await OfficeConverter.ConvertAsync(sourcePath, target, OfficeDocumentFormat.Markdown);
        Assert.True(target.CanWrite);
        Assert.Null(report.OutputPath);
        Assert.Contains("Streamed Word export", Encoding.UTF8.GetString(target.ToArray()));
        Assert.Equal(before, File.ReadAllBytes(sourcePath));
    }

    private sealed class ConversionDirectory : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-stream-file-" + Guid.NewGuid().ToString("N"));
        internal ConversionDirectory() => Directory.CreateDirectory(Root);
        public void Dispose() => Directory.Delete(Root, true);
    }
}
