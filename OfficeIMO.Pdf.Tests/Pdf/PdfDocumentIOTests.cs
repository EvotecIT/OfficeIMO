using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentIOTests {
    [Fact]
    public async Task LoadAsync_ReadsCompleteCallerStreamAndRestoresPosition() {
        byte[] bytes = BuildDocument("Async load").ToBytes();
        using var stream = new MemoryStream(bytes);
        stream.Position = Math.Min(10, stream.Length);

        PdfDocument document = await PdfDocument.LoadAsync(stream);

        Assert.Equal(Math.Min(10, stream.Length), stream.Position);
        Assert.Equal(1, PdfInspector.Inspect(document.ToBytes()).PageCount);
        stream.ReadByte();
    }

    [Fact]
    public async Task LoadAsync_HonorsPreCanceledTokenAndRestoresPosition() {
        using var stream = new MemoryStream(BuildDocument("Canceled load").ToBytes());
        stream.Position = 7;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PdfDocument.LoadAsync(stream, cancellationToken: cancellation.Token));

        Assert.Equal(7, stream.Position);
    }

    [Fact]
    public void Save_WritesPdfToWritableStream() {
        using var stream = new MemoryStream();

        BuildDocument("Stream save").Save(stream);
        Assert.True(stream.Length > 0);

        stream.Position = 0;
        PdfDocumentInfo info = PdfInspector.Inspect(stream);

        Assert.Equal(1, info.PageCount);
        Assert.Equal("Stream save", info.Metadata.Title);
    }

    [Fact]
    public async Task SaveAsync_WritesPdfToWritableStream() {
        using var stream = new MemoryStream();

        await BuildDocument("Async stream save")
            .SaveAsync(stream);

        Assert.True(stream.Length > 0);

        stream.Position = 0;
        PdfDocumentInfo info = PdfInspector.Inspect(stream);

        Assert.Equal(1, info.PageCount);
        Assert.Equal("Async stream save", info.Metadata.Title);
    }

    [Fact]
    public async Task SaveAsync_WithCanceledToken_DoesNotWriteToStream() {
        using var stream = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            BuildDocument("Canceled stream save").SaveAsync(stream, cancellation.Token));

        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public async Task TrySaveAsync_WithCanceledToken_PropagatesCancellation() {
        using var stream = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            BuildDocument("Canceled try-save").TrySaveAsync(stream, cancellation.Token));

        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public async Task Save_WritesPdfToPathOutputs() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-save-path-" + Guid.NewGuid().ToString("N"));
        string syncPath = Path.Combine(directory, "sync", "document.pdf");
        string asyncPath = Path.Combine(directory, "async", "document.pdf");

        try {
            BuildDocument("Path save").Save(syncPath);
            await BuildDocument("Async path save").SaveAsync(asyncPath);

            Assert.True(File.Exists(syncPath));
            Assert.True(File.Exists(asyncPath));

            Assert.Equal("Path save", PdfInspector.Inspect(syncPath).Metadata.Title);
            Assert.Equal("Async path save", PdfInspector.Inspect(asyncPath).Metadata.Title);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public async Task SaveAsync_WithCanceledToken_DoesNotCreatePathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-save-canceled-" + Guid.NewGuid().ToString("N"));
        string outputPath = Path.Combine(directory, "out", "document.pdf");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            BuildDocument("Canceled path save").SaveAsync(outputPath, cancellation.Token));

        Assert.False(Directory.Exists(directory));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public async Task TrySaveAsync_WithCanceledToken_DoesNotCreatePathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-try-save-canceled-" + Guid.NewGuid().ToString("N"));
        string outputPath = Path.Combine(directory, "out", "document.pdf");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            BuildDocument("Canceled path try-save").TrySaveAsync(outputPath, cancellation.Token));

        Assert.False(Directory.Exists(directory));
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public async Task Save_RejectsNullAndReadOnlyStreams() {
        Assert.Throws<ArgumentNullException>(() => BuildDocument("Null").Save((Stream)null!));
        await Assert.ThrowsAsync<ArgumentNullException>(() => BuildDocument("Null").SaveAsync((Stream)null!));

        using var readOnly = new MemoryStream(new byte[8], writable: false);

        Assert.Throws<ArgumentException>(() => BuildDocument("Read only").Save(readOnly));
        await Assert.ThrowsAsync<ArgumentException>(() => BuildDocument("Read only").SaveAsync(readOnly));
    }

    [Fact]
    public async Task Save_RejectsInvalidPathOutputs() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-save-invalid-path-" + Guid.NewGuid().ToString("N"));

        try {
            Directory.CreateDirectory(directory);

            Assert.Throws<ArgumentNullException>(() => BuildDocument("Null").Save((string)null!));
            Assert.Throws<ArgumentException>(() => BuildDocument("Blank").Save(" "));
            var directoryException = Assert.Throws<ArgumentException>(() => BuildDocument("Directory").Save(directory));
            Assert.Equal("path", directoryException.ParamName);

            await Assert.ThrowsAsync<ArgumentNullException>(() => BuildDocument("Null").SaveAsync((string)null!));
            await Assert.ThrowsAsync<ArgumentException>(() => BuildDocument("Blank").SaveAsync(" "));
            var asyncDirectoryException = await Assert.ThrowsAsync<ArgumentException>(() => BuildDocument("Directory").SaveAsync(directory));
            Assert.Equal("path", asyncDirectoryException.ParamName);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static PdfDocument BuildDocument(string title) {
        return PdfDocument.Create()
            .Meta(title: title, author: "OfficeIMO")
            .Paragraph(p => p.Text("PDF stream output."));
    }
}
