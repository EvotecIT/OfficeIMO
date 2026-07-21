using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBoundedObjectBufferTests {
    [Fact]
    public void PageContentStore_SpillsCompletedPagesAndDeletesTemporaryStorage() {
        string spillPath;
        using (var store = new PdfPageContentStore(memoryLimitBytes: 4)) {
            PdfPageContentHandle first = store.Store("abc");
            PdfPageContentHandle second = store.Store("def");
            spillPath = Assert.IsType<string>(store.SpillPath);

            Assert.True(store.IsSpilled);
            Assert.Equal(0, store.RetainedMemoryBytes);
            Assert.True(File.Exists(spillPath));
            Assert.Equal("abc", store.Read(first));
            Assert.Equal("def", store.Read(second));
        }

        Assert.False(File.Exists(spillPath));
    }

    [Fact]
    public void ObjectStore_SpillsCompletedObjectsAndDeletesTemporaryStorage() {
        string spillPath;
        using (var store = new PdfObjectStore(memoryLimitBytes: 4)) {
            store.Add(new byte[] { 1, 2, 3 });
            store.Add(new byte[] { 4, 5, 6 });
            spillPath = Assert.IsType<string>(store.SpillPath);

            Assert.True(store.IsSpilled);
            Assert.Equal(0, store.RetainedMemoryBytes);
            Assert.True(File.Exists(spillPath));
            Assert.Equal(new byte[] { 1, 2, 3 }, store[0]);
            Assert.Equal(new byte[] { 4, 5, 6 }, store[1]);

            store[0] = new byte[] { 7, 8 };
            Assert.Equal(new byte[] { 7, 8 }, store[0]);
        }

        Assert.False(File.Exists(spillPath));
    }

    [Fact]
    public void ObjectStore_CopiesSpilledSegmentsDirectlyToDestination() {
        using var store = new PdfObjectStore(memoryLimitBytes: 0);
        store.AddSegments(new byte[] { 1, 2 }, new byte[] { 3, 4, 5 });
        using var destination = new MemoryStream();

        store.CopyTo(0, destination);

        Assert.True(store.IsSpilled);
        Assert.Equal(0, store.RetainedMemoryBytes);
        Assert.Equal(5, store.GetLength(0));
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, destination.ToArray());
    }

    [Fact]
    public void Save_WithForcedSpill_WritesReadablePdfToDestinationStream() {
        var options = new PdfOptions {
            ObjectBufferMemoryLimitBytes = 0,
            PageContentMemoryLimitBytes = 0
        };
        using var output = new MemoryStream();

        PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Bounded object output"))
            .Save(output);

        byte[] bytes = output.ToArray();
        Assert.StartsWith("%PDF-", PdfEncoding.Latin1GetString(bytes), StringComparison.Ordinal);
        Assert.Contains("Bounded object output", PdfTextExtractor.ExtractAllText(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Save_WithForcedPageAndObjectSpill_WritesEveryPage() {
        const int pageCount = 24;
        var options = new PdfOptions {
            ObjectBufferMemoryLimitBytes = 0,
            PageContentMemoryLimitBytes = 0
        };
        PdfDocument document = PdfDocument.Create(options);
        for (int page = 1; page <= pageCount; page++) {
            document.Paragraph(paragraph => paragraph.Text("Bounded page " + page));
            if (page < pageCount) document.PageBreak();
        }
        using var output = new MemoryStream();

        document.Save(output);

        byte[] bytes = output.ToArray();
        Assert.Equal(pageCount, PdfInspector.Inspect(bytes).PageCount);
        string text = PdfTextExtractor.ExtractAllText(bytes);
        Assert.Contains("Bounded page 1", text, StringComparison.Ordinal);
        Assert.Contains("Bounded page 24", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Save_WithForcedSpillAndEncryption_WritesReadableEncryptedPdf() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        };
        var options = new PdfOptions { ObjectBufferMemoryLimitBytes = 0 }.SetEncryption(encryption);
        using var output = new MemoryStream();

        PdfSaveResult result = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Bounded encrypted output"))
            .Save(output);

        PdfSerializationReport serialization = Assert.IsType<PdfSerializationReport>(result.Serialization);
        Assert.True(serialization.ObjectBufferSpilled);
        Assert.Equal(0, serialization.PeakRetainedObjectBytes);
        byte[] bytes = output.ToArray();
        Assert.Contains(
            "Bounded encrypted output",
            PdfTextExtractor.ExtractAllText(bytes, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "open" }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void BufferLimits_ValidateAndClone() {
        var options = new PdfOptions {
            ObjectBufferMemoryLimitBytes = 1234,
            PageContentMemoryLimitBytes = 5678
        };

        Assert.Equal(1234, options.Clone().ObjectBufferMemoryLimitBytes);
        Assert.Equal(5678, options.Clone().PageContentMemoryLimitBytes);
        Assert.Throws<ArgumentOutOfRangeException>(() => options.ObjectBufferMemoryLimitBytes = -1);
        Assert.Throws<ArgumentOutOfRangeException>(() => options.PageContentMemoryLimitBytes = -1);
    }
}
