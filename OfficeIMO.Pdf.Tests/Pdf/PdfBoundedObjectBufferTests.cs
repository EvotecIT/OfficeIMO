using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBoundedObjectBufferTests {
    [Fact]
    public void ObjectStore_SpillsCompletedObjectsAndDeletesTemporaryStorage() {
        string spillPath;
        using (var store = new PdfObjectStore(memoryLimitBytes: 4)) {
            store.Add(new byte[] { 1, 2, 3 });
            store.Add(new byte[] { 4, 5, 6 });
            spillPath = Assert.IsType<string>(store.SpillPath);

            Assert.True(store.IsSpilled);
            Assert.True(File.Exists(spillPath));
            Assert.Equal(new byte[] { 1, 2, 3 }, store[0]);
            Assert.Equal(new byte[] { 4, 5, 6 }, store[1]);

            store[0] = new byte[] { 7, 8 };
            Assert.Equal(new byte[] { 7, 8 }, store[0]);
        }

        Assert.False(File.Exists(spillPath));
    }

    [Fact]
    public void Save_WithForcedSpill_WritesReadablePdfToDestinationStream() {
        var options = new PdfOptions { ObjectBufferMemoryLimitBytes = 0 };
        using var output = new MemoryStream();

        PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Bounded object output"))
            .Save(output);

        byte[] bytes = output.ToArray();
        Assert.StartsWith("%PDF-", PdfEncoding.Latin1GetString(bytes), StringComparison.Ordinal);
        Assert.Contains("Bounded object output", PdfTextExtractor.ExtractAllText(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Save_WithForcedSpillAndEncryption_WritesReadableEncryptedPdf() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        };
        var options = new PdfOptions { ObjectBufferMemoryLimitBytes = 0 }.SetEncryption(encryption);
        using var output = new MemoryStream();

        PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Bounded encrypted output"))
            .Save(output);

        byte[] bytes = output.ToArray();
        Assert.Contains(
            "Bounded encrypted output",
            PdfTextExtractor.ExtractAllText(bytes, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "open" }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void ObjectBufferLimit_ValidatesAndClones() {
        var options = new PdfOptions { ObjectBufferMemoryLimitBytes = 1234 };

        Assert.Equal(1234, options.Clone().ObjectBufferMemoryLimitBytes);
        Assert.Throws<ArgumentOutOfRangeException>(() => options.ObjectBufferMemoryLimitBytes = -1);
    }
}
