using OfficeIMO.Pdf;
using System.Collections;
using System.Text;
using System.Threading;
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
            Assert.ThrowsAny<IOException>(() => File.OpenRead(spillPath));
#if NET6_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(spillPath));
            }
#endif
            Assert.Equal("abc", store.Read(first));
            Assert.Equal("def", store.Read(second));
        }

        Assert.False(File.Exists(spillPath));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1024)]
    public void PageContentStore_StoresStringBuilderAsEquivalentLatin1Bytes(long memoryLimitBytes) {
        const string content = "BT\n/F1 12 Tf\n(caf\u00e9) Tj\nET\n";
        var builder = new StringBuilder(content);
        using var store = new PdfPageContentStore(memoryLimitBytes);

        PdfPageContentHandle handle = store.Store(builder);
        builder.Clear();

        Assert.Equal(PdfEncoding.Latin1GetBytes(content), store.ReadBytes(handle));
        Assert.Equal(content, store.Read(handle));
        Assert.Equal(memoryLimitBytes == 0, store.IsSpilled);
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
            Assert.ThrowsAny<IOException>(() => File.OpenRead(spillPath));
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
    public void ObjectStore_CopyStopsBetweenInMemoryChunksWhenCancelled() {
        using var cancellation = new CancellationTokenSource();
        using var store = new PdfObjectStore(memoryLimitBytes: 256 * 1024);
        store.Add(new byte[160 * 1024]);
        using var destination = new CancellingWriteStream(cancellation);

        Assert.ThrowsAny<OperationCanceledException>(() =>
            store.CopyTo(0, destination, cancellationToken: cancellation.Token));
        Assert.True(cancellation.IsCancellationRequested);
        Assert.True(destination.Length < store.GetLength(0));
    }

    [Fact]
    public void BufferedAssembly_ForwardsCancellationIntoEncryption() {
        using var cancellation = new CancellationTokenSource();
        var objects = new CancelOnReadObjectList(cancellation, cancelOnRead: 5,
            Encoding.ASCII.GetBytes("1 0 obj\n<< /Type /Catalog >>\nendobj\n"),
            Encoding.ASCII.GetBytes("2 0 obj\n<< >>\nendobj\n"));
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        };

        Assert.ThrowsAny<OperationCanceledException>(() => PdfFileAssembler.AssembleWithEvidence(
            objects,
            catalogId: 1,
            infoId: 0,
            PdfFileVersion.Pdf14,
            encryption,
            PdfObjectStore.DefaultMemoryLimitBytes,
            cancellation.Token,
            out _));
        Assert.True(cancellation.IsCancellationRequested);
    }

    [Fact]
    public void BufferedAssembly_StopsBetweenLargeTrailerWritesWhenCancelled() {
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelOnLargeWriteStream(cancellation);
        byte[][] objects = { Encoding.ASCII.GetBytes("1 0 obj\n<< /Type /Catalog >>\nendobj\n") };
        string trailerId = " /ID [<" + new string('A', 160_000) + "> <00>]";

        Assert.ThrowsAny<OperationCanceledException>(() => PdfFileAssembler.Assemble(
            destination, objects, 1, 0, trailerIdEntry: trailerId, cancellationToken: cancellation.Token));
        Assert.True(cancellation.IsCancellationRequested);
        Assert.InRange(destination.Length, 1, 100_000);
    }

    [Fact]
    public void ForwardOnlyAssembly_MatchesBufferedTrailerOutput() {
        byte[] catalog = Encoding.ASCII.GetBytes("1 0 obj\n<< /Type /Catalog >>\nendobj\n");
        using var destination = new MemoryStream();
        using (var store = new PdfForwardOnlyObjectStore(destination, PdfFileVersion.Pdf14)) {
            store.Add(catalog);
            store.Complete(1, 0);
        }

        Assert.Equal(PdfFileAssembler.Assemble(new[] { catalog }, 1, 0), destination.ToArray());
    }

    [Fact]
    public void BufferedAssembly_MatchesStreamOutputForTrailerVariants() {
        byte[][] objects = {
            Encoding.ASCII.GetBytes("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"),
            Encoding.ASCII.GetBytes("2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n")
        };
        const string trailerId = " /ID [<01020304> <05060708>]";
        byte[] permanentId = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };

        using var defaultOutput = new MemoryStream();
        PdfFileAssembler.Assemble(defaultOutput, objects, 1, 0, PdfFileVersion.Pdf14);
        Assert.Equal(defaultOutput.ToArray(), PdfFileAssembler.Assemble(objects, 1, 0, PdfFileVersion.Pdf14));

        using var explicitIdOutput = new MemoryStream();
        PdfFileAssembler.Assemble(explicitIdOutput, objects, 1, 0, PdfFileVersion.Pdf17, trailerIdEntry: trailerId);
        Assert.Equal(explicitIdOutput.ToArray(), PdfFileAssembler.Assemble(objects, 1, 0, PdfFileVersion.Pdf17, trailerIdEntry: trailerId));

        using var permanentIdOutput = new MemoryStream();
        PdfFileAssembler.AssemblePreservingPermanentId(permanentIdOutput, objects, 1, 0, PdfFileVersion.Pdf20, null, permanentId);
        byte[] permanentIdBytes = PdfFileAssembler.AssemblePreservingPermanentId(objects, 1, 0, PdfFileVersion.Pdf20, null, permanentId);
        Assert.Equal(permanentIdOutput.ToArray(), permanentIdBytes);
        Assert.Equal(0, PdfInspector.Inspect(permanentIdBytes).PageCount);
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
            PdfTextExtractor.ExtractAllText(bytes, (PdfTextLayoutOptions?)null, new PdfLoadOptions { Password = "open" }),
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

    private sealed class CancellingWriteStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancellingWriteStream(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            _cancellation.Cancel();
        }
    }

    private sealed class CancelOnLargeWriteStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnLargeWriteStream(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            if (count >= 64 * 1024) _cancellation.Cancel();
        }
    }

    private sealed class CancelOnReadObjectList : IReadOnlyList<byte[]> {
        private readonly CancellationTokenSource _cancellation;
        private readonly int _cancelOnRead;
        private readonly byte[][] _objects;
        private int _readCount;

        internal CancelOnReadObjectList(
            CancellationTokenSource cancellation,
            int cancelOnRead,
            params byte[][] objects) {
            _cancellation = cancellation;
            _cancelOnRead = cancelOnRead;
            _objects = objects;
        }

        public int Count => _objects.Length;

        public byte[] this[int index] {
            get {
                if (++_readCount == _cancelOnRead) _cancellation.Cancel();
                return _objects[index];
            }
        }

        public IEnumerator<byte[]> GetEnumerator() => ((IEnumerable<byte[]>)_objects).GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
