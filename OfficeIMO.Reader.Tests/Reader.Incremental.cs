using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Excel;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderIncrementalTests {
    [Fact]
    public void WorkbookIncrementalPathUsesOwnerRowChunksAndReleasesTheFile() {
        string path = Path.Combine(Path.GetTempPath(), "reader-incremental-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (var workbook = ExcelDocument.Create(path)) {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1, "Name");
                for (int row = 2; row <= 20; row++) sheet.Cell(row, 1, "row " + row);
                workbook.Save();
            }
            var reader = new OfficeDocumentReaderBuilder().AddExcelHandler(new ReaderExcelOptions { ChunkRows = 4 }).Build();
            var options = new ReaderOptions { ComputeHashes = false };
            var rich = reader.ReadDocument(path, options);
            var incremental = reader.EnumerateChunks(path, options).ToArray();
            Assert.Equal(rich.Chunks.Select(c => c.Markdown), incremental.Select(c => c.Markdown));
            Assert.True(incremental.Length > 1);
            using (var iterator = reader.EnumerateChunks(path, options).GetEnumerator()) Assert.True(iterator.MoveNext());
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void TextStreamProducesFirstChunkBeforeReadingWholeInputAndPreservesOwnership() {
        using var input = new CountingStream(Encoding.UTF8.GetBytes(new string('a', 100_000)));
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        using (var iterator = reader.EnumerateChunks(input, "large.txt",
                   new ReaderOptions { ComputeHashes = false, MaxChars = 256 }).GetEnumerator()) {
            Assert.True(iterator.MoveNext());
            Assert.Equal(256, iterator.Current.Text.Length);
            Assert.InRange(input.BytesRead, 256, 8192);
        }
        Assert.True(input.CanRead);
    }

    [Fact]
    public void IncrementalSeekableInputRestoresPositionAfterEarlyDisposal() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(new string('a', 10_000)));
        input.Position = 7;
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        using (var iterator = reader.EnumerateChunks(input, "large.txt",
                   new ReaderOptions { ComputeHashes = false, MaxChars = 256 }).GetEnumerator()) Assert.True(iterator.MoveNext());
        Assert.Equal(7, input.Position);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void IncrementalInputEnforcesBoundBeforeReadingBeyondLimitPlusOne() {
        using var input = new CountingStream(new byte[100_000]);
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        Assert.Throws<IOException>(() => reader.EnumerateChunks(input, "large.txt",
            new ReaderOptions { ComputeHashes = false, MaxInputBytes = 128 }).ToArray());
        Assert.InRange(input.BytesRead, 1, 129);
    }

    [Fact]
    public void IncrementalReadsRespectDocumentProcessorsAndCancellation() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("original"));
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddProcessor(new DelegateOfficeDocumentProcessor("replace", (document, _) => {
                document.Chunks[0].Text = "processed"; return document;
            })).Build();
        Assert.Equal("processed", Assert.Single(reader.EnumerateChunks(input, "sample.txt")).Text);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => reader.EnumerateChunks(input, "sample.txt",
            cancellationToken: cancellation.Token).ToArray());
    }

    [Fact]
    public void CsvPathYieldsValidatedRowsAndReportsItsActualStreamingBoundary() {
        string path = Path.Combine(Path.GetTempPath(), "reader-stream-" + Guid.NewGuid().ToString("N") + ".csv");
        File.WriteAllText(path, "Name,Count\nalpha,1\nbeta,2\n");
        try {
            var reader = new OfficeDocumentReaderBuilder().AddCsvHandler(new CsvReadOptions { ChunkRows = 1 }).Build();
            var capability = Assert.Single(reader.GetCapabilities());
            Assert.True(capability.SupportsIncrementalPath);
            Assert.False(capability.SupportsIncrementalStream);
            var chunks = reader.EnumerateChunks(path, new ReaderOptions { ComputeHashes = false }).ToArray();
            Assert.Equal(2, chunks.Length);
            Assert.Equal("alpha", Assert.Single(Assert.Single(chunks[0].Tables!).Rows)[0]);
            Assert.Equal("beta", Assert.Single(Assert.Single(chunks[1].Tables!).Rows)[0]);
        } finally { File.Delete(path); }
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task AsyncIncrementalEarlyDisposalReleasesTheReaderGate() {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().WithMaxConcurrentReads(1).Build();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        for (int attempt = 0; attempt < 2; attempt++) {
            using var input = new CountingStream(Encoding.UTF8.GetBytes(new string('a', 10_000)));
            await foreach (var chunk in reader.EnumerateChunksAsync(input, "sample.txt",
                               new ReaderOptions { ComputeHashes = false, MaxChars = 256 }, cancellation.Token)) {
                Assert.Equal(256, chunk.Text.Length);
                break;
            }
            Assert.True(input.CanRead);
        }
    }
#endif

    [Theory]
    [InlineData("utf-8")]
    [InlineData("utf-16")]
    [InlineData("utf-16BE")]
    [InlineData("utf-32")]
    public void BomOverridesExplicitEncodingWithoutBreakingUnicode(string encodingName) {
        string text = new string('a', 255) + "\U0001F600\r\nsecond\rthird";
        Encoding encoding = Encoding.GetEncoding(encodingName);
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(text)).ToArray();
        var result = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build().Read(bytes, "text.txt",
            new ReaderOptions { MaxChars = 256, TextEncoding = Encoding.ASCII, ThrowOnInvalidTextBytes = true });
        Assert.Equal(text.Replace("\r\n", "\n").Replace('\r', '\n'), string.Concat(result.Select(chunk => chunk.Text)));
    }

    [Fact]
    public void ExplicitEncodingAndInvalidBytePoliciesHaveObservableResults() {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        Assert.Equal("café", Assert.Single(reader.Read(new byte[] { 99, 97, 102, 233 }, "latin.txt",
            new ReaderOptions { TextEncoding = Encoding.GetEncoding("iso-8859-1") })).Text);
        byte[] malformed = { 0xef, 0xbb, 0xbf, 0xff };
        var chunk = Assert.Single(reader.Read(malformed, "invalid.txt"));
        Assert.Equal("\ufffd", chunk.Text);
        Assert.Contains(chunk.Warnings!, warning => warning.Contains("invalid byte", StringComparison.Ordinal));
        Assert.Throws<DecoderFallbackException>(() => reader.Read(malformed, "invalid.txt",
            new ReaderOptions { ThrowOnInvalidTextBytes = true }).ToArray());
    }

    [Theory]
    [InlineData(65001, new byte[] { 0xe2, 0x82 })]
    [InlineData(1200, new byte[] { 0x41 })]
    [InlineData(1201, new byte[] { 0x00 })]
    [InlineData(1200, new byte[] { 0x00, 0xd8 })]
    [InlineData(1201, new byte[] { 0xd8, 0x00 })]
    [InlineData(12000, new byte[] { 0x41, 0x00, 0x00 })]
    [InlineData(12001, new byte[] { 0x00, 0x00, 0x00 })]
    public void IncompleteUnicodeSequencesAtEndOfInputAreReportedOrRejected(int codePage, byte[] tail) {
        var encoding = Encoding.GetEncoding(codePage);
        string text = new string('a', 255) + "\U0001F600 tail ";
        byte[] content = encoding.GetBytes(text).Concat(tail).ToArray();
        byte[] withBom = encoding.GetPreamble().Concat(content).ToArray();
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        using var input = new CountingStream(withBom, maxRead: 1);
        var chunks = reader.EnumerateChunks(input, "invalid.txt", new ReaderOptions {
            ComputeHashes = false, MaxChars = 256, TextEncoding = Encoding.ASCII
        }).ToArray();
        Assert.Equal(text + "\ufffd", string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.Contains(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("invalid byte", StringComparison.Ordinal));
        Assert.True(input.CanRead);
        Assert.Throws<DecoderFallbackException>(() => reader.Read(withBom, "invalid.txt",
            new ReaderOptions { ThrowOnInvalidTextBytes = true }).ToArray());
        using var strictInput = new CountingStream(content, maxRead: 1);
        Assert.Throws<DecoderFallbackException>(() => reader.EnumerateChunks(strictInput, "invalid.txt",
            new ReaderOptions { ComputeHashes = false, TextEncoding = encoding, ThrowOnInvalidTextBytes = true }).ToArray());
        Assert.True(strictInput.CanRead);
    }

    private sealed class CountingStream : Stream {
        private readonly MemoryStream _inner;
        private readonly int _maxRead;
        internal CountingStream(byte[] bytes, int maxRead = int.MaxValue) { _inner = new MemoryStream(bytes); _maxRead = maxRead; }
        internal long BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = _inner.Read(buffer, offset, Math.Min(count, _maxRead)); BytesRead += read; return read;
        }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
