using OfficeIMO.Email;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailContentSourceTests {
    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void WritersConsumeReopenableAttachmentSources(EmailFileFormat format) {
        byte[] payload = Encoding.UTF8.GetBytes("lazy attachment");
        var source = new CountingContentSource(payload);
        var document = new EmailDocument { Format = format, Subject = "lazy" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "lazy.txt",
            ContentType = "text/plain",
            ContentSource = source,
            Length = payload.Length
        });

        byte[] artifact = new EmailDocumentWriter().ToBytes(document, format);
        if (format == EmailFileFormat.OutlookMsg) {
            Assert.True(OfficeIMO.Drawing.Internal.OfficeCompoundFileReader.TryRead(
                artifact, out OfficeIMO.Drawing.Internal.OfficeCompoundFile? compound, out string? compoundError),
                compoundError);
            Assert.True(compound!.Streams.ContainsKey(
                    "__attach_version1.0_#00000000/__substg1.0_37010102"),
                string.Join(" | ", compound.Streams.Keys));
        }
        EmailAttachment attachment = Assert.Single(new EmailDocumentReader().Read(artifact).Document.Attachments);

        Assert.Equal(payload, attachment.Content);
        Assert.Equal(1, source.OpenCount);
    }

    [Fact]
    public async Task AttachmentCanOpenSourceSynchronouslyAndAsynchronously() {
        byte[] payload = new byte[] { 1, 2, 3 };
        var attachment = new EmailAttachment { ContentSource = new CountingContentSource(payload) };

        using Stream synchronous = attachment.OpenContentStream();
        using Stream asynchronous = await attachment.OpenContentStreamAsync();

        Assert.Equal(payload, ReadAll(synchronous));
        Assert.Equal(payload, ReadAll(asynchronous));
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.OutlookTemplate)]
    [InlineData(EmailFileFormat.Tnef)]
    public void WriteStreamsAttachmentContentIntoTheDestination(EmailFileFormat format) {
        const int payloadLength = 1024 * 1024 + 17;
        var source = new GeneratedContentSource(payloadLength, allowSynchronous: true);
        var document = CreateStreamingDocument(source, payloadLength, format);
        using var output = new GuardedWriteStream(maximumWriteSize: 128 * 1024);

        EmailWriteResult result = new EmailDocumentWriter().Write(document, output, format);
        EmailAttachment attachment = Assert.Single(new EmailDocumentReader().Read(output.ToArray()).Document.Attachments);

        Assert.False(result.HasErrors);
        Assert.Equal(payloadLength, attachment.Content!.Length);
        Assert.Equal(1, source.SynchronousOpenCount);
        Assert.True(source.MaximumSynchronousReadSize <= 81920);
        Assert.True(output.MaximumSynchronousWriteSize <= 81920);
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml)]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.OutlookTemplate)]
    [InlineData(EmailFileFormat.Tnef)]
    public async Task WriteAsyncUsesAsyncAttachmentAndDestinationIo(EmailFileFormat format) {
        const int payloadLength = 512 * 1024 + 11;
        var source = new GeneratedContentSource(payloadLength, allowSynchronous: false);
        var document = CreateStreamingDocument(source, payloadLength, format);
        using var output = new AsyncOnlyWriteStream();

        EmailWriteResult result = await new EmailDocumentWriter().WriteAsync(
            document, output, format);
        EmailAttachment attachment = Assert.Single(new EmailDocumentReader().Read(output.ToArray()).Document.Attachments);

        Assert.False(result.HasErrors);
        Assert.Equal(payloadLength, attachment.Content!.Length);
        Assert.Equal(0, source.SynchronousOpenCount);
        Assert.Equal(1, source.AsynchronousOpenCount);
        Assert.True(source.AsynchronousReadCount > 1);
        Assert.True(output.AsynchronousWriteCount > 1);
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml, "message.eml")]
    [InlineData(EmailFileFormat.OutlookMsg, "message.msg")]
    [InlineData(EmailFileFormat.OutlookTemplate, "message.oft")]
    [InlineData(EmailFileFormat.Tnef, "winmail.dat")]
    public void StreamingReaderExternalizesAttachmentContent(EmailFileFormat format, string sourceName) {
        const int payloadLength = 1024 * 1024 + 23;
        var document = CreateStreamingDocument(
            new GeneratedContentSource(payloadLength, allowSynchronous: true), payloadLength, format);
        byte[] artifact = new EmailDocumentWriter().ToBytes(document, format);
        using var input = new MemoryStream(artifact);
        EmailReadResult result = new EmailDocumentReader().ReadStreaming(input, sourceName);
        EmailAttachment attachment = Assert.Single(result.Document.Attachments);

        Assert.True(result.UsesFileBackedContent);
        Assert.Equal(format, result.Document.Format);
        Assert.Null(attachment.Content);
        Assert.NotNull(attachment.ContentSource);
        Assert.Equal(payloadLength, attachment.ContentSource!.Length);
        using (Stream content = attachment.OpenContentStream()) {
            AssertGeneratedContent(content, payloadLength);
        }

        result.Dispose();
        Assert.Throws<ObjectDisposedException>(() => attachment.OpenContentStream());
    }

    [Theory]
    [InlineData(EmailFileFormat.Eml, "message.eml")]
    [InlineData(EmailFileFormat.OutlookMsg, "message.msg")]
    [InlineData(EmailFileFormat.Tnef, "winmail.dat")]
    public async Task StreamingReaderAsyncUsesAsyncSourceIo(EmailFileFormat format, string sourceName) {
        const int payloadLength = 512 * 1024 + 29;
        var document = CreateStreamingDocument(
            new GeneratedContentSource(payloadLength, allowSynchronous: true), payloadLength,
            format);
        byte[] artifact = new EmailDocumentWriter().ToBytes(document, format);
        using var input = new AsyncOnlyReadStream(artifact);
        using EmailReadResult result = await new EmailDocumentReader().ReadStreamingAsync(input, sourceName);
        EmailAttachment attachment = Assert.Single(result.Document.Attachments);

        Assert.True(input.AsynchronousReadCount > 1);
        Assert.Null(attachment.Content);
        Assert.NotNull(attachment.ContentSource);
        using Stream content = await attachment.OpenContentStreamAsync();
        AssertGeneratedContent(content, payloadLength);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void StreamingReaderKeepsLargeAttachmentRetainedMemoryBounded() {
        const int payloadLength = 16 * 1024 * 1024 + 31;
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-stream-budget-", Guid.NewGuid().ToString("N"), ".eml"));
        try {
            var document = CreateStreamingDocument(
                new GeneratedContentSource(payloadLength, allowSynchronous: true), payloadLength,
                EmailFileFormat.Eml);
            new EmailDocumentWriter().Write(document, path, EmailFileFormat.Eml);
            ForceCollection();
            long before = GC.GetTotalMemory(forceFullCollection: false);
            using EmailReadResult result = new EmailDocumentReader(new EmailReaderOptions(
                maxInputBytes: 32L * 1024 * 1024,
                maxAttachmentBytes: 24L * 1024 * 1024,
                maxTotalAttachmentBytes: 24L * 1024 * 1024)).ReadStreaming(path);
            ForceCollection();
            long retainedGrowth = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - before);
            EmailAttachment attachment = Assert.Single(result.Document.Attachments);

            Assert.True(result.UsesFileBackedContent);
            Assert.Null(attachment.Content);
            Assert.Equal(payloadLength, attachment.ContentSource!.Length);
            Assert.True(retainedGrowth <= 8L * 1024 * 1024,
                $"Retained managed memory grew by {retainedGrowth:N0} bytes for a {payloadLength:N0}-byte attachment.");
        } finally {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
#endif

    private static void AssertGeneratedContent(Stream stream, int expectedLength) {
        var buffer = new byte[81920];
        int offset = 0;
        while (true) {
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            for (int index = 0; index < read; index++) {
                Assert.Equal((byte)((offset + index) % 251), buffer[index]);
            }
            offset += read;
        }
        Assert.Equal(expectedLength, offset);
    }

    private static EmailDocument CreateStreamingDocument(IEmailContentSource source, int payloadLength,
        EmailFileFormat format) {
        var document = new EmailDocument { Format = format, Subject = "streaming" };
        document.Attachments.Add(new EmailAttachment {
            FileName = "large.bin",
            ContentType = "application/octet-stream",
            ContentSource = source,
            Length = payloadLength
        });
        return document;
    }

    private static byte[] ReadAll(Stream stream) {
        using var output = new MemoryStream();
        stream.CopyTo(output);
        return output.ToArray();
    }

    private sealed class CountingContentSource : IEmailContentSource {
        private readonly byte[] _content;
        internal CountingContentSource(byte[] content) { _content = content; }
        public long? Length => _content.LongLength;
        internal int OpenCount { get; private set; }
        public Stream OpenRead() { OpenCount++; return new MemoryStream(_content, writable: false); }
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OpenCount++;
            return Task.FromResult<Stream>(new MemoryStream(_content, writable: false));
        }
    }

    private static void ForceCollection() {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    private sealed class GeneratedContentSource : IEmailContentSource {
        private readonly int _length;
        private readonly bool _allowSynchronous;
        internal GeneratedContentSource(int length, bool allowSynchronous) {
            _length = length;
            _allowSynchronous = allowSynchronous;
        }
        public long? Length => _length;
        internal int SynchronousOpenCount { get; private set; }
        internal int AsynchronousOpenCount { get; private set; }
        internal int MaximumSynchronousReadSize { get; private set; }
        internal int AsynchronousReadCount { get; private set; }

        public Stream OpenRead() {
            SynchronousOpenCount++;
            if (!_allowSynchronous) throw new InvalidOperationException("Synchronous attachment access is forbidden.");
            return new GeneratedReadStream(_length, synchronousRead =>
                MaximumSynchronousReadSize = Math.Max(MaximumSynchronousReadSize, synchronousRead), null);
        }

        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            AsynchronousOpenCount++;
            return Task.FromResult<Stream>(new GeneratedReadStream(_length, null,
                () => AsynchronousReadCount++));
        }
    }

    private sealed class GeneratedReadStream : Stream {
        private readonly int _length;
        private readonly Action<int>? _onSynchronousRead;
        private readonly Action? _onAsynchronousRead;
        private int _position;
        internal GeneratedReadStream(int length, Action<int>? onSynchronousRead, Action? onAsynchronousRead) {
            _length = length;
            _onSynchronousRead = onSynchronousRead;
            _onAsynchronousRead = onAsynchronousRead;
        }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position { get => _position; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_onSynchronousRead == null) throw new InvalidOperationException("Synchronous reads are forbidden.");
            _onSynchronousRead(count);
            return Fill(buffer, offset, count);
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (_onAsynchronousRead == null) return base.ReadAsync(buffer, offset, count, cancellationToken);
            _onAsynchronousRead();
            return Task.FromResult(Fill(buffer, offset, count));
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        private int Fill(byte[] buffer, int offset, int count) {
            int available = Math.Min(count, _length - _position);
            for (int i = 0; i < available; i++) buffer[offset + i] = (byte)((_position + i) % 251);
            _position += available;
            return available;
        }
    }

    private sealed class GuardedWriteStream : MemoryStream {
        private readonly int _maximumWriteSize;
        internal GuardedWriteStream(int maximumWriteSize) { _maximumWriteSize = maximumWriteSize; }
        internal int MaximumSynchronousWriteSize { get; private set; }
        public override void Write(byte[] buffer, int offset, int count) {
            if (count > _maximumWriteSize) throw new InvalidOperationException("The writer attempted a whole-artifact write.");
            MaximumSynchronousWriteSize = Math.Max(MaximumSynchronousWriteSize, count);
            base.Write(buffer, offset, count);
        }
    }

    private sealed class AsyncOnlyWriteStream : Stream {
        private readonly MemoryStream _content = new MemoryStream();
        internal int AsynchronousWriteCount { get; private set; }
        internal byte[] ToArray() => _content.ToArray();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _content.Length;
        public override long Position { get => _content.Position; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new InvalidOperationException("Synchronous destination writes are forbidden.");
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            AsynchronousWriteCount++;
            _content.Write(buffer, offset, count);
            return Task.CompletedTask;
        }
        protected override void Dispose(bool disposing) {
            if (disposing) _content.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class AsyncOnlyReadStream : Stream {
        private readonly byte[] _content;
        private int _position;
        internal AsyncOnlyReadStream(byte[] content) { _content = content; }
        internal int AsynchronousReadCount { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _content.LongLength;
        public override long Position { get => _position; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) =>
            throw new InvalidOperationException("Synchronous source reads are forbidden.");
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            AsynchronousReadCount++;
            int copy = Math.Min(count, _content.Length - _position);
            if (copy > 0) Buffer.BlockCopy(_content, _position, buffer, offset, copy);
            _position += copy;
            return Task.FromResult(copy);
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
