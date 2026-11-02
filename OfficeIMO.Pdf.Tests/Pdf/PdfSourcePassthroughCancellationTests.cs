using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSourcePassthroughCancellationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task OpenedSaveStopsAfterDestinationCancelsAndCanBeRetried(bool asynchronous, bool seekable) {
        byte[] source = CreateSource(256 * 1024);
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(source));
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancellingWriteStream(cancellation, seekable);

        if (asynchronous) {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.SaveAsync(destination, cancellationToken: cancellation.Token));
        } else {
            Assert.ThrowsAny<OperationCanceledException>(() => document.Save(destination, cancellationToken: cancellation.Token));
        }

        byte[] partial = destination.ToArray();
        Assert.InRange(partial.Length, 1, source.Length - 1);
        Assert.Equal(source.Take(partial.Length), partial);
        Assert.True(destination.CanWrite);

        using var retry = new MemoryStream();
        PdfSaveResult result = asynchronous
            ? await document.SaveWithEvidenceAsync(retry)
            : document.SaveWithEvidence(retry);

        Assert.True(result.Succeeded);
        Assert.Equal(source, retry.ToArray());
        Assert.Equal(source.LongLength, result.BytesWritten);
        Assert.Equal(0, retry.Position);
        Assert.True(result.Serialization!.SourcePassthrough);
        Assert.False(result.Serialization.FinalArtifactBuffered);
        Assert.Equal(result.Pipeline.Input!.Sha256, result.Pipeline.Output!.Sha256);
        Assert.Equal(1, PdfInspector.Inspect(retry.ToArray()).PageCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OpenedSaveResultPropagatesCancellationDuringItsFinalWrite(bool asynchronous) {
        PdfDocument document = PdfDocument.Load(new System.IO.MemoryStream(CreateSource(16)));
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancellingWriteStream(cancellation, seekable: true);

        if (asynchronous) {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.TrySaveWithEvidenceAsync(destination, cancellation.Token));
        } else {
            Assert.ThrowsAny<OperationCanceledException>(() => document.TrySaveWithEvidence(destination, cancellation.Token));
        }

        Assert.True(destination.CanWrite);
    }

    private static byte[] CreateSource(int attachmentBytes) {
        var payload = new byte[attachmentBytes];
        new Random(71283).NextBytes(payload);
        var options = new PdfCreateOptions().AddEmbeddedFile("payload.bin", payload, "application/octet-stream");
        return PdfDocument.Create(pdf => pdf.Content(content => content
            .Paragraph(paragraph => paragraph.Text("Opened PDF passthrough cancellation."))), options).ToBytes();
    }

    private sealed class CancellingWriteStream(CancellationTokenSource cancellation, bool seekable) : Stream {
        private readonly MemoryStream _inner = new MemoryStream();

        internal byte[] ToArray() => _inner.ToArray();
        public override bool CanRead => false;
        public override bool CanSeek => seekable;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => seekable ? _inner.Length : throw new NotSupportedException();
        public override long Position {
            get => seekable ? _inner.Position : throw new NotSupportedException();
            set { if (!seekable) throw new NotSupportedException(); _inner.Position = value; }
        }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => seekable ? _inner.Seek(offset, origin) : throw new NotSupportedException();
        public override void SetLength(long value) {
            if (!seekable) throw new NotSupportedException();
            _inner.SetLength(value);
        }
        public override void Write(byte[] buffer, int offset, int count) {
            _inner.Write(buffer, offset, count);
            cancellation.Cancel();
        }
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
