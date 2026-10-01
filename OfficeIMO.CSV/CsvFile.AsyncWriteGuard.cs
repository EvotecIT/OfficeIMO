#nullable enable
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal static partial class CsvFile
{
    // StreamWriter/compression disposal can flush after a failed async write.
    // Stop that cleanup flush from publishing buffered data or masking cancellation.
    internal sealed class AsyncWriteGuard : Stream
    {
        private readonly Stream _destination;
        private readonly CancellationToken _operationToken;
        private readonly MemoryStream _pending = new MemoryStream();
        internal AsyncWriteGuard(Stream destination, CancellationToken operationToken) {
            _destination = destination;
            _operationToken = operationToken;
        }
        internal bool SuppressWrites { get; set; }
        public override bool CanRead => false;
        public override bool CanSeek => _destination.CanSeek;
        public override bool CanWrite => _destination.CanWrite;
        public override long Length => _destination.Length;
        public override long Position { get => _destination.Position; set => _destination.Position = value; }
        public override void Flush() { }
        public override async Task FlushAsync(CancellationToken cancellationToken) {
            if (SuppressWrites) return;
            using var linked = LinkToken(cancellationToken, out CancellationToken effectiveToken);
            await DrainAsync(effectiveToken).ConfigureAwait(false);
            await _destination.FlushAsync(effectiveToken).ConfigureAwait(false);
        }
        public override void Write(byte[] buffer, int offset, int count) { if (!SuppressWrites) _pending.Write(buffer, offset, count); }
        public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            if (SuppressWrites) return;
            using var linked = LinkToken(cancellationToken, out CancellationToken effectiveToken);
            await DrainAsync(effectiveToken).ConfigureAwait(false);
            await _destination.WriteAsync(buffer, offset, count, effectiveToken).ConfigureAwait(false);
        }
#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) { if (!SuppressWrites) _pending.Write(buffer); }
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) {
            if (SuppressWrites) return;
            using var linked = LinkToken(cancellationToken, out CancellationToken effectiveToken);
            await DrainAsync(effectiveToken).ConfigureAwait(false);
            await _destination.WriteAsync(buffer, effectiveToken).ConfigureAwait(false);
        }
#endif
        internal Task DrainAsync() => DrainAsync(_operationToken);
        private CancellationTokenSource? LinkToken(CancellationToken token, out CancellationToken effectiveToken) {
            if (!token.CanBeCanceled || token == _operationToken) { effectiveToken = _operationToken; return null; }
            if (!_operationToken.CanBeCanceled) { effectiveToken = token; return null; }
            var linked = CancellationTokenSource.CreateLinkedTokenSource(_operationToken, token);
            effectiveToken = linked.Token;
            return linked;
        }
        private async Task DrainAsync(CancellationToken token) {
            if (SuppressWrites || _pending.Length == 0) return;
            token.ThrowIfCancellationRequested();
            if (!_pending.TryGetBuffer(out ArraySegment<byte> bytes)) throw new InvalidOperationException("Async output buffer is inaccessible.");
            await _destination.WriteAsync(bytes.Array!, bytes.Offset, bytes.Count, token).ConfigureAwait(false);
            _pending.SetLength(0);
            _pending.Position = 0;
        }
        public override long Seek(long offset, SeekOrigin origin) => _destination.Seek(offset, origin);
        public override void SetLength(long value) => _destination.SetLength(value);
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _pending.Dispose();
            base.Dispose(disposing);
        }
    }
}
