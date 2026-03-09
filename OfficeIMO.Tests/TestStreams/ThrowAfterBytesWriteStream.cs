using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Tests.TestStreams {
    internal sealed class ThrowAfterBytesWriteStream : Stream {
        private readonly long _maxBytesBeforeThrow;
        private readonly MemoryStream _inner = new();
        private long _written;

        internal ThrowAfterBytesWriteStream(long maxBytesBeforeThrow) {
            if (maxBytesBeforeThrow < 0) {
                throw new ArgumentOutOfRangeException(nameof(maxBytesBeforeThrow));
            }

            _maxBytesBeforeThrow = maxBytesBeforeThrow;
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
            _inner.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin) {
            throw new NotSupportedException();
        }

        public override void SetLength(long value) {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (_written + count > _maxBytesBeforeThrow) {
                throw new IOException("Simulated stream write failure.");
            }

            _inner.Write(buffer, offset, count);
            _written += count;
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            Write(buffer, offset, count);
            return Task.CompletedTask;
        }
    }
}
