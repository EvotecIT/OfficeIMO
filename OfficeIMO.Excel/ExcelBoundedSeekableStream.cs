using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>Bounds a seekable staging stream while retaining package-writer compatibility.</summary>
    internal sealed class ExcelBoundedSeekableStream : Stream {
        private readonly Stream _inner;
        private readonly long _maximumBytes;
        private readonly bool _leaveOpen;
        private readonly CancellationToken _cancellationToken;
        private readonly Func<long, Exception>? _limitExceededExceptionFactory;
        private readonly bool _restoreEmptyStreamOnFailure;
        private bool _completed;

        internal ExcelBoundedSeekableStream(
            Stream inner,
            long maximumBytes,
            bool leaveOpen = false,
            CancellationToken cancellationToken = default,
            Func<long, Exception>? limitExceededExceptionFactory = null,
            bool restoreEmptyStreamOnFailure = false) {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            if (!inner.CanSeek || !inner.CanWrite) {
                throw new ArgumentException("The staging stream must be seekable and writable.", nameof(inner));
            }
            if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            _maximumBytes = maximumBytes;
            _leaveOpen = leaveOpen;
            _cancellationToken = cancellationToken;
            _limitExceededExceptionFactory = limitExceededExceptionFactory;
            _restoreEmptyStreamOnFailure = restoreEmptyStreamOnFailure;
            if (restoreEmptyStreamOnFailure && (inner.Length != 0 || inner.Position != 0)) {
                throw new ArgumentException("Rollback to an empty stream requires an empty destination positioned at the beginning.", nameof(inner));
            }
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set {
                EnsureWithinLimit(value);
                _inner.Position = value;
            }
        }

        public override void Flush() => _inner.Flush();
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) {
            long position;
            switch (origin) {
                case SeekOrigin.Begin:
                    position = offset;
                    break;
                case SeekOrigin.Current:
                    position = checked(_inner.Position + offset);
                    break;
                case SeekOrigin.End:
                    position = checked(_inner.Length + offset);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(origin));
            }
            EnsureWithinLimit(position);
            return _inner.Seek(offset, origin);
        }

        public override void SetLength(long value) {
            _cancellationToken.ThrowIfCancellationRequested();
            EnsureWithinLimit(value);
            _inner.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count) {
            EnsureWriteWithinLimit(count);
            _inner.Write(buffer, offset, count);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            EnsureWriteWithinLimit(count);
            return _inner.WriteAsync(buffer, offset, count, cancellationToken);
        }

        public override void WriteByte(byte value) {
            EnsureWriteWithinLimit(1);
            _inner.WriteByte(value);
        }

        internal void Complete() => _completed = true;

        protected override void Dispose(bool disposing) {
            if (disposing && _restoreEmptyStreamOnFailure && !_completed) {
                try {
                    _inner.Position = 0;
                    _inner.SetLength(0);
                } catch {
                }
            }

            if (disposing && !_leaveOpen) _inner.Dispose();
            base.Dispose(disposing);
        }

        private void EnsureWriteWithinLimit(int count) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            EnsureWithinLimit(checked(_inner.Position + count));
        }

        private void EnsureWithinLimit(long value) {
            if (value < 0 || value > _maximumBytes) {
                if (_limitExceededExceptionFactory != null) {
                    throw _limitExceededExceptionFactory(_maximumBytes);
                }

                throw new IOException($"The staged Excel package exceeds the {_maximumBytes}-byte temporary package limit.");
            }
        }
    }
}
