#nullable enable

using System.IO;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Owns one open file identity and creates independent-position read views over it.
    /// This keeps package metadata and fast ZIP-part reads on the same snapshot even when
    /// the path is atomically replaced while the workbook is open.
    /// </summary>
    internal sealed class SharedReadOnlyFileSnapshot : IDisposable {
        private readonly object _gate = new object();
        private readonly FileStream _stream;
        private readonly long _length;
        private bool _disposed;

        private SharedReadOnlyFileSnapshot(FileStream stream) {
            _stream = stream;
            _length = stream.Length;
        }

        internal long Length => _length;

        internal static SharedReadOnlyFileSnapshot Open(string path) {
            var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete,
                bufferSize: 1,
                FileOptions.RandomAccess);
            try {
                return new SharedReadOnlyFileSnapshot(stream);
            } catch {
                stream.Dispose();
                throw;
            }
        }

        internal Stream CreateView(int bufferSize = 4096) {
            ThrowIfDisposed();
            var view = new View(this);
            return bufferSize > 1
                ? new BufferedStream(view, bufferSize)
                : view;
        }

        private int Read(long position, byte[] buffer, int offset, int count) {
            ThrowIfDisposed();
            if (position >= _length || count == 0) {
                return 0;
            }
            count = checked((int)Math.Min(count, _length - position));
#if NET8_0_OR_GREATER
            return RandomAccess.Read(_stream.SafeFileHandle, buffer.AsSpan(offset, count), position);
#else
            lock (_gate) {
                ThrowIfDisposed();
                _stream.Position = position;
                return _stream.Read(buffer, offset, count);
            }
#endif
        }

#if NET8_0_OR_GREATER
        private int Read(long position, Span<byte> buffer) {
            ThrowIfDisposed();
            if (position >= _length || buffer.Length == 0) {
                return 0;
            }
            int count = checked((int)Math.Min(buffer.Length, _length - position));
            return RandomAccess.Read(_stream.SafeFileHandle, buffer.Slice(0, count), position);
        }
#endif

        private void ThrowIfDisposed() {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(SharedReadOnlyFileSnapshot));
            }
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }
            _disposed = true;
            _stream.Dispose();
        }

        private sealed class View : Stream {
            private readonly SharedReadOnlyFileSnapshot _snapshot;
            private long _position;
            private bool _disposed;

            internal View(SharedReadOnlyFileSnapshot snapshot) {
                _snapshot = snapshot;
            }

            public override bool CanRead => !_disposed;
            public override bool CanSeek => !_disposed;
            public override bool CanWrite => false;
            public override long Length {
                get {
                    ThrowIfDisposed();
                    return _snapshot._length;
                }
            }
            public override long Position {
                get {
                    ThrowIfDisposed();
                    return _position;
                }
                set {
                    ThrowIfDisposed();
                    if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
                    _position = value;
                }
            }

            public override int Read(byte[] buffer, int offset, int count) {
                ThrowIfDisposed();
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                    throw new ArgumentOutOfRangeException();
                }
                int read = _snapshot.Read(_position, buffer, offset, count);
                _position = checked(_position + read);
                return read;
            }

#if NET8_0_OR_GREATER
            public override int Read(Span<byte> buffer) {
                ThrowIfDisposed();
                int read = _snapshot.Read(_position, buffer);
                _position = checked(_position + read);
                return read;
            }
#endif

            public override long Seek(long offset, SeekOrigin origin) {
                ThrowIfDisposed();
                long position = origin switch {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => checked(_position + offset),
                    SeekOrigin.End => checked(_snapshot._length + offset),
                    _ => throw new ArgumentOutOfRangeException(nameof(origin))
                };
                if (position < 0) throw new IOException("Attempted to seek before the file snapshot.");
                _position = position;
                return position;
            }

            public override void Flush() { }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                _disposed = true;
                base.Dispose(disposing);
            }

            private void ThrowIfDisposed() {
                if (_disposed) {
                    throw new ObjectDisposedException(nameof(View));
                }
            }
        }
    }
}
