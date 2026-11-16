using System.Buffers;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Read-only seekable buffering over an unbuffered stream, with the buffer rented from the shared pool. A
    /// <see cref="FileStream"/> allocates its own buffer for every open; reading a package opens the file once per
    /// read, so the pooled buffer keeps that allocation out of each read.
    /// </summary>
    /// <remarks>
    /// Like the stream it replaces, it serves one reader at a time. The buffer goes back to the pool only once no read is
    /// using it, so a dispose that races a background read never hands live memory to another renter.
    /// </remarks>
    internal sealed class ExcelPooledReadBufferStream : Stream {
        private const int Idle = 0;
        private const int Reading = 1;
        private const int Released = 2;

        private readonly Stream _inner;
        private byte[]? _buffer;
        private long _bufferStart;
        private int _bufferLength;
        private long _position;
        private int _state;
        private bool _disposed;

        internal ExcelPooledReadBufferStream(Stream inner, int bufferSize) {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            if (!inner.CanRead || !inner.CanSeek) throw new ArgumentException("The stream must be readable and seekable.", nameof(inner));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(bufferSize);
            _buffer = ArrayPool<byte>.Shared.Rent(bufferSize);
            _position = inner.Position;
        }

        /// <summary>Opens a file for sequential, seekable reads through a pooled buffer.</summary>
        internal static ExcelPooledReadBufferStream OpenFile(string path, int bufferSize) {
            // A buffer size of one turns off the file stream's own buffer.
            var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1);
            try {
                return new ExcelPooledReadBufferStream(file, bufferSize);
            } catch {
                file.Dispose();
                throw;
            }
        }

        public override bool CanRead => !_disposed;
        public override bool CanSeek => !_disposed;
        public override bool CanWrite => false;

        public override long Length {
            get {
                ThrowIfDisposed();
                return _inner.Length;
            }
        }

        public override long Position {
            get {
                ThrowIfDisposed();
                return _position;
            }
            set {
                ThrowIfDisposed();
                ArgumentOutOfRangeException.ThrowIfNegative(value);
                _position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            ArgumentNullException.ThrowIfNull(buffer);
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
            EnterRead();
            try {
                return ReadCore(buffer, offset, count);
            } finally {
                ExitRead();
            }
        }

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer) {
            EnterRead();
            try {
                if (buffer.Length == 0) return 0;
                int buffered = CopyBuffered(buffer);
                if (buffered > 0) return buffered;
                if (buffer.Length >= _buffer!.Length) {
                    _inner.Position = _position;
                    int direct = _inner.Read(buffer);
                    _position += direct;
                    return direct;
                }
                if (!Refill()) return 0;
                return CopyBuffered(buffer);
            } finally {
                ExitRead();
            }
        }

        private int CopyBuffered(Span<byte> destination) {
            long index = _position - _bufferStart;
            if (index < 0 || index >= _bufferLength) return 0;
            int take = Math.Min(destination.Length, _bufferLength - (int)index);
            _buffer.AsSpan((int)index, take).CopyTo(destination);
            _position += take;
            return take;
        }
#endif

        public override int ReadByte() {
            EnterRead();
            try {
                long index = _position - _bufferStart;
                if (index < 0 || index >= _bufferLength) {
                    if (!Refill()) return -1;
                    index = 0;
                }
                _position++;
                return _buffer![(int)index];
            } finally {
                ExitRead();
            }
        }

        private int ReadCore(byte[] destination, int offset, int count) {
            if (count == 0) return 0;
            int buffered = CopyBuffered(destination, offset, count);
            if (buffered > 0) return buffered;
            // A request at least as large as the buffer gains nothing from a copy through it.
            if (count >= _buffer!.Length) {
                _inner.Position = _position;
                int direct = _inner.Read(destination, offset, count);
                _position += direct;
                return direct;
            }
            if (!Refill()) return 0;
            return CopyBuffered(destination, offset, count);
        }

        private int CopyBuffered(byte[] destination, int offset, int count) {
            long index = _position - _bufferStart;
            if (index < 0 || index >= _bufferLength) return 0;
            int take = Math.Min(count, _bufferLength - (int)index);
            Buffer.BlockCopy(_buffer!, (int)index, destination, offset, take);
            _position += take;
            return take;
        }

        private bool Refill() {
            byte[] buffer = _buffer!;
            _inner.Position = _position;
            int read = _inner.Read(buffer, 0, buffer.Length);
            _bufferStart = _position;
            _bufferLength = read;
            return read > 0;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            ThrowIfDisposed();
            long target = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => _inner.Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            if (target < 0) throw new IOException("An attempt was made to move the position before the beginning of the stream.");
            _position = target;
            return target;
        }

        public override void Flush() { }

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing && !_disposed) {
                _disposed = true;
                _inner.Dispose();
                if (Interlocked.CompareExchange(ref _state, Released, Idle) == Idle) ReturnBuffer();
            }
            base.Dispose(disposing);
        }

        private void EnterRead() {
            int state = Interlocked.CompareExchange(ref _state, Reading, Idle);
            if (state == Idle) return;
            if (state == Released) throw new ObjectDisposedException(nameof(ExcelPooledReadBufferStream));
            throw new InvalidOperationException("The stream does not support concurrent reads.");
        }

        private void ExitRead() {
            Volatile.Write(ref _state, Idle);
            // A dispose that arrived during this read left the buffer for the read to return.
            if (_disposed && Interlocked.CompareExchange(ref _state, Released, Idle) == Idle) ReturnBuffer();
        }

        private void ReturnBuffer() {
            byte[]? buffer = Interlocked.Exchange(ref _buffer, null);
            if (buffer != null) ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(ExcelPooledReadBufferStream));
        }
    }
}
