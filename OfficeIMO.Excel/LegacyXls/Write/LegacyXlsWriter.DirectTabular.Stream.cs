using System.Buffers;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        internal sealed class DirectTabularWorkbookStream : Stream {
            private readonly ArrayPool<byte> _pool;
            private byte[]? _buffer;
            private int _length;
            private int _position;
            private int _clearLength;

            internal DirectTabularWorkbookStream(int initialCapacity, ArrayPool<byte>? pool = null) {
                _pool = pool ?? ArrayPool<byte>.Shared;
                _buffer = _pool.Rent(Math.Max(1, initialCapacity));
            }

            internal byte[] Buffer => _buffer ?? throw new ObjectDisposedException(nameof(DirectTabularWorkbookStream));

            public override bool CanRead => false;

            public override bool CanSeek => _buffer != null;

            public override bool CanWrite => _buffer != null;

            public override long Length {
                get {
                    ThrowIfDisposed();
                    return _length;
                }
            }

            public override long Position {
                get {
                    ThrowIfDisposed();
                    return _position;
                }
                set {
                    ThrowIfDisposed();
                    if (value < 0 || value > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
                    _position = checked((int)value);
                }
            }

            internal void EnsureCapacity(int requiredCapacity) {
                byte[] buffer = Buffer;
                if (requiredCapacity <= buffer.Length) return;

                int doubledCapacity = buffer.Length <= int.MaxValue / 2
                    ? buffer.Length * 2
                    : int.MaxValue;
                int newCapacity = Math.Max(requiredCapacity, doubledCapacity);
                byte[] expanded = _pool.Rent(newCapacity);
                System.Buffer.BlockCopy(buffer, 0, expanded, 0, _clearLength);
                Array.Clear(buffer, 0, _clearLength);
                _buffer = expanded;
                _pool.Return(buffer);
            }

            internal void RegisterDirectWriteUpperBound(int exclusiveEnd) {
                byte[] buffer = Buffer;
                if (exclusiveEnd < _position || exclusiveEnd > buffer.Length) {
                    throw new ArgumentOutOfRangeException(nameof(exclusiveEnd));
                }
                if (exclusiveEnd > _clearLength) _clearLength = exclusiveEnd;
            }

            internal byte[] ToArray() {
                var bytes = new byte[_length];
                System.Buffer.BlockCopy(Buffer, 0, bytes, 0, _length);
                return bytes;
            }

            public override void Flush() {
                ThrowIfDisposed();
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) {
                ThrowIfDisposed();
                long target = origin switch {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => checked((long)_position + offset),
                    SeekOrigin.End => checked((long)_length + offset),
                    _ => throw new ArgumentOutOfRangeException(nameof(origin))
                };
                Position = target;
                return target;
            }

            public override void SetLength(long value) {
                ThrowIfDisposed();
                if (value < 0 || value > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
                int newLength = checked((int)value);
                EnsureCapacity(newLength);
                if (newLength > _length) {
                    Array.Clear(Buffer, _length, newLength - _length);
                }
                _length = newLength;
                if (newLength > _clearLength) _clearLength = newLength;
                if (_position > newLength) _position = newLength;
            }

            public override void Write(byte[] buffer, int offset, int count) {
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || buffer.Length - offset < count) {
                    throw new ArgumentOutOfRangeException(offset < 0 ? nameof(offset) : nameof(count));
                }

                int end = checked(_position + count);
                EnsureCapacity(end);
                System.Buffer.BlockCopy(buffer, offset, Buffer, _position, count);
                _position = end;
                if (end > _length) _length = end;
                if (end > _clearLength) _clearLength = end;
            }

            public override void WriteByte(byte value) {
                int end = checked(_position + 1);
                EnsureCapacity(end);
                Buffer[_position] = value;
                _position = end;
                if (end > _length) _length = end;
                if (end > _clearLength) _clearLength = end;
            }

            protected override void Dispose(bool disposing) {
                byte[]? buffer = _buffer;
                if (buffer != null) {
                    _buffer = null;
                    Array.Clear(buffer, 0, _clearLength);
                    _pool.Return(buffer);
                    _length = 0;
                    _position = 0;
                    _clearLength = 0;
                }
                base.Dispose(disposing);
            }

            private void ThrowIfDisposed() {
                if (_buffer == null) throw new ObjectDisposedException(nameof(DirectTabularWorkbookStream));
            }
        }
    }
}
