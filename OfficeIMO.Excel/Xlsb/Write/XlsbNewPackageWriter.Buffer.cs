using System.Buffers;

namespace OfficeIMO.Excel.Xlsb.Write {
    internal static partial class XlsbNewPackageWriter {
        /// <summary>
        /// Owns a compressed package until validation finishes. Fixed pooled chunks
        /// avoid growth copies; seeking lets ZipArchive finalize its ordinary headers.
        /// Only ZIP creation and publication use this private stream.
        /// </summary>
        private sealed class XlsbPackageBuffer : Stream {
            private const int ChunkBits = 15;
            private const int ChunkSize = 1 << ChunkBits;
            private const int ChunkMask = ChunkSize - 1;
            private readonly List<byte[]> _chunks = new List<byte[]>();
            private int _length;
            private int _position;
            private bool _disposed;

            public override bool CanRead => false;
            public override bool CanSeek => !_disposed;
            public override bool CanWrite => !_disposed;
            public override long Length { get { ThrowIfDisposed(); return _length; } }
            public override long Position {
                get { ThrowIfDisposed(); return _position; }
                set {
                    ThrowIfDisposed();
                    if (value < 0 || value > _length) throw new ArgumentOutOfRangeException(nameof(value));
                    _position = (int)value;
                }
            }

            public override void Write(byte[] buffer, int offset, int count) {
                ThrowIfDisposed();
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                    throw new ArgumentOutOfRangeException(nameof(count));
                }
                WriteCore(buffer.AsSpan(offset, count));
            }

#if NET8_0_OR_GREATER
            public override void Write(ReadOnlySpan<byte> buffer) => WriteCore(buffer);
#endif

            private void WriteCore(ReadOnlySpan<byte> buffer) {
                ThrowIfDisposed();
                if ((long)_position + buffer.Length > int.MaxValue) {
                    throw new IOException("The native XLSB package exceeds the in-memory stream limit.");
                }
                while (!buffer.IsEmpty) {
                    int chunk = _position >> ChunkBits;
                    if (chunk == _chunks.Count) _chunks.Add(ArrayPool<byte>.Shared.Rent(ChunkSize));
                    int target = _position & ChunkMask;
                    int copied = Math.Min(buffer.Length, ChunkSize - target);
                    buffer.Slice(0, copied).CopyTo(_chunks[chunk].AsSpan(target, copied));
                    _position += copied;
                    buffer = buffer.Slice(copied);
                }
                _length = Math.Max(_length, _position);
            }

            public override long Seek(long offset, SeekOrigin origin) {
                long start;
                switch (origin) {
                    case SeekOrigin.Begin: start = 0; break;
                    case SeekOrigin.Current: start = Position; break;
                    case SeekOrigin.End: start = Length; break;
                    default: throw new ArgumentOutOfRangeException(nameof(origin));
                }
                Position = checked(start + offset);
                return _position;
            }

            internal void WriteTo(Stream destination) {
                ThrowIfDisposed();
                int remaining = _length;
                for (int index = 0; remaining != 0; index++) {
                    int count = Math.Min(ChunkSize, remaining);
                    destination.Write(_chunks[index], 0, count);
                    remaining -= count;
                }
            }

            public override void Flush() => ThrowIfDisposed();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                if (!_disposed) {
                    _disposed = true;
                    foreach (byte[] chunk in _chunks) ArrayPool<byte>.Shared.Return(chunk, clearArray: true);
                    _chunks.Clear();
                }
                base.Dispose(disposing);
            }

            private void ThrowIfDisposed() {
                if (_disposed) throw new ObjectDisposedException(nameof(XlsbPackageBuffer));
            }
        }
    }
}
