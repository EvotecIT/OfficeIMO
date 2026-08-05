using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Read {
    /// <summary>Bounded random-access view over a BIFF workbook stream.</summary>
    internal sealed class LegacyBiffSource : IDisposable {
        private const int PageSize = 32 * 1024;
        private const int MaximumPooledBufferSize = 64 * 1024 * 1024;
        private Stream? _stream;
        private byte[]? _buffer;
        private byte[]? _page;
        private long _pageOffset = -1;
        private int _pageLength;
        private bool _disposed;

        internal LegacyBiffSource(Stream stream, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead || !stream.CanSeek) {
                stream.Dispose();
                throw new ArgumentException("A BIFF source must be readable and seekable.", nameof(stream));
            }
            if (stream.Length > int.MaxValue) {
                stream.Dispose();
                throw new InvalidDataException("A BIFF workbook stream cannot exceed 2 GiB.");
            }
            Length = checked((int)stream.Length);
            if (Length > MaximumPooledBufferSize) {
                _stream = stream;
                _page = new byte[PageSize];
                return;
            }

            byte[] buffer = ArrayPool<byte>.Shared.Rent(Math.Max(1, Length));
            try {
                stream.Position = 0;
                int offset = 0;
                while (offset < Length) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = stream.Read(buffer, offset, Length - offset);
                    if (read == 0) {
                        throw new EndOfStreamException("The BIFF stream ended before its declared length.");
                    }
                    offset += read;
                }
                cancellationToken.ThrowIfCancellationRequested();
                _buffer = buffer;
            } catch {
                ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                throw;
            } finally {
                stream.Dispose();
            }
        }

        internal int Length { get; }

        /// <summary>
        /// Gets the contiguous workbook buffer when the source was small enough to load eagerly.
        /// Callers must treat the pooled buffer as immutable and observe <see cref="Length"/>.
        /// </summary>
        internal byte[]? ContiguousBuffer => _buffer;

        internal byte this[int offset] => ReadByte(offset);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal byte ReadByte(int offset) {
            if ((uint)offset >= (uint)Length) throw new InvalidDataException("Unexpected end of BIFF record.");
            byte[]? buffer = _buffer;
            if (buffer != null) return buffer[offset];
            EnsurePage(offset);
            return _page![offset - checked((int)_pageOffset)];
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal ushort ReadUInt16(int offset) {
            if (offset < 0 || offset > Length - sizeof(ushort)) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            byte[]? buffer = _buffer;
            return buffer != null
                ? (ushort)(buffer[offset] | buffer[offset + 1] << 8)
                : ReadUInt16Paged(offset);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal uint ReadUInt32(int offset) {
            if (offset < 0 || offset > Length - sizeof(uint)) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            byte[]? buffer = _buffer;
            return buffer != null
                ? (uint)(buffer[offset]
                    | buffer[offset + 1] << 8
                    | buffer[offset + 2] << 16
                    | buffer[offset + 3] << 24)
                : ReadUInt32Paged(offset);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal double ReadDouble(int offset) {
            if (offset < 0 || offset > Length - sizeof(long)) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            byte[]? buffer = _buffer;
            if (buffer != null) {
                return BitConverter.Int64BitsToDouble(
                    BinaryPrimitives.ReadInt64LittleEndian(buffer.AsSpan(offset, sizeof(long))));
            }
            ulong bits = ReadUInt32(offset) | (ulong)ReadUInt32(checked(offset + 4)) << 32;
            return BitConverter.Int64BitsToDouble(unchecked((long)bits));
        }

        internal byte[] Copy(int offset, int count) {
            ThrowIfDisposed();
            if (offset < 0 || count < 0 || offset > Length - count) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            var result = new byte[count];
            if (_buffer != null) {
                Buffer.BlockCopy(_buffer, offset, result, 0, count);
                return result;
            }
            int copied = 0;
            while (copied < count) {
                EnsurePage(checked(offset + copied));
                int within = checked(offset + copied - (int)_pageOffset);
                int take = Math.Min(count - copied, _pageLength - within);
                Buffer.BlockCopy(_page!, within, result, copied, take);
                copied += take;
            }
            return result;
        }

        private void EnsurePage(int offset) {
            ThrowIfDisposed();
            if (_pageOffset >= 0 && offset >= _pageOffset && offset < _pageOffset + _pageLength) return;
            long pageOffset = offset - offset % PageSize;
            _stream!.Position = pageOffset;
            int wanted = Math.Min(PageSize, Length - checked((int)pageOffset));
            int total = 0;
            while (total < wanted) {
                int read = _stream.Read(_page!, total, wanted - total);
                if (read <= 0) throw new EndOfStreamException("The BIFF stream ended before its declared length.");
                total += read;
            }
            _pageOffset = pageOffset;
            _pageLength = total;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private ushort ReadUInt16Paged(int offset) =>
            (ushort)(ReadByte(offset) | ReadByte(checked(offset + 1)) << 8);

        [MethodImpl(MethodImplOptions.NoInlining)]
        private uint ReadUInt32Paged(int offset) =>
            (uint)(ReadByte(offset)
                | ReadByte(checked(offset + 1)) << 8
                | ReadByte(checked(offset + 2)) << 16
                | ReadByte(checked(offset + 3)) << 24);

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            _stream?.Dispose();
            _stream = null;
            byte[]? buffer = _buffer;
            _buffer = null;
            if (buffer != null) {
                Array.Clear(buffer, 0, Length);
                ArrayPool<byte>.Shared.Return(buffer);
            }
            _page = null;
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(LegacyBiffSource));
        }
    }
}
