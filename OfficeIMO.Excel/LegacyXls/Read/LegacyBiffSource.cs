using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.LegacyXls.Read {
    /// <summary>Bounded random-access view over a BIFF workbook stream.</summary>
    internal sealed class LegacyBiffSource : IDisposable {
        private const int PageSize = 32 * 1024;
        private const int MaximumPooledBufferSize = 64 * 1024 * 1024;
        private readonly CancellationToken _cancellationToken;
        private Stream? _stream;
        private byte[]? _buffer;
        private byte[]? _page;
        private long _pageOffset = -1;
        private int _pageLength;
        // The buffer holds bytes [0, _bufferedLength); a deferred source pages the rest from its stream and a
        // slice reads the rest from its parent.
        private int _bufferedLength;
        private readonly LegacyBiffSource? _parent;
        private readonly int _parentOffset;
        private readonly bool _deferred;
        private bool _disposed;

        internal LegacyBiffSource(Stream stream, CancellationToken cancellationToken = default)
            : this(stream, deferLoading: false, cancellationToken) {
        }

        /// <summary>
        /// Opens the stream without loading it, for a reader of one worksheet: records are paged until
        /// <see cref="LoadGlobals"/> loads the workbook globals, and the worksheet is read through a
        /// <see cref="Slice"/>.
        /// </summary>
        internal static LegacyBiffSource OpenDeferred(Stream stream, CancellationToken cancellationToken) =>
            new(stream, deferLoading: true, cancellationToken);

        internal bool IsDeferred => _deferred;

        /// <summary>Whether <see cref="Slice"/> can buffer a substream of <paramref name="length"/> bytes.</summary>
        internal static bool CanSlice(int length) => length <= MaximumPooledBufferSize;

        private LegacyBiffSource(LegacyBiffSource parent, int start, int bufferedEnd) {
            _parent = parent;
            _parentOffset = start;
            _cancellationToken = parent._cancellationToken;
            Length = parent.Length - start;
            int count = bufferedEnd - start;
            byte[] buffer = ArrayPool<byte>.Shared.Rent(Math.Max(1, count));
            try {
                parent.CopyTo(start, buffer, count);
            } catch {
                ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                throw;
            }
            _buffer = buffer;
            _bufferedLength = count;
        }

        private LegacyBiffSource(Stream stream, bool deferLoading, CancellationToken cancellationToken) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead || !stream.CanSeek) {
                stream.Dispose();
                throw new ArgumentException("A BIFF source must be readable and seekable.", nameof(stream));
            }
            if (stream.Length > int.MaxValue) {
                stream.Dispose();
                throw new InvalidDataException("A BIFF workbook stream cannot exceed 2 GiB.");
            }
            _cancellationToken = cancellationToken;
            Length = checked((int)stream.Length);
            if (deferLoading || Length > MaximumPooledBufferSize) {
                _stream = stream;
                _page = new byte[PageSize];
                _deferred = deferLoading && Length <= MaximumPooledBufferSize;
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
                _bufferedLength = Length;
            } catch {
                ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                throw;
            } finally {
                stream.Dispose();
            }
        }

        internal int Length { get; }

        /// <summary>
        /// Gets the contiguous buffer, which holds every byte of a loaded source and of a slice's worksheet.
        /// A slice's reader indexes it only within its worksheet substream. Callers must treat the pooled
        /// buffer as immutable and observe <see cref="Length"/>.
        /// </summary>
        internal byte[]? ContiguousBuffer => _bufferedLength == Length || _parent != null ? _buffer : null;

        internal byte this[int offset] => ReadByte(offset);

        // Offsets are validated against Length before this check, so the sum cannot overflow.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool IsBuffered(int offset, int count) => offset + count <= _bufferedLength;

        /// <summary>Buffers the workbook globals of a deferred source, which end before <paramref name="globalsEnd"/>.</summary>
        internal void LoadGlobals(int globalsEnd) {
            ThrowIfDisposed();
            if (!IsDeferred || _buffer != null) return;
            int count = Math.Min(Math.Max(globalsEnd, 0), Length);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(Math.Max(1, count));
            try {
                CopyTo(0, buffer, count);
            } catch {
                ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                throw;
            }
            _buffer = buffer;
            _bufferedLength = count;
        }

        /// <summary>
        /// Copies the worksheet substream [<paramref name="start"/>, <paramref name="end"/>) into a source whose
        /// offset 0 is <paramref name="start"/>. The slice keeps the rest of this stream as its length, so a
        /// record that crosses the substream end is reported as it is here; those bytes come from this source.
        /// </summary>
        internal LegacyBiffSource Slice(int start, int end) {
            ThrowIfDisposed();
            // A record header may straddle the substream end; keep it in the buffer.
            return new LegacyBiffSource(this, start, (int)Math.Min(Length, (long)end + 4));
        }

        private void CopyTo(int offset, byte[] destination, int count) {
            if (_buffer != null && IsBuffered(offset, count)) {
                Buffer.BlockCopy(_buffer, offset, destination, 0, count);
                return;
            }
            if (_parent != null) {
                _parent.CopyTo(checked(offset + _parentOffset), destination, count);
                return;
            }
            _stream!.Position = offset;
            int copied = 0;
            while (copied < count) {
                _cancellationToken.ThrowIfCancellationRequested();
                int read = _stream.Read(destination, copied, count - copied);
                if (read == 0) throw new EndOfStreamException("The BIFF stream ended before its declared length.");
                copied += read;
            }
            _pageOffset = -1;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal byte ReadByte(int offset) {
            if ((uint)offset >= (uint)Length) throw new InvalidDataException("Unexpected end of BIFF record.");
            byte[]? buffer = _buffer;
            if (buffer != null && IsBuffered(offset, 1)) return buffer[offset];
            if (_parent != null) return _parent.ReadByte(checked(offset + _parentOffset));
            EnsurePage(offset);
            return _page![offset - checked((int)_pageOffset)];
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal ushort ReadUInt16(int offset) {
            if (offset < 0 || offset > Length - sizeof(ushort)) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            byte[]? buffer = _buffer;
            return buffer != null && IsBuffered(offset, sizeof(ushort))
                ? (ushort)(buffer[offset] | buffer[offset + 1] << 8)
                : ReadUInt16Paged(offset);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal uint ReadUInt32(int offset) {
            if (offset < 0 || offset > Length - sizeof(uint)) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            byte[]? buffer = _buffer;
            return buffer != null && IsBuffered(offset, sizeof(uint))
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
            if (buffer != null && IsBuffered(offset, sizeof(long))) {
                return BitConverter.Int64BitsToDouble(
                    BinaryPrimitives.ReadInt64LittleEndian(buffer.AsSpan(offset, sizeof(long))));
            }
            ulong bits = ReadUInt32(offset) | (ulong)ReadUInt32(checked(offset + 4)) << 32;
            return BitConverter.Int64BitsToDouble(unchecked((long)bits));
        }

        /// <summary>
        /// Gets a record slice for decoding. A buffered source returns a view of its pooled
        /// buffer, which is valid until the source is disposed; a paged source returns a copy.
        /// </summary>
        internal ArraySegment<byte> GetSegment(int offset, int count) {
            ThrowIfDisposed();
            if (offset < 0 || count < 0 || offset > Length - count) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            byte[]? buffer = _buffer;
            if (buffer != null && IsBuffered(offset, count)) return new ArraySegment<byte>(buffer, offset, count);
            return new ArraySegment<byte>(Copy(offset, count));
        }

        internal byte[] Copy(int offset, int count) {
            ThrowIfDisposed();
            if (offset < 0 || count < 0 || offset > Length - count) {
                throw new InvalidDataException("Unexpected end of BIFF record.");
            }
            var result = new byte[count];
            if (_buffer != null && IsBuffered(offset, count)) {
                Buffer.BlockCopy(_buffer, offset, result, 0, count);
                return result;
            }
            if (_parent != null) return _parent.Copy(checked(offset + _parentOffset), count);
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
                Array.Clear(buffer, 0, _bufferedLength);
                ArrayPool<byte>.Shared.Return(buffer);
            }
            _page = null;
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(LegacyBiffSource));
        }
    }
}
