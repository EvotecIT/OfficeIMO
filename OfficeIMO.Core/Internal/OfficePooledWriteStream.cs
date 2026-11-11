#if NET8_0_OR_GREATER
using System.Buffers;

namespace OfficeIMO.Core.Internal;

/// <summary>Collects sequential output in pooled chunks and returns a detached byte array.</summary>
/// <remarks>The caller retains responsibility for output quotas. No pooled storage escapes the stream.</remarks>
internal sealed class OfficePooledWriteStream : Stream {
    private const int ChunkSize = 64 * 1024;
    private readonly List<byte[]> _chunks = new List<byte[]>();
    private int _length;
    private bool _disposed;

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => !_disposed;
    public override long Length { get { EnsureOpen(); return _length; } }
    public override long Position { get => Length; set => throw new NotSupportedException(); }

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureOpen();
        ArgumentNullException.ThrowIfNull(buffer);
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
        if (count > int.MaxValue - _length) throw new IOException("Stream was too long.");
        while (count > 0) {
            int chunkOffset = _length % ChunkSize;
            if (chunkOffset == 0) AddChunk();
            int take = Math.Min(count, ChunkSize - chunkOffset);
            buffer.AsSpan(offset, take).CopyTo(_chunks[_chunks.Count - 1].AsSpan(chunkOffset, take));
            _length += take;
            offset += take;
            count -= take;
        }
    }

    public override void WriteByte(byte value) {
        EnsureOpen();
        if (_length == int.MaxValue) throw new IOException("Stream was too long.");
        int offset = _length % ChunkSize;
        if (offset == 0) AddChunk();
        _chunks[_chunks.Count - 1][offset] = value;
        _length++;
    }

    /// <summary>Copies only produced bytes into storage owned by the caller.</summary>
    internal byte[] ToArray() {
        EnsureOpen();
        if (_length == 0) return Array.Empty<byte>();
        byte[] result = GC.AllocateUninitializedArray<byte>(_length);
        int offset = 0;
        foreach (byte[] chunk in _chunks) {
            int take = Math.Min(ChunkSize, _length - offset);
            chunk.AsSpan(0, take).CopyTo(result.AsSpan(offset, take));
            offset += take;
        }
        return result;
    }

    public override void Flush() => EnsureOpen();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (disposing && !_disposed) {
            _disposed = true;
            foreach (byte[] chunk in _chunks) ArrayPool<byte>.Shared.Return(chunk);
            _chunks.Clear();
            _length = 0;
        }
        base.Dispose(disposing);
    }

    private void AddChunk() {
        byte[] chunk = ArrayPool<byte>.Shared.Rent(ChunkSize);
        try {
            _chunks.Add(chunk);
        } catch {
            ArrayPool<byte>.Shared.Return(chunk);
            throw;
        }
    }

    private void EnsureOpen() => ObjectDisposedException.ThrowIf(_disposed, this);
}
#endif
