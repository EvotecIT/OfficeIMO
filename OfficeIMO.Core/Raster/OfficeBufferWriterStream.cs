#if NET8_0_OR_GREATER
using System;
using System.Buffers;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Adapts a caller-owned buffer writer to the stream-based codec core.</summary>
internal sealed class OfficeBufferWriterStream : Stream {
    private const int MaximumWriteSizeHint = 64 * 1024;
    private readonly IBufferWriter<byte> _writer;
    private long _position;

    internal OfficeBufferWriterStream(IBufferWriter<byte> writer) {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _position;

    public override long Position {
        get => _position;
        set => throw new NotSupportedException();
    }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) =>
        throw new NotSupportedException();

    public override void SetLength(long value) =>
        throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        if (offset > buffer.Length - count) throw new ArgumentException("The buffer range is invalid.", nameof(buffer));

        while (count > 0) {
            Span<byte> destination = _writer.GetSpan(Math.Min(count, MaximumWriteSizeHint));
            if (destination.Length == 0) {
                throw new InvalidOperationException("The buffer writer returned an empty span.");
            }
            int written = Math.Min(count, destination.Length);
            new ReadOnlySpan<byte>(buffer, offset, written).CopyTo(destination);
            _writer.Advance(written);
            offset += written;
            count -= written;
            _position = checked(_position + written);
        }
    }

    public override void WriteByte(byte value) {
        Span<byte> destination = _writer.GetSpan(1);
        if (destination.Length == 0) {
            throw new InvalidOperationException("The buffer writer returned an empty span.");
        }
        destination[0] = value;
        _writer.Advance(1);
        _position = checked(_position + 1);
    }
}
#endif
