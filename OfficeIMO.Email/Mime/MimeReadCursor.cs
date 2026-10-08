namespace OfficeIMO.Email;

/// <summary>Buffers codec byte scans with a logical cursor so each byte does not query a platform stream position.</summary>
internal sealed class MimeReadCursor {
    private readonly Stream _source;
    private readonly byte[] _buffer = new byte[64 * 1024];
    private long _bufferStart = -1;
    private int _bufferLength;
    private long _position;

    internal MimeReadCursor(Stream source) {
        _source = source;
        _position = source.Position;
        Length = source.Length;
    }

    internal long Length { get; }
    internal long Position {
        get => _position;
        set {
            if (value < 0 || value > Length) throw new IOException("The MIME cursor cannot seek outside the input.");
            _position = value;
        }
    }

    internal int ReadByte() {
        if (_position >= Length) return -1;
        if (!EnsureBuffer()) return -1;
        return _buffer[(int)(_position++ - _bufferStart)];
    }

    internal int Read(byte[] output, int offset, int count) {
        if (output == null) throw new ArgumentNullException(nameof(output));
        if (offset < 0 || count < 0 || offset > output.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
        int requested = (int)Math.Min(count, Length - _position);
        int total = 0;
        while (total < requested && EnsureBuffer()) {
            int available = _bufferLength - (int)(_position - _bufferStart);
            int take = Math.Min(requested - total, available);
            Buffer.BlockCopy(_buffer, (int)(_position - _bufferStart), output, offset + total, take);
            _position += take;
            total += take;
        }
        return total;
    }

    private bool EnsureBuffer() {
        if (_position >= _bufferStart && _position - _bufferStart < _bufferLength) return true;
        _source.Position = _position;
        _bufferStart = _position;
        _bufferLength = _source.Read(_buffer, 0, (int)Math.Min(_buffer.Length, Length - _position));
        return _bufferLength > 0;
    }
}
