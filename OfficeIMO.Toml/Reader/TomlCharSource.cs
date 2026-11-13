namespace OfficeIMO.Toml;

/// <summary>Supplies decoded characters to the reader's sliding buffer.</summary>
internal abstract class TomlCharSource {
    /// <summary>
    /// Reads up to <paramref name="count"/> characters (at least two are always requested). Returns 0 only at the end of
    /// the input and -1 when the next input bytes are not valid UTF-8; every valid character before them is returned first.
    /// </summary>
    public abstract int Read(char[] buffer, int offset, int count);

    /// <summary>Gets the input consumed so far: bytes for a stream, characters for a text reader.</summary>
    public abstract long Consumed { get; }

    /// <summary>Returns pooled buffers. The caller's reader or stream is never disposed.</summary>
    public virtual void Release() { }
}

/// <summary>Reads characters from a caller-owned <see cref="TextReader"/>.</summary>
internal sealed class TomlTextReaderSource : TomlCharSource {
    private readonly TextReader _reader;
    private long _consumed;

    public TomlTextReaderSource(TextReader reader) => _reader = reader;

    public override long Consumed => _consumed;

    public override int Read(char[] buffer, int offset, int count) {
        int read = _reader.Read(buffer, offset, count);
        if (read > 0) _consumed += read;
        return read;
    }
}

/// <summary>
/// Decodes a caller-owned UTF-8 <see cref="Stream"/> strictly, skipping a leading byte-order mark. The decoder stops
/// exactly before the first invalid or truncated sequence, so the reader can report its position.
/// </summary>
internal sealed class TomlStreamSource : TomlCharSource {
    private const int ByteBufferSize = 16 * 1024;
    private readonly Stream _stream;
    private byte[]? _bytes;
    private int _bytePos;
    private int _byteCount;
    private bool _eof;
    private bool _started;
    private bool _invalid;
    private long _consumed;

    public TomlStreamSource(Stream stream) {
        _stream = stream;
        _bytes = ArrayPool<byte>.Shared.Rent(ByteBufferSize);
    }

    /// <summary>Gets whether the stream started with a UTF-8 byte-order mark.</summary>
    public bool HasByteOrderMark { get; private set; }

    public override long Consumed => _consumed;

    public override int Read(char[] buffer, int offset, int count) {
        byte[] bytes = _bytes ?? throw new ObjectDisposedException(nameof(TomlReader));
        if (!_started) Start(bytes);
        while (true) {
            if (_invalid) return -1;
            int written = Decode(bytes, buffer, offset, count);
            if (written > 0) return written;
            if (_invalid) return -1;
            if (_eof) {
                if (_bytePos < _byteCount) {
                    // The input ends inside a multi-byte sequence.
                    _invalid = true;
                    return -1;
                }
                return 0;
            }
            Refill(bytes);
        }
    }

    private void Start(byte[] bytes) {
        _started = true;
        while (_byteCount < 3 && !_eof) Refill(bytes);
        if (_byteCount >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
            _bytePos = 3;
            HasByteOrderMark = true;
        }
    }

    private void Refill(byte[] bytes) {
        int remaining = _byteCount - _bytePos;
        if (_bytePos > 0) {
            if (remaining > 0) Buffer.BlockCopy(bytes, _bytePos, bytes, 0, remaining);
            _bytePos = 0;
            _byteCount = remaining;
        }
        int read = _stream.Read(bytes, _byteCount, bytes.Length - _byteCount);
        if (read <= 0) {
            _eof = true;
        } else {
            _byteCount += read;
            _consumed += read;
        }
    }

    /// <summary>Decodes whole sequences into the buffer; stops at an incomplete sequence, a full buffer or invalid input.</summary>
    private int Decode(byte[] bytes, char[] buffer, int offset, int count) {
        int i = _bytePos, end = _byteCount, d = offset, dEnd = offset + count;
        while (i < end && d < dEnd) {
            byte lead = bytes[i];
            if (lead < 0x80) {
                buffer[d++] = (char)lead;
                i++;
                continue;
            }
            int need, scalar;
            byte low = 0x80, high = 0xBF;
            if (lead >= 0xC2 && lead <= 0xDF) {
                need = 1;
                scalar = lead & 0x1F;
            } else if (lead >= 0xE0 && lead <= 0xEF) {
                need = 2;
                scalar = lead & 0x0F;
                if (lead == 0xE0) low = 0xA0; // overlong
                else if (lead == 0xED) high = 0x9F; // surrogates
            } else if (lead >= 0xF0 && lead <= 0xF4) {
                need = 3;
                scalar = lead & 0x07;
                if (lead == 0xF0) low = 0x90; // overlong
                else if (lead == 0xF4) high = 0x8F; // beyond U+10FFFF
            } else {
                _invalid = true;
                break;
            }
            if (need == 3 && dEnd - d < 2) break;
            int j = 1;
            for (; j <= need && i + j < end; j++) {
                byte next = bytes[i + j];
                if (j == 1 ? next < low || next > high : (next & 0xC0) != 0x80) {
                    _invalid = true;
                    break;
                }
                scalar = (scalar << 6) | (next & 0x3F);
            }
            if (_invalid || j <= need) break;
            if (scalar < 0x10000) {
                buffer[d++] = (char)scalar;
            } else {
                scalar -= 0x10000;
                buffer[d++] = (char)(0xD800 + (scalar >> 10));
                buffer[d++] = (char)(0xDC00 + (scalar & 0x3FF));
            }
            i += need + 1;
        }
        _bytePos = i;
        return d - offset;
    }

    public override void Release() {
        byte[]? bytes = _bytes;
        _bytes = null;
        if (bytes is not null) ArrayPool<byte>.Shared.Return(bytes);
    }
}
