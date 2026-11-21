namespace OfficeIMO.Yaml;

/// <summary>
/// A text writer over one pooled <see cref="char"/> array that the emitter appends to without virtual dispatch. For a
/// string it grows and ends as one string, never copying chunk lists as <see cref="StringWriter"/> does. Over a stream
/// it encodes UTF-8 in blocks whenever the buffer fills.
/// </summary>
internal sealed class YamlCharBuffer : TextWriter {
    private const int MinStreamBlock = 16 * 1024;
#if NET8_0_OR_GREATER
    // The shared pool keeps large arrays here, so a document up to this size reaches the stream in one write.
    private const int MaxStreamBlock = 1 << 20;
#else
    private const int MaxStreamBlock = MinStreamBlock;
#endif
    private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

    private Stream? _stream;
    private readonly Encoder? _encoder;
    // Characters buffered before a stream write; a string buffer never drains.
    private int _block = int.MaxValue;
    private byte[]? _bytes;
    private char[] _buffer;
    private int _length;
    private int _flushed;

    public YamlCharBuffer(int capacity = 4096) {
        _buffer = ArrayPool<char>.Shared.Rent(Math.Max(capacity, 4096));
    }

    // A reused buffer keeps arrays up to these sizes between calls instead of returning them to the pool.
    private const int KeptChars = 16 * 1024;
    private const int KeptBytes = 48 * 1024;

    /// <summary>Starts a string buffer again after <see cref="ReturnArray"/>, as a new one of this capacity would.</summary>
    public void Reset(int capacity) {
        capacity = Math.Max(capacity, 4096);
        if (_buffer.Length < capacity) {
            if (_buffer.Length > 0) ArrayPool<char>.Shared.Return(_buffer);
            _buffer = ArrayPool<char>.Shared.Rent(capacity);
        }
        _length = 0;
    }

    /// <summary>
    /// Lets go of a stream and returns large arrays to the pool, keeping the object, and small arrays, for a reset.
    /// </summary>
    public void ReturnArray() {
        char[] buffer = _buffer;
        _length = 0;
        if (buffer.Length > KeptChars) {
            _buffer = Array.Empty<char>();
            ArrayPool<char>.Shared.Return(buffer);
        }
        if (_bytes is not null && _bytes.Length > KeptBytes) {
            ArrayPool<byte>.Shared.Return(_bytes);
            _bytes = null;
        }
        if (_stream is not null) _stream = Stream.Null;
    }

    /// <summary>Gets the number of characters written, including those already encoded to a stream.</summary>
    public int Length => _flushed + _length;

    /// <summary>
    /// Creates a buffer that writes UTF-8 without a byte-order mark to a caller-owned stream, in blocks sized for a
    /// document of about <paramref name="lengthHint"/> characters: a document of the size seen before reaches the
    /// stream in one write, so a growing stream such as <see cref="MemoryStream"/> sizes itself once.
    /// </summary>
    public YamlCharBuffer(Stream stream, int lengthHint) {
        _stream = stream;
        _encoder = Utf8.GetEncoder();
        _block = BlockFor(lengthHint);
        _buffer = ArrayPool<char>.Shared.Rent(_block + 1024);
    }

    private static int BlockFor(int lengthHint) => Math.Min(Math.Max(lengthHint + lengthHint / 8, MinStreamBlock), MaxStreamBlock);

    /// <summary>Starts a stream buffer again after <see cref="ReturnArray"/>, as a new one over <paramref name="stream"/> would.</summary>
    public void Reset(Stream stream, int lengthHint) {
        _stream = stream;
        _encoder!.Reset();
        _block = BlockFor(lengthHint);
        if (_buffer.Length < _block + 1024) {
            if (_buffer.Length > 0) ArrayPool<char>.Shared.Return(_buffer);
            _buffer = ArrayPool<char>.Shared.Rent(_block + 1024);
        }
        _length = 0;
        _flushed = 0;
    }

    /// <summary>Whether the buffer writes to a stream.</summary>
    public bool IsStream => _stream is not null;

    public override Encoding Encoding => Encoding.Unicode;

    public override void Write(char value) {
        if (_length == _buffer.Length) {
            if (_stream is not null) Drain(final: false);
            else Grow(1);
        }
        _buffer[_length++] = value;
    }

    public override void Write(string? value) {
        if (value is null) return;
        Append(value.AsSpan());
    }

    public override void Write(char[] buffer, int index, int count) => Append(buffer.AsSpan(index, count));

#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<char> buffer) => Append(buffer);
#endif

    public void Append(ReadOnlySpan<char> text) {
        if (_length + text.Length > _block) Drain(final: false);
        if (_buffer.Length - _length < text.Length) Grow(text.Length);
        // Most runs are a few characters (indentation, indicators, short scalars), where a loop beats a memmove call.
        if (text.Length <= 16) {
            char[] buffer = _buffer;
            int position = _length;
            for (int i = 0; i < text.Length; i++) buffer[position + i] = text[i];
        } else {
            text.CopyTo(_buffer.AsSpan(_length));
        }
        _length += text.Length;
    }

    /// <summary>
    /// Appends a block mapping entry on a new line, the line start (line break and indentation), the key, <c>": "</c>
    /// and the value, with one capacity check instead of one per piece.
    /// </summary>
    public void AppendEntry(ReadOnlySpan<char> lineStart, string key, ReadOnlySpan<char> value) {
        int total = lineStart.Length + key.Length + 2 + value.Length;
        if (_length + total > _block) Drain(final: false);
        if (_buffer.Length - _length < total) Grow(total);
        char[] buffer = _buffer;
        int position = _length;
        Copy(lineStart, buffer, ref position);
        Copy(key.AsSpan(), buffer, ref position);
        buffer[position++] = ':';
        buffer[position++] = ' ';
        Copy(value, buffer, ref position);
        _length = position;
    }

    /// <summary>Appends <paramref name="count"/> spaces, the indentation of a deep line, in one step.</summary>
    public void AppendSpaces(int count) {
        if (_length + count > _block) Drain(final: false);
        if (_buffer.Length - _length < count) Grow(count);
        _buffer.AsSpan(_length, count).Fill(' ');
        _length += count;
    }

    /// <summary>Appends a block mapping key on a new line: the line start (or padding), the key and <c>':'</c>.</summary>
    public void AppendKey(ReadOnlySpan<char> lineStart, string key) {
        int total = lineStart.Length + key.Length + 1;
        if (_length + total > _block) Drain(final: false);
        if (_buffer.Length - _length < total) Grow(total);
        char[] buffer = _buffer;
        int position = _length;
        Copy(lineStart, buffer, ref position);
        Copy(key.AsSpan(), buffer, ref position);
        buffer[position++] = ':';
        _length = position;
    }

    /// <summary>
    /// Appends a block sequence item on a new line, the line start, <c>-</c>, the padding after it and the value, with
    /// one capacity check instead of one per piece.
    /// </summary>
    public void AppendItem(ReadOnlySpan<char> lineStart, ReadOnlySpan<char> padding, ReadOnlySpan<char> value) {
        int total = lineStart.Length + 1 + padding.Length + value.Length;
        if (_length + total > _block) Drain(final: false);
        if (_buffer.Length - _length < total) Grow(total);
        char[] buffer = _buffer;
        int position = _length;
        Copy(lineStart, buffer, ref position);
        buffer[position++] = '-';
        Copy(padding, buffer, ref position);
        Copy(value, buffer, ref position);
        _length = position;
    }

    private static void Copy(ReadOnlySpan<char> text, char[] buffer, ref int position) {
        if (text.Length <= 16) {
            for (int i = 0; i < text.Length; i++) buffer[position + i] = text[i];
        } else {
            text.CopyTo(buffer.AsSpan(position));
        }
        position += text.Length;
    }

    public override string ToString() => new string(_buffer, 0, _length);

    /// <summary>Encodes the characters written to a string buffer as UTF-8 into an array of exactly their length.</summary>
    public byte[] ToUtf8Bytes() {
        if (_length <= SinglePassLimit) {
            // Encoding once into pooled room and copying is cheaper than counting first when the text is not ASCII,
            // and costs a copy of a small array when it is.
            int needed = Utf8.GetMaxByteCount(_length);
            byte[]? room = _bytes;
            if (room is null || room.Length < needed) {
                if (room is not null) ArrayPool<byte>.Shared.Return(room);
                _bytes = room = ArrayPool<byte>.Shared.Rent(needed);
            }
            int count = Utf8.GetBytes(_buffer, 0, _length, room, 0);
            var exact = new byte[count];
            Buffer.BlockCopy(room, 0, exact, 0, count);
            return exact;
        }
        var bytes = new byte[Utf8.GetByteCount(_buffer, 0, _length)];
        Utf8.GetBytes(_buffer, 0, _length, bytes, 0);
        return bytes;
    }

    // Up to this many characters, ToUtf8Bytes encodes once into pooled room of three bytes per character.
    private const int SinglePassLimit = 64 * 1024;

    /// <summary>Encodes everything buffered to the stream; for a string buffer it does nothing.</summary>
    public override void Flush() {
        if (_stream is null) return;
        Drain(final: true);
        _stream.Flush();
    }

    private void Drain(bool final) {
        if (_length == 0 && !final) return;
        byte[]? bytes = _bytes;
        int needed = Utf8.GetMaxByteCount(_length);
        if (bytes is null || needed > bytes.Length) {
            if (bytes is not null) ArrayPool<byte>.Shared.Return(bytes);
            _bytes = bytes = ArrayPool<byte>.Shared.Rent(needed);
        }
        int written = _encoder!.GetBytes(_buffer, 0, _length, bytes, 0, flush: final);
        _stream!.Write(bytes, 0, written);
        _flushed += _length;
        _length = 0;
    }

    private void Grow(int needed) {
        char[] next = ArrayPool<char>.Shared.Rent(Math.Max(_buffer.Length * 2, _length + needed));
        _buffer.AsSpan(0, _length).CopyTo(next);
        ArrayPool<char>.Shared.Return(_buffer);
        _buffer = next;
    }

    protected override void Dispose(bool disposing) {
        if (disposing && _buffer.Length > 0) {
            ArrayPool<char>.Shared.Return(_buffer);
            _buffer = Array.Empty<char>();
            _length = 0;
            if (_bytes is not null) {
                ArrayPool<byte>.Shared.Return(_bytes);
                _bytes = null;
            }
        }
        base.Dispose(disposing);
    }
}
