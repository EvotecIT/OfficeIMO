namespace OfficeIMO.Yaml;

/// <summary>
/// A text writer over one pooled <see cref="char"/> array that the emitter appends to without virtual dispatch. For a
/// string it grows and ends as one string, never copying chunk lists as <see cref="StringWriter"/> does. Over a stream
/// it encodes UTF-8 in blocks whenever the buffer fills.
/// </summary>
internal sealed class YamlCharBuffer : TextWriter {
    private const int StreamBlock = 16 * 1024;
    private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

    private readonly Stream? _stream;
    private readonly Encoder? _encoder;
    private byte[]? _bytes;
    private char[] _buffer;
    private int _length;

    public YamlCharBuffer(int capacity = 4096) {
        _buffer = ArrayPool<char>.Shared.Rent(capacity);
    }

    /// <summary>Creates a buffer that writes UTF-8 without a byte-order mark to a caller-owned stream.</summary>
    public YamlCharBuffer(Stream stream) {
        _stream = stream;
        _encoder = Utf8.GetEncoder();
        _buffer = ArrayPool<char>.Shared.Rent(StreamBlock + 1024);
        _bytes = ArrayPool<byte>.Shared.Rent(Utf8.GetMaxByteCount(_buffer.Length));
    }

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
        if (_stream is not null && _length + text.Length > StreamBlock) Drain(final: false);
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

    public override string ToString() => new string(_buffer, 0, _length);

    /// <summary>Encodes everything buffered to the stream; for a string buffer it does nothing.</summary>
    public override void Flush() {
        if (_stream is null) return;
        Drain(final: true);
        _stream.Flush();
    }

    private void Drain(bool final) {
        if (_length == 0 && !final) return;
        byte[] bytes = _bytes!;
        int written;
        if (_buffer.Length > 0 && Utf8.GetMaxByteCount(_length) > bytes.Length) {
            ArrayPool<byte>.Shared.Return(bytes);
            _bytes = bytes = ArrayPool<byte>.Shared.Rent(Utf8.GetMaxByteCount(_length));
        }
        written = _encoder!.GetBytes(_buffer, 0, _length, bytes, 0, flush: final);
        _stream!.Write(bytes, 0, written);
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
