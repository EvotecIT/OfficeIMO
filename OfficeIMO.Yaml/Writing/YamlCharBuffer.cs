namespace OfficeIMO.Yaml;

/// <summary>
/// A text writer over one pooled, growing <see cref="char"/> array, for output that ends as a single string. Unlike
/// <see cref="StringWriter"/> it never copies chunk lists, and the emitter calls it without virtual dispatch.
/// </summary>
internal sealed class YamlCharBuffer : TextWriter {
    private char[] _buffer;
    private int _length;

    public YamlCharBuffer(int capacity = 4096) {
        _buffer = ArrayPool<char>.Shared.Rent(capacity);
    }

    public override Encoding Encoding => Encoding.Unicode;

    public override void Write(char value) {
        if (_length == _buffer.Length) Grow(1);
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
        }
        base.Dispose(disposing);
    }
}
