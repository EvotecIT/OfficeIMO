namespace OfficeIMO.Bibliography;

internal sealed class EndNoteCancellableTextReader : TextReader {
    private const int MaximumReadSize = 4096;
    private readonly string _value;
    private readonly CancellationToken _cancellationToken;
    private int _position;

    internal EndNoteCancellableTextReader(string value, CancellationToken cancellationToken, int startPosition = 0) {
        _value = value ?? throw new ArgumentNullException(nameof(value));
        if (startPosition < 0 || startPosition > value.Length) throw new ArgumentOutOfRangeException(nameof(startPosition));
        _position = startPosition;
        _cancellationToken = cancellationToken;
    }

    public override int Peek() {
        _cancellationToken.ThrowIfCancellationRequested();
        return _position < _value.Length ? _value[_position] : -1;
    }

    public override int Read() {
        _cancellationToken.ThrowIfCancellationRequested();
        return _position < _value.Length ? _value[_position++] : -1;
    }

    public override int Read(char[] buffer, int index, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (index < 0 || index > buffer.Length) throw new ArgumentOutOfRangeException(nameof(index));
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        if (buffer.Length - index < count) throw new ArgumentException("The buffer range is invalid.");
        _cancellationToken.ThrowIfCancellationRequested();
        if (_position >= _value.Length) return 0;
        int read = Math.Min(Math.Min(count, MaximumReadSize), _value.Length - _position);
        _value.CopyTo(_position, buffer, index, read);
        _position += read;
        return read;
    }
}
