namespace OfficeIMO.Bibliography;

internal sealed class EndNoteCancellableTextReader : TextReader {
    private const int MaximumReadSize = 4096;
    private readonly string _value;
    private readonly CancellationToken _cancellationToken;
    private int _position;

    internal EndNoteCancellableTextReader(string value, CancellationToken cancellationToken, int startPosition = 0) {
        _value = value ?? throw new ArgumentNullException(nameof(value));
        ArgumentOutOfRangeException.ThrowIfNegative(startPosition);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(startPosition, value.Length);
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
        ArgumentNullException.ThrowIfNull(buffer);
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(index, buffer.Length);
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        if (buffer.Length - index < count) throw new ArgumentException("The buffer range is invalid.");
        _cancellationToken.ThrowIfCancellationRequested();
        if (_position >= _value.Length) return 0;
        int read = Math.Min(Math.Min(count, MaximumReadSize), _value.Length - _position);
        _value.CopyTo(_position, buffer, index, read);
        _position += read;
        return read;
    }
}
