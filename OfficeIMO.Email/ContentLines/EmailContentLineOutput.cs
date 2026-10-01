namespace OfficeIMO.Email;

/// <summary>Bounds accumulated UTF-8 output in the Outlook portable-content codecs.</summary>
internal sealed class EmailContentLineOutput {
    private readonly StringBuilder _output = new StringBuilder();
    private readonly long _maximum;
    private readonly CancellationToken _cancellationToken;
    private long _bytes;

    internal EmailContentLineOutput(long maximum, CancellationToken cancellationToken) {
        _maximum = maximum; _cancellationToken = cancellationToken;
    }

    internal EmailContentLineOutput Append(StringBuilder value) => Append(value.ToString());

    internal EmailContentLineOutput Append(string value) {
        _cancellationToken.ThrowIfCancellationRequested();
        int count = Encoding.UTF8.GetByteCount(value);
        if (count > _maximum - _bytes)
            throw new EmailLimitExceededException("MaxOutputBytes", _bytes + count, _maximum);
        _output.Append(value); _bytes += count;
        return this;
    }

    public override string ToString() => _output.ToString();
}
