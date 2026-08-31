namespace OfficeIMO.Bibliography;

internal sealed class EndNoteBoundedStringWriter : TextWriter {
    private const int ChunkSize = 4096;
    private readonly StringBuilder _builder;
    private readonly IList<BibliographyItem> _partial;
    private readonly BibliographyLimitGuard _limits;
    private readonly int _offset;
    private readonly CancellationToken _cancellationToken;

    internal EndNoteBoundedStringWriter(StringBuilder builder, IList<BibliographyItem> partial, BibliographyLimitGuard limits, int offset, CancellationToken cancellationToken) {
        _builder = builder;
        _partial = partial;
        _limits = limits;
        _offset = offset;
        _cancellationToken = cancellationToken;
    }

    public override Encoding Encoding => Encoding.Unicode;

    public override void Write(char value) {
        _cancellationToken.ThrowIfCancellationRequested();
        _limits.CheckAdditionalValueLength(_partial, _builder.Length, 1, _offset);
        _builder.Append(value);
    }

    public override void Write(char[] buffer, int index, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        while (count > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(ChunkSize, count);
            _limits.CheckAdditionalValueLength(_partial, _builder.Length, length, _offset);
            _builder.Append(buffer, index, length);
            index += length;
            count -= length;
        }
    }

    public override void Write(string? value) {
        if (value == null) return;
        for (int index = 0; index < value.Length;) {
            _cancellationToken.ThrowIfCancellationRequested();
            int length = Math.Min(ChunkSize, value.Length - index);
            _limits.CheckAdditionalValueLength(_partial, _builder.Length, length, _offset);
            _builder.Append(value, index, length);
            index += length;
        }
    }
}
