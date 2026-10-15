namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Finite native text projection shared by the independent ODF stories.</summary>
internal sealed class OpenDocumentRichBlocks {
    private readonly ReaderOpenDocumentOptions _options;
    private readonly CancellationToken _token;
    private readonly List<OfficeDocumentBlock> _blocks = new();
    private long _characters;
    internal OpenDocumentRichBlocks(ReaderOpenDocumentOptions options, CancellationToken token) {
        if (options.MaxRichBlocks < 0 || options.MaxExtractedCharacters < 1) throw new ArgumentOutOfRangeException(nameof(options));
        _options = options; _token = token;
    }
    internal IReadOnlyList<OfficeDocumentBlock> Blocks => _blocks;
    internal long RemainingCharacters => _options.MaxExtractedCharacters - _characters;
    internal void Retain(IEnumerable<OfficeDocumentBlock> blocks) {
        foreach (OfficeDocumentBlock block in blocks) { Count(block.Text); _blocks.Add(block); }
    }
    internal void Add(string text, string kind, ReaderLocation location, OfficeDocumentRegion? region = null) {
        Count(text);
        _blocks.Add(new OfficeDocumentBlock { Id = "odf-native-block-" + _blocks.Count.ToString("D6", CultureInfo.InvariantCulture),
            Kind = kind, Text = text, Location = location, Region = region });
    }
    private void Count(string text) {
        _token.ThrowIfCancellationRequested();
        if (_blocks.Count >= _options.MaxRichBlocks) throw new ReaderResourceLimitException("OpenDocument.MaxRichBlocks", _options.MaxRichBlocks);
        if (text.Length > _options.MaxExtractedCharacters - _characters) throw new ReaderResourceLimitException("OpenDocument.MaxExtractedCharacters", _options.MaxExtractedCharacters);
        _characters += text.Length;
    }
}
