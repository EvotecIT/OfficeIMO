namespace OfficeIMO.AsciiDoc;

/// <summary>Operation-local text and provenance builder; the processor owns output-budget enforcement.</summary>
internal sealed class AsciiDocProcessingOutput {
    private readonly StringBuilder _text;
    private readonly List<AsciiDocSourceMapping> _entries = new List<AsciiDocSourceMapping>();
    private string? _sourceName;
    private AsciiDocSourceSpan _span;
    private bool _exact;
    internal AsciiDocProcessingOutput(int capacity) { _text = new StringBuilder(capacity); }
    internal int Length => _text.Length;
    internal IReadOnlyList<AsciiDocSourceMapping> Entries => _entries;
    internal void SetOrigin(string? sourceName, AsciiDocSourceSpan span, bool exact) { _sourceName = sourceName; _span = span; _exact = exact; }
    internal void Append(string value, bool exact) {
        if (value.Length == 0) return;
        _entries.Add(new AsciiDocSourceMapping(_text.Length, value.Length, _sourceName, _span, exact && _exact && value.Length == _span.Length));
        _text.Append(value);
    }
    internal void Append(AsciiDocProcessedSource source, System.Threading.CancellationToken token) {
        int offset = _text.Length;
        foreach (AsciiDocSourceMapping entry in source.Entries) { token.ThrowIfCancellationRequested(); _entries.Add(entry.Shift(offset)); }
        _text.Append(source.Content);
    }
    internal AsciiDocProcessedSource Build() => new AsciiDocProcessedSource(_text.ToString(), _entries.AsReadOnly());
}

internal sealed class AsciiDocProcessedSource {
    internal AsciiDocProcessedSource(string content, IReadOnlyList<AsciiDocSourceMapping> entries) { Content = content; Entries = entries; }
    internal string Content { get; }
    internal IReadOnlyList<AsciiDocSourceMapping> Entries { get; }
}

internal sealed class AsciiDocSelectedLine {
    private readonly AsciiDocSourceLine _line;
    private string? _replacement;
    internal AsciiDocSelectedLine(AsciiDocSourceLine line) {
        _line = line;
        OriginalSpan = new AsciiDocSourceSpan(new AsciiDocSourcePosition(line.Start, line.LineNumber, 1),
            line.LineEndingLength > 0 ? new AsciiDocSourcePosition(line.End, line.LineNumber + 1, 1) :
            new AsciiDocSourcePosition(line.End, line.LineNumber, line.ContentLength + 1));
        IsExact = true;
    }
    internal string Text => _replacement == null ? _line.FullText : _replacement + LineEnding;
    internal string Content => _replacement ?? _line.Content;
    internal string LineEnding => _line.LineEnding;
    internal AsciiDocSourceSpan OriginalSpan { get; }
    internal bool IsExact { get; private set; }
    internal void ReplaceContent(string content) { _replacement = content; IsExact = false; }
}
