namespace OfficeIMO.Latex;

// Parser and projection readers need offsets and values, not a retained public
// object per token. These immutable records keep exact source ownership intact.
internal readonly struct LatexTokenRecord {
    private readonly int _kindAndTermination;

    internal LatexTokenRecord(LatexTokenKind kind, string? value, int start, int end, bool terminated) {
        _kindAndTermination = (int)kind | (terminated ? 0 : 256);
        Value = value;
        StartOffset = start;
        EndOffset = end;
    }

    internal LatexTokenKind Kind => (LatexTokenKind)(_kindAndTermination & 255);
    internal bool IsTerminated => (_kindAndTermination & 256) == 0;
    internal string? Value { get; }
    internal int StartOffset { get; }
    internal int EndOffset { get; }
}

internal readonly struct LatexTokenView {
    private readonly LatexSourceText _source;
    private readonly LatexTokenRecord _record;

    internal LatexTokenView(LatexSourceText source, LatexTokenRecord record) {
        _source = source;
        _record = record;
    }

    internal LatexTokenKind Kind => _record.Kind;
    internal bool IsTerminated => _record.IsTerminated;
    internal string? Value => _record.Value;
    internal int StartOffset => _record.StartOffset;
    internal int EndOffset => _record.EndOffset;
    internal LatexSourceSpan Span => _source.CreateSpan(StartOffset, EndOffset);
    internal string Text => _source.Text.Substring(StartOffset, EndOffset - StartOffset);
}
