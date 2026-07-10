namespace OfficeIMO.AsciiDoc;

internal sealed class AsciiDocSourceLine {
    private readonly string _source;

    internal AsciiDocSourceLine(string source, int lineNumber, int start, int contentEnd, int end) {
        _source = source;
        LineNumber = lineNumber;
        Start = start;
        ContentEnd = contentEnd;
        End = end;
    }

    internal int LineNumber { get; }
    internal int Start { get; }
    internal int ContentEnd { get; }
    internal int End { get; }
    internal int ContentLength => ContentEnd - Start;
    internal int LineEndingLength => End - ContentEnd;
    internal string Content => _source.Substring(Start, ContentLength);
    internal string LineEnding => _source.Substring(ContentEnd, LineEndingLength);
    internal string FullText => _source.Substring(Start, End - Start);
}
