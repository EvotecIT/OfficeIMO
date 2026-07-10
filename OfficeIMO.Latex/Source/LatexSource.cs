namespace OfficeIMO.Latex;

/// <summary>One-based position in decoded LaTeX source.</summary>
public readonly struct LatexSourcePosition : IEquatable<LatexSourcePosition> {
    internal LatexSourcePosition(int offset, int line, int column) {
        Offset = offset;
        Line = line;
        Column = column;
    }

    /// <summary>Zero-based UTF-16 offset.</summary>
    public int Offset { get; }
    /// <summary>One-based line.</summary>
    public int Line { get; }
    /// <summary>One-based column.</summary>
    public int Column { get; }

    /// <inheritdoc />
    public bool Equals(LatexSourcePosition other) => Offset == other.Offset && Line == other.Line && Column == other.Column;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is LatexSourcePosition other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => Offset;
    /// <inheritdoc />
    public override string ToString() => Line + ":" + Column;
}

/// <summary>Half-open span in decoded LaTeX source.</summary>
public readonly struct LatexSourceSpan : IEquatable<LatexSourceSpan> {
    internal LatexSourceSpan(LatexSourcePosition start, LatexSourcePosition end) {
        Start = start;
        End = end;
    }

    /// <summary>Start position.</summary>
    public LatexSourcePosition Start { get; }
    /// <summary>Exclusive end position.</summary>
    public LatexSourcePosition End { get; }
    /// <summary>Character length.</summary>
    public int Length => End.Offset - Start.Offset;
    /// <summary>Tests whether an offset is inside the span.</summary>
    public bool ContainsOffset(int offset) => offset >= Start.Offset && offset < End.Offset;
    /// <summary>Returns the represented source slice.</summary>
    public string Slice(string source) => source.Substring(Start.Offset, Length);
    /// <inheritdoc />
    public bool Equals(LatexSourceSpan other) => Start.Equals(other.Start) && End.Equals(other.End);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is LatexSourceSpan other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => Start.GetHashCode() ^ End.GetHashCode();
    /// <inheritdoc />
    public override string ToString() => Start + "-" + End;
}

/// <summary>Complete decoded LaTeX source with offset mapping.</summary>
public sealed class LatexSourceText {
    private readonly int[] _lineStarts;

    internal LatexSourceText(string text) {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        var starts = new List<int> { 0 };
        for (int index = 0; index < text.Length; index++) {
            if (text[index] == '\r') {
                if (index + 1 < text.Length && text[index + 1] == '\n') index++;
                starts.Add(index + 1);
            } else if (text[index] == '\n') {
                starts.Add(index + 1);
            }
        }
        _lineStarts = starts.ToArray();
        PreferredLineEnding = DetectLineEnding(text);
    }

    /// <summary>Complete source characters.</summary>
    public string Text { get; }
    /// <summary>Number of logical lines.</summary>
    public int LineCount => _lineStarts.Length;
    /// <summary>First source line ending, or environment newline when absent.</summary>
    public string PreferredLineEnding { get; }

    /// <summary>Maps an offset to line and column.</summary>
    public LatexSourcePosition GetPosition(int offset) {
        if (offset < 0 || offset > Text.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        int line = Array.BinarySearch(_lineStarts, offset);
        if (line < 0) line = ~line - 1;
        return new LatexSourcePosition(offset, line + 1, offset - _lineStarts[line] + 1);
    }

    /// <summary>Creates a checked half-open span.</summary>
    public LatexSourceSpan CreateSpan(int start, int end) {
        if (start < 0 || end < start || end > Text.Length) throw new ArgumentOutOfRangeException(nameof(start));
        return new LatexSourceSpan(GetPosition(start), GetPosition(end));
    }

    private static string DetectLineEnding(string text) {
        for (int index = 0; index < text.Length; index++) {
            if (text[index] == '\r') return index + 1 < text.Length && text[index + 1] == '\n' ? "\r\n" : "\r";
            if (text[index] == '\n') return "\n";
        }
        return Environment.NewLine;
    }
}
