namespace OfficeIMO.Markdown;

/// <summary>
/// Source span for markdown syntax nodes.
/// </summary>
public readonly struct MarkdownSourceSpan : IEquatable<MarkdownSourceSpan> {
    private const int Missing = int.MinValue;
    private readonly int _startLine;
    private readonly int _startColumn;
    private readonly int _endLine;
    private readonly int _endColumn;
    private readonly int _startOffset;
    private readonly int _endOffset;

    /// <summary>1-based start line.</summary>
    public int StartLine => _startLine;
    /// <summary>1-based start column.</summary>
    public int? StartColumn => _startLine == 0 || _startColumn == Missing ? null : _startColumn;
    /// <summary>1-based end line.</summary>
    public int EndLine => _endLine;
    /// <summary>1-based end column.</summary>
    public int? EndColumn => _startLine == 0 || _endColumn == Missing ? null : _endColumn;
    /// <summary>0-based start offset in the normalized markdown text.</summary>
    public int? StartOffset => _startLine == 0 || _startOffset == Missing ? null : _startOffset;
    /// <summary>0-based end offset in the normalized markdown text.</summary>
    public int? EndOffset => _startLine == 0 || _endOffset == Missing ? null : _endOffset;

    /// <summary>Create a line-based source span.</summary>
    public MarkdownSourceSpan(int startLine, int endLine) {
        if (startLine < 1) {
            startLine = 1;
        }
        if (endLine < startLine) {
            endLine = startLine;
        }

        _startLine = startLine;
        _startColumn = Missing;
        _endLine = endLine;
        _endColumn = Missing;
        _startOffset = Missing;
        _endOffset = Missing;
    }

    /// <summary>Create a source span with line, column, and optional normalized-text offsets.</summary>
    public MarkdownSourceSpan(int startLine, int startColumn, int endLine, int endColumn, int? startOffset = null, int? endOffset = null) {
        if (startLine < 1) {
            startLine = 1;
        }
        if (endLine < startLine) {
            endLine = startLine;
        }
        if (startColumn < 1) {
            startColumn = 1;
        }
        if (endColumn < 1) {
            endColumn = 1;
        }
        if (endLine == startLine && endColumn < startColumn) {
            endColumn = startColumn;
        }

        _startLine = startLine;
        _startColumn = startColumn;
        _endLine = endLine;
        _endColumn = endColumn;
        _startOffset = startOffset ?? Missing;
        _endOffset = endOffset ?? Missing;
    }

    /// <summary>Returns true when the span contains the given 1-based line number.</summary>
    public bool ContainsLine(int lineNumber) {
        if (lineNumber < 1) return false;
        return lineNumber >= StartLine && lineNumber <= EndLine;
    }

    /// <summary>Returns true when this span fully contains the given span.</summary>
    public bool Contains(MarkdownSourceSpan other) {
        if (StartOffset.HasValue && EndOffset.HasValue && other.StartOffset.HasValue && other.EndOffset.HasValue) {
            return StartOffset.Value <= other.StartOffset.Value && EndOffset.Value >= other.EndOffset.Value;
        }

        if (!StartColumn.HasValue || !EndColumn.HasValue || !other.StartColumn.HasValue || !other.EndColumn.HasValue) {
            return other.StartLine >= StartLine && other.EndLine <= EndLine;
        }

        return CompareStartTo(other) <= 0 && CompareEndTo(other) >= 0;
    }

    /// <summary>Returns true when this span overlaps the given span.</summary>
    public bool Overlaps(MarkdownSourceSpan other) {
        if (StartOffset.HasValue && EndOffset.HasValue && other.StartOffset.HasValue && other.EndOffset.HasValue) {
            return StartOffset.Value <= other.EndOffset.Value && EndOffset.Value >= other.StartOffset.Value;
        }

        if (!StartColumn.HasValue || !EndColumn.HasValue || !other.StartColumn.HasValue || !other.EndColumn.HasValue) {
            return other.EndLine >= StartLine && other.StartLine <= EndLine;
        }

        return CompareStartTo(other, useOtherEnd: true) <= 0 && CompareEndTo(other, useOtherStart: true) >= 0;
    }

    /// <summary>Returns true when the span contains the given 1-based line and column.</summary>
    public bool ContainsPosition(int lineNumber, int columnNumber) {
        if (!ContainsLine(lineNumber)) {
            return false;
        }

        if (!StartColumn.HasValue || !EndColumn.HasValue) {
            return true;
        }

        if (lineNumber == StartLine && columnNumber < StartColumn.Value) {
            return false;
        }

        if (lineNumber == EndLine && columnNumber > EndColumn.Value) {
            return false;
        }

        return columnNumber >= 1;
    }

    /// <inheritdoc />
    public override string ToString() {
        if (StartColumn.HasValue && EndColumn.HasValue) {
            return StartLine == EndLine
                ? $"L{StartLine}:C{StartColumn}-C{EndColumn}"
                : $"L{StartLine}:C{StartColumn}-L{EndLine}:C{EndColumn}";
        }

        return StartLine == EndLine
            ? $"L{StartLine}"
            : $"L{StartLine}-L{EndLine}";
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is MarkdownSourceSpan other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = StartLine;
            hash = (hash * 397) ^ (StartColumn ?? 0);
            hash = (hash * 397) ^ EndLine;
            hash = (hash * 397) ^ (EndColumn ?? 0);
            return hash;
        }
    }

    /// <inheritdoc />
    public bool Equals(MarkdownSourceSpan other) =>
        StartLine == other.StartLine
        && StartColumn == other.StartColumn
        && EndLine == other.EndLine
        && EndColumn == other.EndColumn;

    private int CompareStartTo(MarkdownSourceSpan other, bool useOtherEnd = false) {
        int otherLine = useOtherEnd ? other.EndLine : other.StartLine;
        int? otherColumn = useOtherEnd ? other.EndColumn : other.StartColumn;
        return ComparePosition(StartLine, StartColumn, otherLine, otherColumn, treatMissingAsMax: false);
    }

    private int CompareEndTo(MarkdownSourceSpan other, bool useOtherStart = false) {
        int otherLine = useOtherStart ? other.StartLine : other.EndLine;
        int? otherColumn = useOtherStart ? other.StartColumn : other.EndColumn;
        return ComparePosition(EndLine, EndColumn, otherLine, otherColumn, treatMissingAsMax: true);
    }

    private static int ComparePosition(int leftLine, int? leftColumn, int rightLine, int? rightColumn, bool treatMissingAsMax) {
        int lineCompare = leftLine.CompareTo(rightLine);
        if (lineCompare != 0) {
            return lineCompare;
        }

        int normalizedLeft = NormalizeColumn(leftColumn, treatMissingAsMax);
        int normalizedRight = NormalizeColumn(rightColumn, treatMissingAsMax);
        return normalizedLeft.CompareTo(normalizedRight);
    }

    private static int NormalizeColumn(int? column, bool treatMissingAsMax) =>
        column ?? (treatMissingAsMax ? int.MaxValue : 1);
}
