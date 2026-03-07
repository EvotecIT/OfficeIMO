namespace OfficeIMO.Markdown;

/// <summary>
/// Line-based source span for markdown syntax nodes.
/// </summary>
public readonly struct MarkdownSourceSpan {
    /// <summary>1-based start line.</summary>
    public int StartLine { get; }
    /// <summary>1-based end line.</summary>
    public int EndLine { get; }

    /// <summary>Create a line-based source span.</summary>
    public MarkdownSourceSpan(int startLine, int endLine) {
        if (startLine < 1) startLine = 1;
        if (endLine < startLine) endLine = startLine;
        StartLine = startLine;
        EndLine = endLine;
    }

    /// <inheritdoc />
    public override string ToString() => StartLine == EndLine
        ? $"L{StartLine}"
        : $"L{StartLine}-L{EndLine}";
}
