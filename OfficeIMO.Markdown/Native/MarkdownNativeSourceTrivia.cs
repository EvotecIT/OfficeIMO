namespace OfficeIMO.Markdown;

/// <summary>
/// Identifies source trivia projected by <see cref="MarkdownNativeDocument"/>.
/// </summary>
public enum MarkdownNativeSourceTriviaKind {
    /// <summary>A blank source line, including whitespace-only lines.</summary>
    BlankLine
}

/// <summary>
/// Source-backed document trivia that is not owned by a semantic block or inline.
/// </summary>
public sealed class MarkdownNativeSourceTrivia {
    internal MarkdownNativeSourceTrivia(MarkdownNativeSourceTriviaKind kind, string text, MarkdownSourceSpan sourceSpan) {
        Kind = kind;
        Text = text ?? string.Empty;
        SourceSpan = sourceSpan;
    }

    /// <summary>Trivia kind.</summary>
    public MarkdownNativeSourceTriviaKind Kind { get; }

    /// <summary>Exact normalized line content represented by this trivia, excluding the line ending.</summary>
    public string Text { get; }

    /// <summary>Source span for the trivia content.</summary>
    public MarkdownSourceSpan SourceSpan { get; }

    /// <summary>1-based source line number.</summary>
    public int LineNumber => SourceSpan.StartLine;
}
