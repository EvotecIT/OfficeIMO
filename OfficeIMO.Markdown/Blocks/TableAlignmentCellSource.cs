namespace OfficeIMO.Markdown;

internal readonly struct TableAlignmentCellSource {
    internal TableAlignmentCellSource(string markdown, MarkdownSourceSpan sourceSpan) {
        Markdown = markdown ?? string.Empty;
        SourceSpan = sourceSpan;
    }

    internal string Markdown { get; }

    internal MarkdownSourceSpan SourceSpan { get; }
}
