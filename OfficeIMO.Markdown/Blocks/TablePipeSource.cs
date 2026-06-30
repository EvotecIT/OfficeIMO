namespace OfficeIMO.Markdown;

internal readonly struct TablePipeSource {
    internal TablePipeSource(int rowIndex, int columnIndex, MarkdownSourceSpan sourceSpan) {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        SourceSpan = sourceSpan;
    }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }

    internal MarkdownSourceSpan SourceSpan { get; }
}
