namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Extracts Markdown tables using the configured reader parsing contract.</summary>
    public IReadOnlyList<ReaderTable> ExtractMarkdownTables(
        string markdown,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractMarkdownTables(markdown, options, cancellationToken);
    }

}
