namespace OfficeIMO.Reader.Csv;

/// <summary>
/// Options for CSV/TSV adapter behavior.
/// </summary>
public sealed class CsvReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Treat first row as headers.
    /// </summary>
    public bool HeadersInFirstRow { get; set; } = true;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
