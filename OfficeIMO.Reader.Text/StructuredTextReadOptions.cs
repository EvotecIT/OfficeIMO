namespace OfficeIMO.Reader.Text;

/// <summary>
/// Options for structured-text adapter behavior.
/// </summary>
public sealed class StructuredTextReadOptions {
    /// <summary>
    /// CSV: rows per emitted chunk.
    /// </summary>
    public int CsvChunkRows { get; set; } = 200;

    /// <summary>
    /// CSV: treat first row as headers.
    /// </summary>
    public bool CsvHeadersInFirstRow { get; set; } = true;

    /// <summary>
    /// CSV: include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeCsvMarkdown { get; set; } = true;
}
