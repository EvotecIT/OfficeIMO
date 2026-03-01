namespace OfficeIMO.Reader.Xml;

/// <summary>
/// Options for XML adapter behavior.
/// </summary>
public sealed class XmlReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
