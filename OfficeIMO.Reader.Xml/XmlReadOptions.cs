namespace OfficeIMO.Reader.Xml;

/// <summary>
/// Options for XML adapter behavior.
/// </summary>
public sealed class XmlReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>Maximum element nesting, with the root at depth one.</summary>
    public int MaxDepth { get; set; } = 128;

    /// <summary>Maximum XML nodes and attributes accepted before emitting chunks.</summary>
    public int MaxNodes { get; set; } = 200_000;

    /// <summary>Maximum characters in an attribute or an element's combined direct text.</summary>
    public int MaxScalarLength { get; set; } = 1_048_576;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
