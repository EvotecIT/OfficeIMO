namespace OfficeIMO.Reader.Json;

/// <summary>
/// Options for JSON adapter behavior.
/// </summary>
/// <remarks>
/// The adapter checks the limits in one pass over the input before it builds the document or emits chunks.
/// Exceeding a limit throws <see cref="ReaderResourceLimitException"/> named after the property.
/// </remarks>
public sealed class JsonReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Maximum nesting of objects and arrays, with the root container at depth one.
    /// </summary>
    public int MaxDepth { get; set; } = 64;

    /// <summary>
    /// Maximum JSON values (objects, arrays, strings, numbers, booleans and nulls) in the document.
    /// </summary>
    public int MaxNodes { get; set; } = 1_000_000;

    /// <summary>
    /// Maximum characters in one decoded string value or property name, or in one number as written.
    /// </summary>
    public int MaxScalarLength { get; set; } = 1_048_576;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
