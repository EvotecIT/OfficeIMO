namespace OfficeIMO.Reader;

/// <summary>A source-ordered content observation containing exactly one block, table, or fallback text chunk.</summary>
public sealed class OfficeDocumentContentItem {
    internal OfficeDocumentContentItem(OfficeDocumentBlock? block, ReaderTable? table, ReaderChunk? chunk, ReaderLocation? location) {
        Block = block; Table = table; Chunk = chunk; Location = location;
    }

    /// <summary>Normalized source block, or null for a table or fallback chunk.</summary>
    public OfficeDocumentBlock? Block { get; }
    /// <summary>Structured table, including its rows, or null for another observation kind.</summary>
    public ReaderTable? Table { get; }
    /// <summary>Fallback text chunk used when no normalized source blocks contain text.</summary>
    public ReaderChunk? Chunk { get; }
    /// <summary>Effective source location, including page, chunk, or stable block-anchor fallback.</summary>
    public ReaderLocation? Location { get; }
}
