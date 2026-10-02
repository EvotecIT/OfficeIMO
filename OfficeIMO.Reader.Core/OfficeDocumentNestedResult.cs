namespace OfficeIMO.Reader;

/// <summary>A nested document and the path that locates it inside its containing source.</summary>
public sealed class OfficeDocumentNestedResult {
    /// <summary>Container-relative or virtual path. IDs inside Document are local to that document.</summary>
    public string Path { get; set; } = string.Empty;
    /// <summary>The complete nested result, including links, forms, assets, metadata and further nested documents.</summary>
    public OfficeDocumentReadResult Document { get; set; } = new OfficeDocumentReadResult();
}
