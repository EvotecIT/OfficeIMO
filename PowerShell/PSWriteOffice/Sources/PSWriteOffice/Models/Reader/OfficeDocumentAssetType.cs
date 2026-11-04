namespace PSWriteOffice.Models.Reader;

/// <summary>Document content to materialize outside the source document.</summary>
public enum OfficeDocumentAssetType {
    /// <summary>Images and embedded files.</summary>
    Media,
    /// <summary>Structured tables.</summary>
    Table,
    /// <summary>Source visuals and their metadata.</summary>
    Visual
}

/// <summary>Representation of an exported document table.</summary>
public enum OfficeTableExportFormat {
    /// <summary>Comma-separated values.</summary>
    Csv,
    /// <summary>Markdown table.</summary>
    Markdown,
    /// <summary>JSON sidecar.</summary>
    Json
}
