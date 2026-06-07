namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalDocument {
    /// <summary>Catalog XMP metadata stream discovered from /Metadata.</summary>
    public PdfXmpMetadataInfo? XmpMetadata { get; }

    /// <summary>True when readable catalog XMP metadata was discovered.</summary>
    public bool HasReadableXmpMetadata => XmpMetadata is not null;
}
