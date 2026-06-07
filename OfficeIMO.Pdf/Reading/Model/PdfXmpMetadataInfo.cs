namespace OfficeIMO.Pdf;

/// <summary>
/// Readback metadata for a catalog XMP metadata stream.
/// </summary>
public sealed class PdfXmpMetadataInfo {
    internal PdfXmpMetadataInfo(
        int? objectNumber,
        string? subtype,
        string? filter,
        int streamSizeBytes,
        int decodedSizeBytes,
        IReadOnlyList<string> unsupportedFilters,
        string? rawXml,
        bool isWellFormedXml,
        string? title,
        string? creator,
        string? description,
        IReadOnlyList<string> subjects,
        string? producer,
        string? keywords,
        int? pdfAPart,
        string? pdfAConformance,
        int? pdfUaPart,
        string? electronicInvoiceDocumentType,
        string? electronicInvoiceDocumentFileName,
        string? electronicInvoiceVersion,
        string? electronicInvoiceConformanceLevel) {
        ObjectNumber = objectNumber;
        Subtype = subtype;
        Filter = filter;
        StreamSizeBytes = streamSizeBytes;
        DecodedSizeBytes = decodedSizeBytes;
        UnsupportedFilters = unsupportedFilters;
        RawXml = rawXml;
        IsWellFormedXml = isWellFormedXml;
        Title = title;
        Creator = creator;
        Description = description;
        Subjects = subjects;
        Producer = producer;
        Keywords = keywords;
        PdfAPart = pdfAPart;
        PdfAConformance = pdfAConformance;
        PdfUaPart = pdfUaPart;
        ElectronicInvoiceDocumentType = electronicInvoiceDocumentType;
        ElectronicInvoiceDocumentFileName = electronicInvoiceDocumentFileName;
        ElectronicInvoiceVersion = electronicInvoiceVersion;
        ElectronicInvoiceConformanceLevel = electronicInvoiceConformanceLevel;
    }

    /// <summary>Metadata stream object number when the catalog entry is indirect.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Metadata stream /Subtype name, usually XML.</summary>
    public string? Subtype { get; }

    /// <summary>Metadata stream filter name or simple filter value, when present.</summary>
    public string? Filter { get; }

    /// <summary>Raw metadata stream size in bytes.</summary>
    public int StreamSizeBytes { get; }

    /// <summary>Decoded metadata stream size in bytes after supported filters are applied.</summary>
    public int DecodedSizeBytes { get; }

    /// <summary>Unsupported stream filters discovered on the metadata stream.</summary>
    public IReadOnlyList<string> UnsupportedFilters { get; }

    /// <summary>True when at least one unsupported stream filter was discovered.</summary>
    public bool HasUnsupportedFilters => UnsupportedFilters.Count > 0;

    /// <summary>Decoded XMP XML text when the metadata stream could be decoded as text.</summary>
    public string? RawXml { get; }

    /// <summary>True when the decoded XMP text parsed as XML.</summary>
    public bool IsWellFormedXml { get; }

    /// <summary>Dublin Core title, usually dc:title/x-default.</summary>
    public string? Title { get; }

    /// <summary>First Dublin Core creator value.</summary>
    public string? Creator { get; }

    /// <summary>Dublin Core description, usually dc:description/x-default.</summary>
    public string? Description { get; }

    /// <summary>Dublin Core subject bag values.</summary>
    public IReadOnlyList<string> Subjects { get; }

    /// <summary>PDF producer metadata from pdf:Producer.</summary>
    public string? Producer { get; }

    /// <summary>PDF keyword metadata from pdf:Keywords.</summary>
    public string? Keywords { get; }

    /// <summary>PDF/A identification part from pdfaid:part.</summary>
    public int? PdfAPart { get; }

    /// <summary>PDF/A identification conformance from pdfaid:conformance.</summary>
    public string? PdfAConformance { get; }

    /// <summary>True when PDF/A identification metadata was found.</summary>
    public bool HasPdfAIdentification => PdfAPart.HasValue || !string.IsNullOrEmpty(PdfAConformance);

    /// <summary>PDF/UA identification part from pdfuaid:part.</summary>
    public int? PdfUaPart { get; }

    /// <summary>True when PDF/UA identification metadata was found.</summary>
    public bool HasPdfUaIdentification => PdfUaPart.HasValue;

    /// <summary>Factur-X/ZUGFeRD document type from the XMP extension metadata.</summary>
    public string? ElectronicInvoiceDocumentType { get; }

    /// <summary>Factur-X/ZUGFeRD embedded document file name from the XMP extension metadata.</summary>
    public string? ElectronicInvoiceDocumentFileName { get; }

    /// <summary>Factur-X/ZUGFeRD XMP extension version.</summary>
    public string? ElectronicInvoiceVersion { get; }

    /// <summary>Factur-X/ZUGFeRD conformance level from the XMP extension metadata.</summary>
    public string? ElectronicInvoiceConformanceLevel { get; }

    /// <summary>True when Factur-X/ZUGFeRD XMP extension metadata was found.</summary>
    public bool HasElectronicInvoiceMetadata =>
        !string.IsNullOrEmpty(ElectronicInvoiceDocumentType) ||
        !string.IsNullOrEmpty(ElectronicInvoiceDocumentFileName) ||
        !string.IsNullOrEmpty(ElectronicInvoiceVersion) ||
        !string.IsNullOrEmpty(ElectronicInvoiceConformanceLevel);
}
