namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata for an embedded or associated PDF file attachment without exposing the attachment payload.
/// </summary>
public sealed class PdfAttachmentInfo {
    internal PdfAttachmentInfo(
        string name,
        string fileName,
        string? unicodeFileName,
        string? description,
        string? mimeType,
        PdfAssociatedFileRelationship relationship,
        string filter,
        int fileSpecObjectNumber,
        int embeddedFileObjectNumber,
        int encodedSizeBytes,
        int? declaredSizeBytes,
        string source,
        DateTimeOffset? creationDate,
        DateTimeOffset? modificationDate) {
        Name = name;
        FileName = fileName;
        UnicodeFileName = unicodeFileName;
        Description = description;
        MimeType = mimeType;
        Relationship = relationship;
        Filter = filter;
        FileSpecObjectNumber = fileSpecObjectNumber;
        EmbeddedFileObjectNumber = embeddedFileObjectNumber;
        EncodedSizeBytes = encodedSizeBytes;
        DeclaredSizeBytes = declaredSizeBytes;
        Source = source;
        CreationDate = creationDate;
        ModificationDate = modificationDate;
    }

    /// <summary>Name-tree key or associated-file fallback name for this attachment.</summary>
    public string Name { get; }

    /// <summary>File name from the file specification dictionary.</summary>
    public string FileName { get; }

    /// <summary>Unicode file name from /UF, when present.</summary>
    public string? UnicodeFileName { get; }

    /// <summary>Human-readable file description from /Desc, when present.</summary>
    public string? Description { get; }

    /// <summary>MIME type decoded from the embedded file stream /Subtype name, when present.</summary>
    public string? MimeType { get; }

    /// <summary>Associated-file relationship from /AFRelationship, or Unspecified when absent.</summary>
    public PdfAssociatedFileRelationship Relationship { get; }

    /// <summary>PDF stream filter name or comma-separated filter names when present.</summary>
    public string Filter { get; }

    /// <summary>Object number of the file specification dictionary, or 0 for a direct dictionary.</summary>
    public int FileSpecObjectNumber { get; }

    /// <summary>Object number of the embedded file stream, or 0 for a direct stream.</summary>
    public int EmbeddedFileObjectNumber { get; }

    /// <summary>
    /// Encoded attachment stream size in bytes. This is a compatibility alias for <see cref="EncodedSizeBytes"/>;
    /// it is not the decoded attachment size when stream filters are present.
    /// </summary>
    public int SizeBytes => EncodedSizeBytes;

    /// <summary>Exact encoded attachment stream size stored in the PDF.</summary>
    public int EncodedSizeBytes { get; }

    /// <summary>
    /// Untrusted decoded-size claim from the embedded-file /Params /Size entry, when present.
    /// Callers must not use this value as a resource limit.
    /// </summary>
    public int? DeclaredSizeBytes { get; }

    /// <summary>
    /// Decoded attachment size is unknown during metadata-only inspection. Extract the attachment under
    /// <see cref="PdfReadLimits.MaxTotalAttachmentBytes"/> and inspect its payload length when the exact value is required.
    /// </summary>
    public int? DecodedSizeBytes { get; }

    /// <summary>Catalog source that referenced this attachment, for example Names/EmbeddedFiles or AF.</summary>
    public string Source { get; }
    /// <summary>Embedded-file creation date, when readable.</summary>
    public DateTimeOffset? CreationDate { get; }
    /// <summary>Embedded-file modification date, when readable.</summary>
    public DateTimeOffset? ModificationDate { get; }

    /// <summary>True when the attachment was referenced from the catalog /AF associated-files array.</summary>
    public bool IsAssociatedFile => string.Equals(Source, "AF", StringComparison.Ordinal);
}
