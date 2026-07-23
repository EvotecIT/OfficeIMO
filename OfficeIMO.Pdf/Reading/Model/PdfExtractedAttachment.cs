namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an embedded file attachment extracted from a parsed PDF catalog.
/// </summary>
public sealed class PdfExtractedAttachment {
    private readonly byte[] _bytes;

    internal PdfExtractedAttachment(
        string name,
        string fileName,
        string? unicodeFileName,
        string? description,
        string? mimeType,
        PdfAssociatedFileRelationship relationship,
        string filter,
        int fileSpecObjectNumber,
        int embeddedFileObjectNumber,
        byte[] bytes,
        string source = "Names/EmbeddedFiles",
        DateTimeOffset? creationDate = null,
        DateTimeOffset? modificationDate = null,
        bool copyBytes = true) {
        Name = name;
        FileName = fileName;
        UnicodeFileName = unicodeFileName;
        Description = description;
        MimeType = mimeType;
        Relationship = relationship;
        Filter = filter;
        FileSpecObjectNumber = fileSpecObjectNumber;
        EmbeddedFileObjectNumber = embeddedFileObjectNumber;
        Source = source;
        CreationDate = creationDate;
        ModificationDate = modificationDate;
        _bytes = copyBytes ? (byte[])bytes.Clone() : bytes;
    }

    /// <summary>Name-tree key associated with this embedded file.</summary>
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

    /// <summary>Catalog source that referenced this attachment, for example Names/EmbeddedFiles or AF.</summary>
    public string Source { get; }
    /// <summary>Embedded-file creation date from /Params, when readable.</summary>
    public DateTimeOffset? CreationDate { get; }
    /// <summary>Embedded-file modification date from /Params, when readable.</summary>
    public DateTimeOffset? ModificationDate { get; }

    /// <summary>Decoded embedded file bytes. The returned array is a defensive copy.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Decoded payload size without allocating the defensive public copy.</summary>
    internal int ByteLength => _bytes.Length;
}
