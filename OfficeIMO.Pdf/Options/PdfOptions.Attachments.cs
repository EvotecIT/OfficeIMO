namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private long _maxEmbeddedFileBytes = PdfAttachmentAdmission.DefaultMaximumFileBytes;
    private long _maxTotalEmbeddedFileBytes = PdfAttachmentAdmission.DefaultMaximumTotalBytes;
    private long _totalEmbeddedFileBytes;

    /// <summary>Maximum payload bytes in each catalog attachment admitted or attachment emitted by this generated document. Defaults to 128 MiB.</summary>
    /// <remarks>Must be positive. Set an explicit larger limit before adding trusted larger files. This limits payload bytes, not defensive copies or peak process memory.</remarks>
    public long MaxEmbeddedFileBytes {
        get => _maxEmbeddedFileBytes;
        set {
            PdfAttachmentAdmission.ValidateMaximum(value, nameof(MaxEmbeddedFileBytes));
            ValidateEmbeddedFileLimits(value, _maxTotalEmbeddedFileBytes);
            _maxEmbeddedFileBytes = value;
        }
    }

    /// <summary>Maximum aggregate catalog payload bytes admitted and distinct attachment payload bytes emitted by this generated document. Defaults to 256 MiB.</summary>
    /// <remarks>Must be positive. Emission includes catalog and structure-associated files; identical structure associations sharing one emitted file count once. This does not limit snapshot multiplicity, serialized PDF size, or peak process memory.</remarks>
    public long MaxTotalEmbeddedFileBytes {
        get => _maxTotalEmbeddedFileBytes;
        set {
            PdfAttachmentAdmission.ValidateMaximum(value, nameof(MaxTotalEmbeddedFileBytes));
            ValidateEmbeddedFileLimits(_maxEmbeddedFileBytes, value);
            _maxTotalEmbeddedFileBytes = value;
        }
    }

    /// <summary>Embedded files associated with the generated document catalog.</summary>
    public IReadOnlyList<PdfEmbeddedFile> EmbeddedFiles =>
        (IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? Array.Empty<PdfEmbeddedFile>();

    internal IReadOnlyList<PdfEmbeddedFile> EmbeddedFileSnapshots =>
        (IReadOnlyList<PdfEmbeddedFile>?)CloneEmbeddedFiles(_embeddedFiles) ?? Array.Empty<PdfEmbeddedFile>();

    /// <summary>Borrowed owned descriptions for synchronous read-only writer admission and metadata inspection. Neither descriptions nor payload ownership may be changed.</summary>
    internal IReadOnlyList<PdfEmbeddedFile> EmbeddedFilesForInspection =>
        _embeddedFiles ?? (IReadOnlyList<PdfEmbeddedFile>)Array.Empty<PdfEmbeddedFile>();

    /// <summary>Adds an embedded file associated with the generated PDF catalog.</summary>
    /// <exception cref="InvalidDataException">The payload or projected catalog collection exceeds its configured byte limit.</exception>
    public PdfCreateOptions AddEmbeddedFile(PdfEmbeddedFile file) {
        Guard.NotNull(file, nameof(file));
        long projectedBytes = PrepareEmbeddedFileAdmission(file.FileName, file.ByteLength);
        PdfEmbeddedFile snapshot = file.Clone();
        (_embeddedFiles ??= new List<PdfEmbeddedFile>()).Add(snapshot);
        _totalEmbeddedFileBytes = projectedBytes;
        return this;
    }

    /// <summary>Adds an embedded file associated with the generated PDF catalog, checking byte limits before copying caller-owned bytes.</summary>
    /// <exception cref="InvalidDataException">The payload or projected catalog collection exceeds its configured byte limit.</exception>
    public PdfCreateOptions AddEmbeddedFile(
        string fileName,
        byte[] data,
        string? mimeType = null,
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Unspecified,
        string? description = null) {
        Guard.NotNullOrEmpty(data, nameof(data));
        PrepareEmbeddedFileAdmission(fileName, data.LongLength);
        return AddEmbeddedFile(new PdfEmbeddedFile(fileName, data, mimeType, relationship, description));
    }

    /// <summary>Sets the modification date emitted in the embedded-file stream parameters for an existing attachment.</summary>
    public PdfCreateOptions SetEmbeddedFileModificationDate(string fileName, DateTimeOffset modificationDate) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName));
        if (_embeddedFiles != null) {
            foreach (PdfEmbeddedFile file in _embeddedFiles) {
                if (string.Equals(file.FileName, fileName, StringComparison.Ordinal)) {
                    file.ModificationDate = modificationDate;
                    return this;
                }
            }
        }
        throw new ArgumentException("No embedded file with the requested name is configured.", nameof(fileName));
    }

    /// <summary>Removes all embedded files associated with the generated PDF catalog.</summary>
    public PdfCreateOptions ClearEmbeddedFiles() {
        _embeddedFiles?.Clear();
        _totalEmbeddedFileBytes = 0L;
        return this;
    }

    /// <summary>Rechecks owned catalog payload limits before options cloning or generation without making payload copies.</summary>
    internal void ValidateEmbeddedFileLimits() => ValidateEmbeddedFileLimits(_maxEmbeddedFileBytes, _maxTotalEmbeddedFileBytes);

    private void ValidateEmbeddedFileLimits(long maximumFileBytes, long maximumTotalBytes) {
        long totalBytes = 0L;
        if (_embeddedFiles == null) return;
        foreach (PdfEmbeddedFile file in _embeddedFiles) {
            PdfAttachmentAdmission.EnsureFileBytes(file.ByteLength, maximumFileBytes, file.FileName);
            totalBytes = PdfAttachmentAdmission.GetProjectedBytes(totalBytes, file.ByteLength, maximumTotalBytes);
        }
    }

    private long PrepareEmbeddedFileAdmission(string fileName, long byteLength) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName));
        if (_embeddedFiles?.Any(file => string.Equals(file.FileName, fileName, StringComparison.Ordinal)) == true)
            throw new ArgumentException("PDF embedded file names must be unique within a generated document.", nameof(fileName));
        PdfAttachmentAdmission.EnsureFileBytes(byteLength, MaxEmbeddedFileBytes, fileName);
        return PdfAttachmentAdmission.GetProjectedBytes(_totalEmbeddedFileBytes, byteLength, MaxTotalEmbeddedFileBytes);
    }
}
