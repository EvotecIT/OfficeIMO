namespace OfficeIMO.Pdf;

/// <summary>Mutable attachment collection used by one existing-document attachment edit.</summary>
public sealed class PdfAttachmentEditSession {
    private readonly List<PdfEmbeddedFile> _attachments;
    private readonly List<PdfAttachmentSourceIdentity> _sourceIdentities;
    private readonly Dictionary<string, string> _retainedOriginalNames;
    private readonly long _maximumAttachmentBytes;
    private long _totalAttachmentBytes;

    internal PdfAttachmentEditSession(IEnumerable<PdfEmbeddedFile> attachments,
        long maximumAttachmentBytes = PdfReadLimits.DefaultMaxTotalAttachmentBytes)
        : this(attachments.Select(static file => new PdfAttachmentEditSource(file, 0, 0)), maximumAttachmentBytes) { }

    internal PdfAttachmentEditSession(IEnumerable<PdfAttachmentEditSource> attachments,
        long maximumAttachmentBytes = PdfReadLimits.DefaultMaxTotalAttachmentBytes) {
        Guard.Positive(maximumAttachmentBytes, nameof(maximumAttachmentBytes));
        _maximumAttachmentBytes = maximumAttachmentBytes;
        _attachments = new List<PdfEmbeddedFile>();
        _sourceIdentities = new List<PdfAttachmentSourceIdentity>();
        foreach (PdfAttachmentEditSource source in attachments) {
            long projectedBytes = GetProjectedBytes(0, source.Attachment.ByteLength);
            _attachments.Add(source.Attachment.Clone());
            _sourceIdentities.Add(new PdfAttachmentSourceIdentity(
                source.FileSpecObjectNumber, source.EmbeddedFileObjectNumber));
            _totalAttachmentBytes = projectedBytes;
        }
        _retainedOriginalNames = CreateRetainedOriginalNames(_attachments);
    }

    private static Dictionary<string, string> CreateRetainedOriginalNames(
        IEnumerable<PdfEmbeddedFile> attachments) {
        var retainedOriginalNames = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (PdfEmbeddedFile attachment in attachments) {
            if (!retainedOriginalNames.ContainsKey(attachment.FileName)) {
                retainedOriginalNames.Add(attachment.FileName, attachment.FileName);
            }
        }
        return retainedOriginalNames;
    }

    /// <summary>Current attachment snapshots in edit order.</summary>
    public IReadOnlyList<PdfEmbeddedFile> Attachments => _attachments.Select(static file => file.Clone()).ToArray();

    /// <summary>Adds a uniquely named attachment after admitting its payload against the target collection budget.</summary>
    /// <exception cref="PdfReadLimitException">The target collection exceeds its aggregate attachment-byte limit.</exception>
    public PdfAttachmentEditSession Add(PdfEmbeddedFile attachment) {
        Guard.NotNull(attachment, nameof(attachment));
        EnsureMissing(attachment.FileName);
        long projectedBytes = GetProjectedBytes(0, attachment.ByteLength);
        _attachments.Add(attachment.Clone());
        _sourceIdentities.Add(new PdfAttachmentSourceIdentity(0, 0));
        _totalAttachmentBytes = projectedBytes;
        return this;
    }

    /// <summary>Replaces the attachment identified by file name, admitting the net payload before changing the collection.</summary>
    /// <exception cref="PdfReadLimitException">The target collection exceeds its aggregate attachment-byte limit.</exception>
    public PdfAttachmentEditSession Replace(string fileName, PdfEmbeddedFile replacement) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName)); Guard.NotNull(replacement, nameof(replacement));
        int index = RequireIndex(fileName);
        int conflict = FindIndex(replacement.FileName);
        if (conflict >= 0 && conflict != index) throw new ArgumentException("An attachment with the replacement file name already exists.", nameof(replacement));
        long projectedBytes = GetProjectedBytes(_attachments[index].ByteLength, replacement.ByteLength);
        PdfEmbeddedFile snapshot = replacement.Clone();
        UpdateRetainedName(_attachments[index].FileName, replacement.FileName);
        _attachments[index] = snapshot;
        _totalAttachmentBytes = projectedBytes;
        return this;
    }

    /// <summary>Renames an attachment while preserving payload and metadata.</summary>
    public PdfAttachmentEditSession Rename(string fileName, string newFileName) {
        int index = RequireIndex(fileName);
        if (!string.Equals(fileName, newFileName, StringComparison.Ordinal)) EnsureMissing(newFileName);
        PdfEmbeddedFile current = _attachments[index];
        UpdateRetainedName(current.FileName, newFileName);
        _attachments[index] = new PdfEmbeddedFile(newFileName, current.DataSnapshot, current.MimeType, current.Relationship, current.Description, current.CreationDate, current.ModificationDate);
        return this;
    }

    /// <summary>Removes an attachment by file name.</summary>
    public PdfAttachmentEditSession Remove(string fileName) {
        int index = RequireIndex(fileName);
        return RemoveAt(index);
    }

    /// <summary>Removes one source attachment by its file-spec and embedded-stream identity.</summary>
    public PdfAttachmentEditSession Remove(PdfAttachmentInfo attachment) {
        Guard.NotNull(attachment, nameof(attachment));
        bool hasObjectIdentity = attachment.FileSpecObjectNumber != 0 || attachment.EmbeddedFileObjectNumber != 0;
        int matched = -1;
        for (int index = 0; index < _attachments.Count; index++) {
            PdfAttachmentSourceIdentity source = _sourceIdentities[index];
            bool matches = hasObjectIdentity
                ? source.FileSpecObjectNumber == attachment.FileSpecObjectNumber &&
                  source.EmbeddedFileObjectNumber == attachment.EmbeddedFileObjectNumber &&
                  string.Equals(_attachments[index].FileName, attachment.FileName, StringComparison.Ordinal)
                : source.FileSpecObjectNumber == 0 && source.EmbeddedFileObjectNumber == 0 &&
                  string.Equals(_attachments[index].FileName, attachment.FileName, StringComparison.Ordinal);
            if (!matches) continue;
            if (matched >= 0) throw new InvalidOperationException("The selected PDF attachment identity is ambiguous.");
            matched = index;
        }
        if (matched < 0) throw new KeyNotFoundException("The selected PDF attachment is no longer present.");
        return RemoveAt(matched);
    }

    internal PdfAttachmentEditSession RemoveAt(int index) {
        if (index < 0 || index >= _attachments.Count) throw new ArgumentOutOfRangeException(nameof(index));
        string fileName = _attachments[index].FileName;
        _totalAttachmentBytes -= _attachments[index].ByteLength;
        _attachments.RemoveAt(index);
        _sourceIdentities.RemoveAt(index);
        string? originalName = _retainedOriginalNames.FirstOrDefault(pair => string.Equals(pair.Value, fileName, StringComparison.Ordinal)).Key;
        if (originalName != null && !_attachments.Any(file => string.Equals(file.FileName, fileName, StringComparison.Ordinal)))
            _retainedOriginalNames.Remove(originalName);
        return this;
    }

    internal IReadOnlyList<PdfEmbeddedFile> Snapshot() => _attachments.Select(static file => file.Clone()).ToArray();
    internal IReadOnlyList<PdfAttachmentEditEntry> SnapshotEntries() => _attachments
        .Select((file, index) => new PdfAttachmentEditEntry(file.Clone(), _sourceIdentities[index]))
        .ToArray();
    internal IReadOnlyDictionary<string, string> RetainedOriginalNames => _retainedOriginalNames;

    private long GetProjectedBytes(long removedBytes, long addedBytes) {
        long retainedBytes = _totalAttachmentBytes - removedBytes;
        if (addedBytes > _maximumAttachmentBytes - retainedBytes) {
            long actual = retainedBytes > long.MaxValue - addedBytes ? long.MaxValue : retainedBytes + addedBytes;
            throw PdfReadLimitException.Create(PdfReadLimitKind.AttachmentBytes, _maximumAttachmentBytes, actual);
        }
        return retainedBytes + addedBytes;
    }

    private void UpdateRetainedName(string currentName, string newName) {
        string? originalName = _retainedOriginalNames.FirstOrDefault(pair => string.Equals(pair.Value, currentName, StringComparison.Ordinal)).Key;
        if (originalName != null) _retainedOriginalNames[originalName] = newName;
    }

    private int RequireIndex(string fileName) { Guard.NotNullOrWhiteSpace(fileName, nameof(fileName)); int index = FindIndex(fileName); return index >= 0 ? index : throw new KeyNotFoundException("PDF attachment was not found: " + fileName); }
    private void EnsureMissing(string fileName) { if (FindIndex(fileName) >= 0) throw new ArgumentException("A PDF attachment with this file name already exists: " + fileName, nameof(fileName)); }
    private int FindIndex(string fileName) => _attachments.FindIndex(file => string.Equals(file.FileName, fileName, StringComparison.Ordinal));
}

internal sealed class PdfAttachmentEditSource {
    internal PdfAttachmentEditSource(
        PdfEmbeddedFile attachment,
        int fileSpecObjectNumber,
        int embeddedFileObjectNumber) {
        Attachment = attachment;
        FileSpecObjectNumber = fileSpecObjectNumber;
        EmbeddedFileObjectNumber = embeddedFileObjectNumber;
    }

    internal PdfEmbeddedFile Attachment { get; }
    internal int FileSpecObjectNumber { get; }
    internal int EmbeddedFileObjectNumber { get; }
}

internal sealed class PdfAttachmentEditEntry {
    internal PdfAttachmentEditEntry(
        PdfEmbeddedFile attachment,
        PdfAttachmentSourceIdentity sourceIdentity) {
        Attachment = attachment;
        SourceIdentity = sourceIdentity;
    }

    internal PdfEmbeddedFile Attachment { get; }
    internal PdfAttachmentSourceIdentity SourceIdentity { get; }
}

internal readonly struct PdfAttachmentSourceIdentity {
    internal PdfAttachmentSourceIdentity(
        int fileSpecObjectNumber,
        int embeddedFileObjectNumber) {
        FileSpecObjectNumber = fileSpecObjectNumber;
        EmbeddedFileObjectNumber = embeddedFileObjectNumber;
    }

    internal int FileSpecObjectNumber { get; }
    internal int EmbeddedFileObjectNumber { get; }
}
