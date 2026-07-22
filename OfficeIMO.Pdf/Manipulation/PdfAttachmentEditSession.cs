namespace OfficeIMO.Pdf;

/// <summary>Mutable attachment collection used by one existing-document attachment edit.</summary>
public sealed class PdfAttachmentEditSession {
    private readonly List<PdfEmbeddedFile> _attachments;
    private readonly Dictionary<string, string> _retainedOriginalNames;

    internal PdfAttachmentEditSession(IEnumerable<PdfEmbeddedFile> attachments) {
        _attachments = attachments.Select(static file => file.Clone()).ToList();
        _retainedOriginalNames = _attachments.ToDictionary(
            static file => file.FileName,
            static file => file.FileName,
            StringComparer.Ordinal);
    }

    /// <summary>Current attachment snapshots in edit order.</summary>
    public IReadOnlyList<PdfEmbeddedFile> Attachments => _attachments.Select(static file => file.Clone()).ToArray();

    /// <summary>Adds a uniquely named attachment.</summary>
    public PdfAttachmentEditSession Add(PdfEmbeddedFile attachment) {
        Guard.NotNull(attachment, nameof(attachment));
        EnsureMissing(attachment.FileName);
        _attachments.Add(attachment.Clone());
        return this;
    }

    /// <summary>Replaces the attachment identified by file name.</summary>
    public PdfAttachmentEditSession Replace(string fileName, PdfEmbeddedFile replacement) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName)); Guard.NotNull(replacement, nameof(replacement));
        int index = RequireIndex(fileName);
        int conflict = FindIndex(replacement.FileName);
        if (conflict >= 0 && conflict != index) throw new ArgumentException("An attachment with the replacement file name already exists.", nameof(replacement));
        UpdateRetainedName(_attachments[index].FileName, replacement.FileName);
        _attachments[index] = replacement.Clone();
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
        _attachments.RemoveAt(RequireIndex(fileName));
        string? originalName = _retainedOriginalNames.FirstOrDefault(pair => string.Equals(pair.Value, fileName, StringComparison.Ordinal)).Key;
        if (originalName != null) _retainedOriginalNames.Remove(originalName);
        return this;
    }

    internal IReadOnlyList<PdfEmbeddedFile> Snapshot() => _attachments.Select(static file => file.Clone()).ToArray();
    internal IReadOnlyDictionary<string, string> RetainedOriginalNames => _retainedOriginalNames;

    private void UpdateRetainedName(string currentName, string newName) {
        string? originalName = _retainedOriginalNames.FirstOrDefault(pair => string.Equals(pair.Value, currentName, StringComparison.Ordinal)).Key;
        if (originalName != null) _retainedOriginalNames[originalName] = newName;
    }

    private int RequireIndex(string fileName) { Guard.NotNullOrWhiteSpace(fileName, nameof(fileName)); int index = FindIndex(fileName); return index >= 0 ? index : throw new KeyNotFoundException("PDF attachment was not found: " + fileName); }
    private void EnsureMissing(string fileName) { if (FindIndex(fileName) >= 0) throw new ArgumentException("A PDF attachment with this file name already exists: " + fileName, nameof(fileName)); }
    private int FindIndex(string fileName) => _attachments.FindIndex(file => string.Equals(file.FileName, fileName, StringComparison.Ordinal));
}
