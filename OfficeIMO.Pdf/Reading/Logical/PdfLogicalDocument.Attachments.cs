namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalDocument {
    /// <summary>Embedded and associated file attachment metadata discovered from the document catalog.</summary>
    public IReadOnlyList<PdfAttachmentInfo> Attachments { get; }

    /// <summary>Number of embedded and associated file attachments discovered from the document catalog.</summary>
    public int AttachmentCount => Attachments.Count;

    /// <summary>True when at least one embedded or associated file attachment was discovered.</summary>
    public bool HasAttachments => AttachmentCount > 0;

    /// <summary>Attachment name-tree keys or associated-file fallback names in first-seen order.</summary>
    public IReadOnlyList<string> AttachmentNames => Attachments.Select(attachment => attachment.Name).ToArray();

    /// <summary>Distinct attachment file names in first-seen order.</summary>
    public IReadOnlyList<string> AttachmentFileNames => Attachments.Select(attachment => attachment.FileName).Distinct(StringComparer.Ordinal).ToArray();

    /// <summary>Distinct catalog attachment sources in first-seen order.</summary>
    public IReadOnlyList<string> AttachmentSources => Attachments.Select(attachment => attachment.Source).Distinct(StringComparer.Ordinal).ToArray();

    /// <summary>Returns attachments with a matching file specification file name.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsByFileName(string fileName) {
        Guard.NotNullOrWhiteSpace(fileName, nameof(fileName));
        return Attachments.Where(attachment => string.Equals(attachment.FileName, fileName, StringComparison.Ordinal)).ToArray();
    }

    /// <summary>Returns attachments from a matching catalog source.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsBySource(string source) {
        Guard.NotNullOrWhiteSpace(source, nameof(source));
        return Attachments.Where(attachment => string.Equals(attachment.Source, source, StringComparison.Ordinal)).ToArray();
    }

    /// <summary>Returns attachments with a matching associated-file relationship.</summary>
    public IReadOnlyList<PdfAttachmentInfo> GetAttachmentsByRelationship(PdfAssociatedFileRelationship relationship) {
        return Attachments.Where(attachment => attachment.Relationship == relationship).ToArray();
    }
}
