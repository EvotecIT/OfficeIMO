namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Embedded and associated file attachment metadata discovered from the document catalog.</summary>
    public IReadOnlyList<PdfAttachmentInfo> Attachments => ReadLogicalContent(_attachments);

    private IReadOnlyList<PdfAttachmentInfo> ExtractAttachmentInfos() {
        // Catalog inspection must remain available for preflight even when payload extraction is restricted.
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(_objects, _trailerRaw, _options.Limits);
        if (attachments.Count == 0) {
            return Array.Empty<PdfAttachmentInfo>();
        }

        var result = new List<PdfAttachmentInfo>(attachments.Count);
        for (int i = 0; i < attachments.Count; i++) {
            PdfExtractedAttachment attachment = attachments[i];
            result.Add(new PdfAttachmentInfo(
                attachment.Name,
                attachment.FileName,
                attachment.UnicodeFileName,
                attachment.Description,
                attachment.MimeType,
                attachment.Relationship,
                attachment.Filter,
                attachment.FileSpecObjectNumber,
                attachment.EmbeddedFileObjectNumber,
                attachment.Bytes.Length,
                attachment.Source,
                attachment.CreationDate,
                attachment.ModificationDate));
        }

        return result.AsReadOnly();
    }
}
