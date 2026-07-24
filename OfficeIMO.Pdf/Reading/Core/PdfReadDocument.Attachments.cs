namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Embedded and associated file attachment metadata discovered from the document catalog.</summary>
    public IReadOnlyList<PdfAttachmentInfo> Attachments => ReadLogicalContent(_attachments);

    private IReadOnlyList<PdfAttachmentInfo> ExtractAttachmentInfos() {
        // Catalog inspection must remain available for preflight even when payload extraction is restricted.
        return PdfAttachmentExtractor.InspectAttachments(_objects, _trailerRaw, _options.Limits);
    }
}
