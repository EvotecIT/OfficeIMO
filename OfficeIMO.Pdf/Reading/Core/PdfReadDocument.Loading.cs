namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Opens a PDF from bytes into the canonical typed object model.</summary>
    public static PdfReadDocument Open(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfReadOptions effectiveOptions = PdfReadOptions.Resolve(options);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, effectiveOptions);
        var (map, trailer) = PdfSyntax.ParseObjects(pdf, effectiveOptions, out PdfRepairReport repairReport);
        if (effectiveOptions.Password is not null && security.HasEncryption) {
            security = PdfSyntax.ReadDocumentSecurityInfo(pdf, map, trailer, security);
        }

        return new PdfReadDocument(map, trailer, security, repairReport, effectiveOptions);
    }

    /// <summary>Opens a PDF from a bounded file snapshot.</summary>
    public static PdfReadDocument Open(string path, PdfReadOptions? options = null) =>
        PdfDocumentSource.FromPath(path, options).Read();

    /// <summary>Opens a PDF from a bounded readable stream snapshot.</summary>
    public static PdfReadDocument Open(Stream stream, PdfReadOptions? options = null) =>
        PdfDocumentSource.FromStream(stream, options).Read();

    /// <summary>Extracts full‑document plain text (pages separated by blank lines).</summary>
    public string ExtractText() {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < Pages.Count; i++) {
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].ExtractText());
        }
        return sb.ToString();
    }

    /// <summary>Extracts image XObjects from all pages in page order.</summary>
    public IReadOnlyList<PdfExtractedImage> ExtractImages() => PdfImageExtractor.ExtractImages(this);

    /// <summary>Extracts embedded file attachments from the document catalog.</summary>
    public IReadOnlyList<PdfExtractedAttachment> ExtractAttachments() => PdfAttachmentExtractor.ExtractAttachments(_objects, _trailerRaw, _options.Limits);
}
