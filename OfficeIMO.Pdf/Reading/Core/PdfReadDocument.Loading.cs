namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Opens a PDF from bytes into the canonical typed object model.</summary>
    public static PdfReadDocument Open(byte[] pdf, PdfLoadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            effectiveOptions,
            includeParsedDetails: false);
        var (map, trailer) = PdfSyntax.ParseObjects(pdf, effectiveOptions, out PdfRepairReport repairReport, out long decodedStreamBytes);
        security = PdfSyntax.ReadDocumentSecurityInfo(pdf, map, trailer, security, effectiveOptions);

        return new PdfReadDocument(map, trailer, security, repairReport, effectiveOptions, decodedStreamBytes);
    }

    /// <summary>Opens a PDF from a bounded file snapshot.</summary>
    public static PdfReadDocument Open(string path, PdfLoadOptions? options = null) =>
        PdfDocumentSource.FromPath(path, options).Read();

    /// <summary>Opens a PDF from a bounded readable stream snapshot.</summary>
    public static PdfReadDocument Open(Stream stream, PdfLoadOptions? options = null) =>
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
    public IReadOnlyList<PdfExtractedImage> ExtractImages() {
        DemandContentExtraction("image");
        return PdfImageExtractor.ExtractImages(this);
    }

    /// <summary>Extracts embedded file attachments from the document catalog.</summary>
    public IReadOnlyList<PdfExtractedAttachment> ExtractAttachments() {
        DemandContentExtraction("attachment");
        return PdfAttachmentExtractor.ExtractAttachments(
            this,
            static _ => true,
            _options.Limits.MaxTotalAttachmentBytes,
            _options.Limits.MaxDecodedStreamBytes);
    }

    internal void DemandTextExtraction() => PdfPermissionAuthorization.DemandTextExtraction(Security, _options.PermissionPolicy);

    internal void DemandContentExtraction(string contentName) => PdfPermissionAuthorization.DemandContentExtraction(Security, _options.PermissionPolicy, contentName);
}
