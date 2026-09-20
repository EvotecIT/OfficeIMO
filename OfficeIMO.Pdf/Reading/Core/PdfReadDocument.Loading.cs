using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Opens a PDF from bytes into the canonical typed object model.</summary>
    public static PdfReadDocument Open(byte[] pdf, PdfLoadOptions? options = null) =>
        Open(pdf, options, CancellationToken.None);

    internal static PdfReadDocument Open(
        byte[] pdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) =>
        OpenCore(pdf, options, ownsBytes: false, cancellationToken);

    /// <summary>Opens an immutable buffer already owned by OfficeIMO without duplicating stream payloads.</summary>
    internal static PdfReadDocument OpenOwned(
        byte[] ownedPdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        OpenCore(ownedPdf, options, ownsBytes: true, cancellationToken);

    private static PdfReadDocument OpenCore(
        byte[] pdf,
        PdfLoadOptions? options,
        bool ownsBytes,
        CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            effectiveOptions,
            includeParsedDetails: false,
            out string decodedText,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfRepairReport repairReport;
        long decodedStreamBytes;
        (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) parsed = ownsBytes
            ? PdfSyntax.ParseOwnedObjects(
                pdf,
                effectiveOptions,
                out repairReport,
                out decodedStreamBytes,
                decodedText,
                cancellationToken)
            : PdfSyntax.ParseObjects(
                pdf,
                effectiveOptions,
                out repairReport,
                out decodedStreamBytes,
                decodedText,
                cancellationToken);
        Dictionary<int, PdfIndirectObject> map = parsed.Map;
        string trailer = parsed.TrailerRaw;
        cancellationToken.ThrowIfCancellationRequested();
        security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            map,
            trailer,
            security,
            repairReport,
            effectiveOptions,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        return new PdfReadDocument(map, trailer, security, repairReport, effectiveOptions, decodedStreamBytes, cancellationToken);
    }

    /// <summary>
    /// Opens bytes produced by the canonical clear-text rewrite writer without repeating the
    /// arbitrary-input security marker scan. The complete object parse and semantic repair pass
    /// still run, and the resulting document remains the canonical cached readback.
    /// </summary>
    internal static PdfReadDocument OpenRewrittenOutput(
        byte[] pdf,
        PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        string decodedText = PdfEncoding.Latin1GetString(pdf);
        var (map, trailer) = PdfSyntax.ParseOwnedObjects(
            pdf,
            effectiveOptions,
            out PdfRepairReport repairReport,
            out long decodedStreamBytes,
            decodedText,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentSecurityInfo security = PdfSyntax.ReadRewrittenOutputSecurityInfo(
            decodedText,
            trailer,
            effectiveOptions);

        return new PdfReadDocument(map, trailer, security, repairReport, effectiveOptions, decodedStreamBytes, cancellationToken);
    }

    /// <summary>Opens a PDF from a bounded file snapshot.</summary>
    public static PdfReadDocument Open(string path, PdfLoadOptions? options = null) =>
        PdfDocumentSource.FromPath(path, options).Read();

    /// <summary>Opens a PDF from a bounded readable stream snapshot.</summary>
    public static PdfReadDocument Open(Stream stream, PdfLoadOptions? options = null) =>
        PdfDocumentSource.FromStream(stream, options).Read();

    /// <summary>Extracts full‑document plain text (pages separated by blank lines).</summary>
    public string ExtractText() => ExtractText(CancellationToken.None);

    internal string ExtractText(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < Pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].ExtractText(cancellationToken));
        }
        return sb.ToString();
    }

    /// <summary>Extracts image XObjects from all pages in page order.</summary>
    public IReadOnlyList<PdfExtractedImage> ExtractImages() {
        DemandContentExtraction("image");
        return PdfImageExtractor.ExtractImages(this);
    }

    /// <summary>Extracts embedded file attachments from the document catalog.</summary>
    public IReadOnlyList<PdfExtractedAttachment> ExtractAttachments() =>
        ExtractAttachments(CancellationToken.None);

    internal IReadOnlyList<PdfExtractedAttachment> ExtractAttachments(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        DemandContentExtraction("attachment");
        return PdfAttachmentExtractor.ExtractAttachments(
            this,
            static _ => true,
            _options.Limits.MaxTotalAttachmentBytes,
            _options.Limits.MaxDecodedStreamBytes,
            cancellationToken: cancellationToken);
    }

    internal void DemandTextExtraction() => PdfPermissionAuthorization.DemandTextExtraction(Security, _options.PermissionPolicy);

    internal void DemandContentExtraction(string contentName) => PdfPermissionAuthorization.DemandContentExtraction(Security, _options.PermissionPolicy, contentName);
}
