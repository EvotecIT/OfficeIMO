namespace OfficeIMO.Pdf;

internal sealed partial class PdfReadDocument {
    /// <summary>Opens a PDF from bytes into the canonical typed object model.</summary>
    internal static PdfReadDocument Open(byte[] pdf, PdfLoadOptions? options = null) =>
        Open(pdf, options, CancellationToken.None);

    internal static PdfReadDocument Open(
        byte[] pdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        return Open((ReadOnlyMemory<byte>)pdf, options, cancellationToken);
    }

    internal static PdfReadDocument Open(
        ReadOnlyMemory<byte> pdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken) =>
        OpenCore(pdf, options, ownsBytes: false, detachOutputIntentProfiles: false, cancellationToken);

    /// <summary>Opens an immutable buffer already owned by OfficeIMO without duplicating stream payloads.</summary>
    internal static PdfReadDocument OpenOwned(
        byte[] ownedPdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(ownedPdf, nameof(ownedPdf));
        return OpenCore(ownedPdf, options, ownsBytes: true, detachOutputIntentProfiles: false, cancellationToken);
    }

    /// <summary>
    /// Opens a buffer that remains stable and available for the complete lifetime of the returned internal reader.
    /// </summary>
    internal static PdfReadDocument OpenBorrowed(
        ReadOnlyMemory<byte> borrowedPdf,
        PdfLoadOptions? options,
        CancellationToken cancellationToken = default) =>
        OpenBorrowed(borrowedPdf, options, detachOutputIntentProfiles: false, cancellationToken);

    internal static PdfReadDocument OpenBorrowed(
        ReadOnlyMemory<byte> borrowedPdf,
        PdfLoadOptions? options,
        bool detachOutputIntentProfiles,
        CancellationToken cancellationToken) =>
        OpenCore(borrowedPdf, options, ownsBytes: true, detachOutputIntentProfiles, cancellationToken);

    internal static PdfReadDocument OpenSeekable(PdfSeekableInput input, PdfLoadOptions options, CancellationToken cancellationToken) {
        using PdfSeekableInput.Admission admission = input.BeginAdmission();
        try {
            return OpenCore(new PdfByteContentText(input, cancellationToken), options, ownsBytes: true,
                detachOutputIntentProfiles: true, cancellationToken);
        } finally {
            // Some PDF syntax paths recover from a damaged object. A failed read
            // cannot be treated as damage in the bytes that were successfully read.
            admission.ThrowIfReadFailed();
        }
    }

    private static PdfReadDocument OpenCore(
        ReadOnlyMemory<byte> pdf,
        PdfLoadOptions? options,
        bool ownsBytes,
        bool detachOutputIntentProfiles,
        CancellationToken cancellationToken) =>
        OpenCore(new PdfByteContentText(pdf), options, ownsBytes, detachOutputIntentProfiles, cancellationToken);

    private static PdfReadDocument OpenCore(
        PdfByteContentText pdf,
        PdfLoadOptions? options,
        bool ownsBytes,
        bool detachOutputIntentProfiles,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf, effectiveOptions, includeParsedDetails: false, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var parsed = PdfSyntax.ParseObjectsCore(pdf, effectiveOptions, out PdfRepairReport repairReport,
            out long decodedStreamBytes, ownsBytes, cancellationToken);
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

        return new PdfReadDocument(map, trailer, security, repairReport, effectiveOptions, decodedStreamBytes,
            detachOutputIntentProfiles, cancellationToken);
    }

    /// <summary>
    /// Opens bytes produced by the canonical rewrite writer. Clear-text outputs avoid the
    /// arbitrary-input security marker scan; protected outputs receive complete authenticated
    /// security inspection. Both paths retain the object parse and semantic repair pass.
    /// </summary>
    internal static PdfReadDocument OpenRewrittenOutput(
        byte[] pdf,
        PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions effectiveOptions = PdfLoadOptions.Resolve(options);
        var (map, trailer) = PdfSyntax.ParseOwnedObjects(
            pdf,
            effectiveOptions,
            out PdfRepairReport repairReport,
            out long decodedStreamBytes,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentSecurityInfo security;
        if (PdfSyntax.ReadTrailerReference(trailer, "Encrypt", effectiveOptions.Limits, cancellationToken) is not null) {
            security = PdfSyntax.ReadDocumentSecurityInfo(pdf, effectiveOptions, includeParsedDetails: false, cancellationToken);
            security = PdfSyntax.ReadDocumentSecurityInfo(pdf, map, trailer, security, repairReport, effectiveOptions, cancellationToken);
        } else {
            security = PdfSyntax.ReadRewrittenOutputSecurityInfo(pdf, trailer, effectiveOptions, cancellationToken);
        }

        return new PdfReadDocument(map, trailer, security, repairReport, effectiveOptions, decodedStreamBytes,
            detachOutputIntentProfiles: false, cancellationToken);
    }

    /// <summary>Opens a PDF from a bounded file snapshot.</summary>
    internal static PdfReadDocument Open(string path, PdfLoadOptions? options = null) =>
        PdfDocumentSource.FromPath(path, options).Read();

    /// <summary>Opens a PDF from a bounded readable stream snapshot.</summary>
    internal static PdfReadDocument Open(Stream stream, PdfLoadOptions? options = null) =>
        PdfDocumentSource.FromStream(stream, options).Read();

    /// <summary>Extracts full‑document plain text (pages separated by blank lines).</summary>
    public string ExtractText() => ExtractText(CancellationToken.None);

    internal string ExtractText(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        // Lenient parsing keeps the readable text when a content stream or a font of the document is damaged;
        // ReadForInspection reports what was left out. Strict parsing fails instead.
        bool keepReadableContent = _options.ParsingMode == PdfParsingMode.Lenient;
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < Pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].ExtractText(cancellationToken, keepReadableContent));
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
