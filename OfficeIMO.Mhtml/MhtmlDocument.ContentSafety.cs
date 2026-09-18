using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Email;
using OfficeIMO.Html;

namespace OfficeIMO.Mhtml;

public sealed partial class MhtmlDocument {
    /// <summary>Inspects concealed machine-readable HTML in a bounded MHTML archive.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] archiveBytes,
        OfficeContentSafetyOptions? options = null,
        EmailReaderOptions? mimeOptions = null,
        CancellationToken cancellationToken = default) {
        if (archiveBytes == null) throw new ArgumentNullException(nameof(archiveBytes));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        OfficeContentSafetyInputGuard.ValidateBytes(archiveBytes, effective);
        MhtmlDocument document = LoadForContentSafety(archiveBytes, effective, mimeOptions, cancellationToken);
        return InspectDocumentContentSafetyAsync(document, effective, cancellationToken).GetAwaiter().GetResult();
    }

    /// <summary>Inspects concealed machine-readable HTML in a bounded MHTML file.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null,
        EmailReaderOptions? mimeOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        return InspectContentSafety(OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective), effective, mimeOptions, cancellationToken);
    }

    /// <summary>Removes exact selected concealed HTML findings, then reopens and reinspects the MHTML archive.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] archiveBytes,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        EmailReaderOptions? mimeOptions = null,
        CancellationToken cancellationToken = default) {
        if (archiveBytes == null) throw new ArgumentNullException(nameof(archiveBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        OfficeContentSafetyInputGuard.ValidateBytes(archiveBytes, options.Inspection);
        MhtmlDocument document = LoadForContentSafety(archiveBytes, options.Inspection, mimeOptions, cancellationToken);
        HtmlRenderOptions renderOptions = CreateContentSafetyRenderOptions(document, options.Inspection);
        var part = new HtmlContentSafetyPackagePart("root", document.Html, "MHTML/Root", renderOptions);
        HtmlContentSafetyPackageCleanupResult cleaned = HtmlContentSafety.RemoveSelectedPackagePartsAsync(
            "MHTML",
            new[] { part },
            selection,
            options.Inspection,
            cancellationToken).GetAwaiter().GetResult();
        if (cleaned.Changes.Count == 0) {
            return new OfficeContentCleanupResult((byte[])archiveBytes.Clone(), cleaned.Before, cleaned.Before, cleaned.Changes);
        }
        if (document._mimeDocument.Protection.IsProtected) {
            throw new InvalidOperationException(
                "MHTML cleanup cannot mutate a signed or encrypted MIME wrapper. Verify and unwrap or decrypt it before cleanup.");
        }
        ApplyTransportSignatureMutationPolicy(document._mimeDocument, options.SignatureMutationPolicy);
        RemoveOuterPayloadDependentHeaders(document._mimeDocument);

        document._mimeDocument.Body.Html = cleaned.Parts["root"];
        document._mimeDocument.Body.PreserveHtmlMimeHeadersOnWrite = true;
        foreach (EmailAttachment attachment in document._mimeDocument.Attachments) {
            attachment.PreserveMimeHeadersOnWrite = true;
        }
        byte[] output = document.ToBytes(new EmailWriterOptions(maxOutputBytes: options.Inspection.MaxExpandedPackageBytes));
        OfficeContentSafetyReport after = InspectContentSafety(output, options.Inspection, mimeOptions, cancellationToken);
        return new OfficeContentCleanupResult(output, cleaned.Before, after, cleaned.Changes);
    }

    private static void ApplyTransportSignatureMutationPolicy(
        EmailDocument document,
        OfficeSignatureMutationPolicy policy) {
        bool hasBodyCoveringSignature = document.Headers.Any(header =>
            IsBodyCoveringTransportSignatureHeader(header.Name));
        if (!hasBodyCoveringSignature) return;
        if (policy == OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException(
                "MHTML cleanup cannot retain a DKIM or ARC signature after mutating the signed body. " +
                "Use RemoveInvalidatedSignatures to remove the invalidated transport-signature headers explicitly.");
        }
        if (policy != OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) return;
        for (int index = document.Headers.Count - 1; index >= 0; index--) {
            if (IsTransportSignatureChainHeader(document.Headers[index].Name)) document.Headers.RemoveAt(index);
        }
    }

    private static void RemoveOuterPayloadDependentHeaders(EmailDocument document) {
        for (int index = document.Headers.Count - 1; index >= 0; index--) {
            string name = document.Headers[index].Name;
            if (string.Equals(name, "Content-Length", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "Content-MD5", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "Content-Digest", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "Repr-Digest", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "Digest", StringComparison.OrdinalIgnoreCase)) {
                document.Headers.RemoveAt(index);
            }
        }
    }

    private static bool IsBodyCoveringTransportSignatureHeader(string name) =>
        string.Equals(name, "DKIM-Signature", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "DomainKey-Signature", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "ARC-Message-Signature", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "ARC-Seal", StringComparison.OrdinalIgnoreCase);

    private static bool IsTransportSignatureChainHeader(string name) =>
        IsBodyCoveringTransportSignatureHeader(name)
        || string.Equals(name, "ARC-Authentication-Results", StringComparison.OrdinalIgnoreCase);

    /// <summary>Atomically writes an explicitly cleaned MHTML artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        EmailReaderOptions? mimeOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        byte[] input = OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection);
        OfficeContentCleanupResult result = RemoveSelectedContent(input, selection, options, mimeOptions, cancellationToken);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static MhtmlDocument LoadForContentSafety(
        byte[] archiveBytes,
        OfficeContentSafetyOptions options,
        EmailReaderOptions? mimeOptions,
        CancellationToken cancellationToken) {
        EmailReaderOptions readerOptions = IntersectReaderOptions(mimeOptions ?? EmailReaderOptions.Default, options);
        using var stream = new MemoryStream(archiveBytes, writable: false);
        MhtmlDocument document = Load(stream, readerOptions, cancellationToken: cancellationToken);
        if (!MimeTextCodec.IsSupportedTransferEncoding(document._mimeDocument.Body.HtmlTransferEncoding)
            || document._mimeDocument.Body.HtmlMimeTransferDecodingWasAmbiguous
            || (!string.IsNullOrWhiteSpace(document._mimeDocument.Body.HtmlCharset)
                && document._mimeDocument.Body.HtmlMimeDecodingWasAmbiguous)
            || document._mimeDocument.Body.HtmlWebDecodingWasAmbiguous) {
            throw new InvalidDataException(
                "MHTML content-safety inspection requires the selected HTML root to use an unambiguous supported MIME encoding.");
        }
        EmailDiagnostic? ambiguous = document.MimeDiagnostics.FirstOrDefault(diagnostic =>
            diagnostic.Code == MhtmlDiagnosticCodes.DuplicateContentId
            || diagnostic.Code == MhtmlDiagnosticCodes.DuplicateContentLocation
            || diagnostic.Code == MhtmlDiagnosticCodes.DuplicateResourceIdentity
            || diagnostic.Code == MhtmlDiagnosticCodes.InvalidContentLocation
            || diagnostic.Code == MimeHeaderParser.DuplicateSingletonHeaderDiagnosticCode
            || diagnostic.Code == MimeHeaderParser.InvalidUtf8HeaderDiagnosticCode
            || diagnostic.Code == MimeHeaderParser.EncodedWordInStructuredHeaderDiagnosticCode
            || diagnostic.Code == MimeValueParser.DuplicateSecurityParameterDiagnosticCode
            || diagnostic.Code == MimeValueParser.ParameterContinuationGapDiagnosticCode
            || diagnostic.Code == MimeValueParser.InvalidExtendedParameterDiagnosticCode
            || diagnostic.Code == MimeParser.MultipleHtmlBodyDiagnosticCode
            || diagnostic.Code == MimeParser.RelatedRootMissingDiagnosticCode
            || diagnostic.Code == MimeParser.RelatedRootNotHtmlDiagnosticCode
            || diagnostic.Code == MimeParser.RelatedRootTypeMismatchDiagnosticCode
            || diagnostic.Code == MimeParser.EmptyBoundaryDiagnosticCode
            || diagnostic.Code == MimeParser.BoundaryNotClosedDiagnosticCode);
        if (ambiguous != null) {
            throw new InvalidDataException(
                "MHTML content-safety inspection requires unambiguous MIME content headers and embedded resource identities. " +
                ambiguous.Code + ": " + ambiguous.Message);
        }
        return document;
    }

    private static async Task<OfficeContentSafetyReport> InspectDocumentContentSafetyAsync(
        MhtmlDocument document,
        OfficeContentSafetyOptions options,
        CancellationToken cancellationToken) {
        HtmlRenderOptions renderOptions = CreateContentSafetyRenderOptions(document, options);
        var part = new HtmlContentSafetyPackagePart("root", document.Html, "MHTML/Root", renderOptions);
        return await HtmlContentSafety.InspectPackagePartsAsync(
            "MHTML",
            new[] { part },
            options,
            cancellationToken).ConfigureAwait(false);
    }

    private static HtmlRenderOptions CreateContentSafetyRenderOptions(MhtmlDocument document, OfficeContentSafetyOptions options) {
        var renderOptions = new HtmlRenderOptions {
            MaxInputCharacters = options.MaxCharacters,
            MaxResourceBytes = Math.Min(options.MaxInputBytes, int.MaxValue),
            MaxTotalResourceBytes = options.MaxExpandedPackageBytes,
            MaxResourceCount = options.MaxPackageEntries,
            MaxResourceRequests = options.MaxPackageEntries
        };
        document.ConfigureRenderOptions(renderOptions, MhtmlRemoteResourcePolicy.CreateEmbeddedOnlyProfile());
        return renderOptions;
    }

    private static EmailReaderOptions IntersectReaderOptions(EmailReaderOptions source, OfficeContentSafetyOptions safety) =>
        new EmailReaderOptions(
            maxInputBytes: Math.Min(source.MaxInputBytes, safety.MaxInputBytes),
            maxHeaderBytes: source.MaxHeaderBytes,
            maxHeaderCount: source.MaxHeaderCount,
            maxPartCount: Math.Min(source.MaxPartCount, safety.MaxPackageEntries),
            maxMimeDepth: source.MaxMimeDepth,
            maxAttachmentBytes: Math.Min(source.MaxAttachmentBytes, safety.MaxInputBytes),
            maxTotalAttachmentBytes: Math.Min(source.MaxTotalAttachmentBytes, safety.MaxExpandedPackageBytes),
            maxNestedMessageDepth: source.MaxNestedMessageDepth,
            includeAttachmentContent: true,
            preserveRawSource: false,
            maxCompoundDirectoryEntries: source.MaxCompoundDirectoryEntries,
            maxMapiPropertyCount: source.MaxMapiPropertyCount,
            maxDecodedPropertyBytes: Math.Min(source.MaxDecodedPropertyBytes, safety.MaxExpandedPackageBytes),
            maxTnefAttributeCount: source.MaxTnefAttributeCount,
            maxAttachmentCount: Math.Min(source.MaxAttachmentCount, safety.MaxPackageEntries));
}
