namespace OfficeIMO.Pdf;

/// <summary>Removes or quarantines active content and embedded payloads through a proven full rewrite.</summary>
internal static partial class PdfSanitizer {
    /// <summary>Returns the forbidden-content inventory that the supplied policy would remove.</summary>
    public static IReadOnlyList<PdfSanitizationFinding> Analyze(byte[] pdf, PdfSanitizationOptions? options = null) {
        return Analyze(pdf, options, readOptions: null);
    }

    internal static IReadOnlyList<PdfSanitizationFinding> Analyze(byte[] pdf, PdfSanitizationOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        var parsed = PdfSyntax.ParseObjects(pdf, readOptions);
        return Scan(parsed.Map, options ?? new PdfSanitizationOptions());
    }

    /// <summary>
    /// Produces a normalized PDF with forbidden actions, unsafe URI targets, rich media, and embedded payloads removed.
    /// Quarantine mode returns decoded attachments to the caller but never writes them to disk.
    /// </summary>
    public static PdfSanitizationResult Sanitize(byte[] pdf, PdfSanitizationOptions? options = null) {
        return Sanitize(pdf, options, readOptions: null);
    }

    internal static PdfSanitizationResult Sanitize(byte[] pdf, PdfSanitizationOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfSanitizationOptions policy = options ?? new PdfSanitizationOptions();
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.Sanitize, readOptions);
        IReadOnlyList<PdfSanitizationFinding> before = Analyze(pdf, policy, readOptions);
        IReadOnlyList<PdfExtractedAttachment> quarantined = policy.EmbeddedFiles == PdfEmbeddedFileSanitizationMode.Quarantine
            ? PdfAttachmentExtractor.ExtractAttachments(PdfReadDocument.Open(pdf, readOptions))
            : Array.Empty<PdfExtractedAttachment>();

        byte[] sanitized = PdfDocumentObjectGraphRewriter.Rewrite(
            pdf,
            sourceReadOptions: readOptions,
            outputEncryption: null,
            (objects, security) => {
                SanitizeObjectGraph(objects, policy);
                return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
                    ? security.InfoObjectNumber
                    : null;
            });
        IReadOnlyList<PdfSanitizationFinding> remaining = Analyze(sanitized, policy, readOptions: null);
        if (remaining.Count > 0) {
            throw new InvalidOperationException(
                "PDF sanitization post-save validation found " + remaining.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " forbidden item(s); the artifact was not returned.");
        }

        return new PdfSanitizationResult(sanitized, plan, before, remaining, quarantined);
    }

    /// <summary>Sanitizes a PDF from the current position of a readable stream.</summary>
    public static PdfSanitizationResult Sanitize(Stream stream, PdfSanitizationOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Sanitize(buffer.ToArray(), options);
    }

    /// <summary>Sanitizes a PDF file and returns the result without writing output automatically.</summary>
    public static PdfSanitizationResult Sanitize(string inputPath, PdfSanitizationOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return Sanitize(File.ReadAllBytes(inputPath), options);
    }
}
