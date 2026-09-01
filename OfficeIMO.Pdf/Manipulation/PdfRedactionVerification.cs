namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable post-redaction checks for removed and retained PDF text markers.
/// </summary>
internal static partial class PdfRedactionVerification {
    /// <summary>
    /// Verifies a redacted PDF using the supplied redaction verification profile.
    /// </summary>
    public static PdfRedactionVerificationReport Verify(
        byte[] redactedPdf,
        PdfRedactionVerificationOptions options,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(redactedPdf, nameof(redactedPdf));
        Guard.NotNull(options, nameof(options));

        PdfLoadOptions effectiveReadOptions = PdfLoadOptions.Resolve(readOptions);
        string extractedText = PdfReadDocument.Open(redactedPdf, effectiveReadOptions).ExtractText();
        string rawPdf = options.CheckRawPdfBytes ? PdfEncoding.Latin1GetString(redactedPdf) : string.Empty;
        var issues = new List<PdfRedactionVerificationIssue>();
        var externalResults = new List<PdfRedactionExternalValidationResult>();

        if (options.CheckDecodedPdfStreams &&
            options.FailOnUndecodablePdfStreams &&
            options.RemovedTextMarkers.Count > 0) {
            issues.AddRange(FindUndecodableStreamIssues(redactedPdf, effectiveReadOptions));
        }

        for (int i = 0; i < options.RemovedTextMarkers.Count; i++) {
            string marker = options.RemovedTextMarkers[i];
            if (ContainsMarker(extractedText, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedTextMarker",
                    marker,
                    "Removed text marker remains extractable after redaction: " + marker));
            }

            if (options.CheckRawPdfBytes && ContainsMarker(rawPdf, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedRawMarker",
                    marker,
                    "Removed text marker remains in raw rewritten PDF bytes: " + marker));
            }

            if (options.CheckEncodedPdfStrings && ContainsEncodedPdfMarker(redactedPdf, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedEncodedMarker",
                    marker,
                    "Removed text marker remains in encoded rewritten PDF string bytes: " + marker));
            }

            if (options.CheckDecodedPdfStreams && ContainsDecodedStreamMarker(redactedPdf, marker, options.MatchCase, effectiveReadOptions)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedDecodedStreamMarker",
                    marker,
                    "Removed text marker remains in a decoded rewritten PDF stream: " + marker));
            }
        }

        for (int i = 0; i < options.RetainedTextMarkers.Count; i++) {
            string marker = options.RetainedTextMarkers[i];
            if (!ContainsMarker(extractedText, marker, options.MatchCase)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RetainedTextMarker",
                    marker,
                    "Expected retained text marker is not extractable after redaction: " + marker));
            }
        }

        if (options.CheckManagedRendering) {
            IReadOnlyList<PdfPageRenderResult> renders = PdfPageImageRenderer.RenderPages(redactedPdf, options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg, ContinueOnError = true }, readOptions: effectiveReadOptions);
            for (int i = 0; i < renders.Count; i++) if (!renders[i].Succeeded) issues.Add(new PdfRedactionVerificationIssue("ManagedRendering", renders[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), "Managed rendering failed for redacted page " + renders[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ": " + string.Join("; ", renders[i].Diagnostics)));
        }

        for (int i = 0; i < options.ExternalValidators.Count; i++) {
            PdfRedactionExternalValidationResult result = options.ExternalValidators[i].Validate((byte[])redactedPdf.Clone()); externalResults.Add(result);
            if (!result.IsValid) issues.Add(new PdfRedactionVerificationIssue("ExternalValidation", result.ValidatorName, "External redaction validation failed for " + result.ValidatorName + (string.IsNullOrWhiteSpace(result.Diagnostic) ? "." : ": " + result.Diagnostic)));
        }

        return new PdfRedactionVerificationReport(extractedText, options.CheckRawPdfBytes, options.CheckEncodedPdfStrings, options.CheckDecodedPdfStreams, options.CheckManagedRendering, externalResults.AsReadOnly(), issues.AsReadOnly());
    }

    /// <summary>
    /// Verifies a redacted PDF and throws when removed text remains or retained text disappears.
    /// </summary>
    public static PdfRedactionVerificationReport AssertVerified(
        byte[] redactedPdf,
        PdfRedactionVerificationOptions options,
        PdfLoadOptions? readOptions = null) {
        PdfRedactionVerificationReport report = Verify(redactedPdf, options, readOptions);
        report.ThrowIfFailed();
        return report;
    }

    private static bool ContainsMarker(string text, string marker, bool matchCase) {
        if (string.IsNullOrEmpty(marker)) {
            return false;
        }

#if NETSTANDARD2_0 || NETFRAMEWORK
        return text.IndexOf(
            marker,
            matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) >= 0;
#else
        return text.Contains(
            marker,
            matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
#endif
    }
}
