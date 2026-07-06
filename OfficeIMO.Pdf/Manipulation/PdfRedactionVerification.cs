namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable post-redaction checks for removed and retained PDF text markers.
/// </summary>
public static partial class PdfRedactionVerification {
    /// <summary>
    /// Verifies a redacted PDF using the supplied redaction verification profile.
    /// </summary>
    public static PdfRedactionVerificationReport Verify(byte[] redactedPdf, PdfRedactionVerificationOptions options) {
        Guard.NotNull(redactedPdf, nameof(redactedPdf));
        Guard.NotNull(options, nameof(options));

        string extractedText = PdfReadDocument.Load(redactedPdf).ExtractText();
        string rawPdf = options.CheckRawPdfBytes ? PdfEncoding.Latin1GetString(redactedPdf) : string.Empty;
        var issues = new List<PdfRedactionVerificationIssue>();

        if (options.CheckDecodedPdfStreams &&
            options.FailOnUndecodablePdfStreams &&
            options.RemovedTextMarkers.Count > 0) {
            issues.AddRange(FindUndecodableStreamIssues(redactedPdf));
        }

        for (int i = 0; i < options.RemovedTextMarkers.Count; i++) {
            string marker = options.RemovedTextMarkers[i];
            if (ContainsOrdinal(extractedText, marker)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedTextMarker",
                    marker,
                    "Removed text marker remains extractable after redaction: " + marker));
            }

            if (options.CheckRawPdfBytes && ContainsOrdinal(rawPdf, marker)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedRawMarker",
                    marker,
                    "Removed text marker remains in raw rewritten PDF bytes: " + marker));
            }

            if (options.CheckEncodedPdfStrings && ContainsEncodedPdfMarker(redactedPdf, marker)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedEncodedMarker",
                    marker,
                    "Removed text marker remains in encoded rewritten PDF string bytes: " + marker));
            }

            if (options.CheckDecodedPdfStreams && ContainsDecodedStreamMarker(redactedPdf, marker)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RemovedDecodedStreamMarker",
                    marker,
                    "Removed text marker remains in a decoded rewritten PDF stream: " + marker));
            }
        }

        for (int i = 0; i < options.RetainedTextMarkers.Count; i++) {
            string marker = options.RetainedTextMarkers[i];
            if (!ContainsOrdinal(extractedText, marker)) {
                issues.Add(new PdfRedactionVerificationIssue(
                    "RetainedTextMarker",
                    marker,
                    "Expected retained text marker is not extractable after redaction: " + marker));
            }
        }

        return new PdfRedactionVerificationReport(extractedText, options.CheckRawPdfBytes, options.CheckEncodedPdfStrings, options.CheckDecodedPdfStreams, issues.AsReadOnly());
    }

    /// <summary>
    /// Verifies a redacted PDF and throws when removed text remains or retained text disappears.
    /// </summary>
    public static PdfRedactionVerificationReport AssertVerified(byte[] redactedPdf, PdfRedactionVerificationOptions options) {
        PdfRedactionVerificationReport report = Verify(redactedPdf, options);
        report.ThrowIfFailed();
        return report;
    }

    private static bool ContainsOrdinal(string text, string marker) {
        return !string.IsNullOrEmpty(marker) && text.Contains(marker);
    }
}
