namespace OfficeIMO.Pdf;

/// <summary>
/// Result of verifying that redacted text is removed and expected surrounding content remains readable.
/// </summary>
public sealed class PdfRedactionVerificationReport {
    internal PdfRedactionVerificationReport(
        string extractedText,
        bool rawPdfBytesChecked,
        bool encodedPdfStringsChecked,
        bool decodedPdfStreamsChecked,
        IReadOnlyList<PdfRedactionVerificationIssue> issues) {
        ExtractedText = extractedText;
        RawPdfBytesChecked = rawPdfBytesChecked;
        EncodedPdfStringsChecked = encodedPdfStringsChecked;
        DecodedPdfStreamsChecked = decodedPdfStreamsChecked;
        Issues = issues;
    }

    /// <summary>Extracted text from the redacted PDF used for marker checks.</summary>
    public string ExtractedText { get; }

    /// <summary>True when raw rewritten PDF bytes were also searched for removed markers.</summary>
    public bool RawPdfBytesChecked { get; }

    /// <summary>True when common PDF string byte encodings and hex strings were searched for removed markers.</summary>
    public bool EncodedPdfStringsChecked { get; }

    /// <summary>True when decoded PDF stream content was searched for removed markers.</summary>
    public bool DecodedPdfStreamsChecked { get; }

    /// <summary>Verification issues found in the redacted PDF.</summary>
    public IReadOnlyList<PdfRedactionVerificationIssue> Issues { get; }

    /// <summary>True when all configured redaction checks passed.</summary>
    public bool IsVerified => Issues.Count == 0;

    /// <summary>Human-readable summary suitable for logs, tests, and wrappers.</summary>
    public string Summary {
        get {
            if (IsVerified) {
                return "PDF redaction verification checks passed.";
            }

            return "PDF redaction verification failed: " + string.Join("; ", Issues.Select(issue => issue.Message));
        }
    }

    /// <summary>Throws an InvalidOperationException when verification checks found issues.</summary>
    public void ThrowIfFailed() {
        if (!IsVerified) {
            throw new InvalidOperationException(Summary);
        }
    }
}
