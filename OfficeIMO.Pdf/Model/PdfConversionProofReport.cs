namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable proof snapshot for a source-document to PDF conversion result.
/// </summary>
public sealed class PdfConversionProofReport {
    internal PdfConversionProofReport(
        PdfDocumentInfo? documentInfo,
        PdfLogicalDocument? logicalDocument,
        string extractedText,
        IReadOnlyList<string> logicalSignals,
        long artifactByteCount,
        string artifactSha256,
        PdfConversionReportSummary warningSummary,
        IReadOnlyList<PdfConversionProofIssue> issues) {
        DocumentInfo = documentInfo;
        LogicalDocument = logicalDocument;
        ExtractedText = extractedText ?? string.Empty;
        LogicalSignals = logicalSignals;
        ArtifactByteCount = artifactByteCount;
        ArtifactSha256 = artifactSha256 ?? string.Empty;
        WarningSummary = warningSummary;
        Issues = issues;
    }

    /// <summary>Inspection snapshot for the generated PDF when it could be read.</summary>
    public PdfDocumentInfo? DocumentInfo { get; }

    /// <summary>Logical readback snapshot for the generated PDF when requested or needed for proof.</summary>
    public PdfLogicalDocument? LogicalDocument { get; }

    /// <summary>Text extracted from the generated PDF during proof.</summary>
    public string ExtractedText { get; }

    /// <summary>Stable logical signal names observed in the generated PDF.</summary>
    public IReadOnlyList<string> LogicalSignals { get; }

    /// <summary>Generated PDF artifact size in bytes when artifact proof was captured.</summary>
    public long ArtifactByteCount { get; }

    /// <summary>Lowercase SHA-256 hash for the generated PDF artifact when artifact proof was captured.</summary>
    public string ArtifactSha256 { get; }

    /// <summary>Grouped warning summary captured from the conversion report snapshot.</summary>
    public PdfConversionReportSummary WarningSummary { get; }

    /// <summary>Missing or failed proof items.</summary>
    public IReadOnlyList<PdfConversionProofIssue> Issues { get; }

    /// <summary>True when every requested conversion proof item was satisfied.</summary>
    public bool IsSatisfied => Issues.Count == 0;

    /// <summary>Human-readable summary suitable for tests, logs, wrappers, and proof packs.</summary>
    public string Summary {
        get {
            if (IsSatisfied) {
                return "PDF conversion proof checks passed.";
            }

            return "PDF conversion proof failed: " + string.Join("; ", Issues.Select(issue => issue.Message));
        }
    }

    /// <summary>Throws an InvalidOperationException when conversion proof checks failed.</summary>
    public void ThrowIfFailed() {
        if (!IsSatisfied) {
            throw new InvalidOperationException(Summary);
        }
    }
}
