namespace OfficeIMO.Pdf;

/// <summary>Lightweight structural validation result for one PDF signature value.</summary>
public sealed class PdfSignatureValidationResult {
    internal PdfSignatureValidationResult(
        PdfSignatureInfo signature,
        bool hasCompleteByteRangeShape,
        bool byteRangeSegmentsAreOrdered,
        bool byteRangeCoversEndOfFile,
        long? byteRangeCoveredBytes,
        long? byteRangeGapStart,
        long? byteRangeGapLength,
        IReadOnlyList<PdfSignatureValidationFinding> findings) {
        Signature = signature;
        HasCompleteByteRangeShape = hasCompleteByteRangeShape;
        ByteRangeSegmentsAreOrdered = byteRangeSegmentsAreOrdered;
        ByteRangeCoversEndOfFile = byteRangeCoversEndOfFile;
        ByteRangeCoveredBytes = byteRangeCoveredBytes;
        ByteRangeGapStart = byteRangeGapStart;
        ByteRangeGapLength = byteRangeGapLength;
        Findings = findings;
    }

    /// <summary>Signature metadata read from the PDF.</summary>
    public PdfSignatureInfo Signature { get; }

    /// <summary>True when /ByteRange contains the common four numeric values used by detached PDF signatures.</summary>
    public bool HasCompleteByteRangeShape { get; }

    /// <summary>True when the parsed byte ranges are non-negative, non-overlapping, and ordered.</summary>
    public bool ByteRangeSegmentsAreOrdered { get; }

    /// <summary>True when the final /ByteRange segment ends at the end of the file.</summary>
    public bool ByteRangeCoversEndOfFile { get; }

    /// <summary>Total bytes covered by the parsed /ByteRange values, when readable.</summary>
    public long? ByteRangeCoveredBytes { get; }

    /// <summary>Start of the unsigned gap between the first two /ByteRange segments, when readable.</summary>
    public long? ByteRangeGapStart { get; }

    /// <summary>Length of the unsigned gap between the first two /ByteRange segments, when readable.</summary>
    public long? ByteRangeGapLength { get; }

    /// <summary>Findings for this signature value.</summary>
    public IReadOnlyList<PdfSignatureValidationFinding> Findings { get; }

    /// <summary>True when this signature has no structural validation errors.</summary>
    public bool IsStructurallyValid => Findings.All(static finding => finding.Severity != PdfDiagnosticSeverity.Error);
}
