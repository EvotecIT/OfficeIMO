namespace OfficeIMO.Pdf;

/// <summary>
/// Result for one scenario in a PDF rewrite-preservation proof matrix.
/// </summary>
public sealed class PdfRewritePreservationMatrixEntry {
    internal PdfRewritePreservationMatrixEntry(
        string id,
        string operation,
        PdfRewritePreservationMatrixClassification expectedClassification,
        PdfRewritePreservationMatrixClassification actualClassification,
        IReadOnlyList<string> sourceFeatures,
        PdfRewritePreservationReport? preservationReport,
        string? failureType,
        string? failureMessage) {
        Id = id;
        Operation = operation;
        ExpectedClassification = expectedClassification;
        ActualClassification = actualClassification;
        SourceFeatures = sourceFeatures;
        PreservationReport = preservationReport;
        FailureType = failureType;
        FailureMessage = failureMessage;
    }

    /// <summary>Stable scenario id.</summary>
    public string Id { get; }

    /// <summary>Operation that was tested.</summary>
    public string Operation { get; }

    /// <summary>Expected classification for this scenario.</summary>
    public PdfRewritePreservationMatrixClassification ExpectedClassification { get; }

    /// <summary>Observed classification for this scenario.</summary>
    public PdfRewritePreservationMatrixClassification ActualClassification { get; }

    /// <summary>Source fixture feature labels.</summary>
    public IReadOnlyList<string> SourceFeatures { get; }

    /// <summary>Preservation proof report when the rewrite completed.</summary>
    public PdfRewritePreservationReport? PreservationReport { get; }

    /// <summary>Exception type name when the operation did not complete.</summary>
    public string? FailureType { get; }

    /// <summary>Exception message when the operation did not complete.</summary>
    public string? FailureMessage { get; }

    /// <summary>True when the observed classification matches the expected classification.</summary>
    public bool Passed => ActualClassification == ExpectedClassification;

    /// <summary>Human-readable row summary for logs and proof files.</summary>
    public string Summary {
        get {
            if (PreservationReport is not null) {
                return Id + " " + ActualClassification + ": " + PreservationReport.Summary;
            }

            if (!string.IsNullOrEmpty(FailureMessage)) {
                return Id + " " + ActualClassification + ": " + FailureMessage;
            }

            return Id + " " + ActualClassification + ".";
        }
    }
}
