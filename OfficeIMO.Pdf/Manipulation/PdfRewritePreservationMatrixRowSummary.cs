namespace OfficeIMO.Pdf;

/// <summary>
/// JSON-friendly summary of one PDF rewrite-preservation matrix row.
/// </summary>
public sealed class PdfRewritePreservationMatrixRowSummary {
    internal PdfRewritePreservationMatrixRowSummary(PdfRewritePreservationMatrixEntry entry) {
        Id = entry.Id;
        Operation = entry.Operation;
        ExpectedClassification = entry.ExpectedClassification.ToString();
        ActualClassification = entry.ActualClassification.ToString();
        Passed = entry.Passed;
        SourceFeatures = entry.SourceFeatures.ToArray();
        PreservationSummary = entry.PreservationReport?.Summary;
        FailureType = entry.FailureType;
        FailureMessage = entry.FailureMessage;
        Issues = entry.PreservationReport?.Issues
            .Select(static issue => new PdfRewritePreservationIssueSummary(issue))
            .ToArray();
    }

    /// <summary>Stable scenario id.</summary>
    public string Id { get; }

    /// <summary>Operation that was tested.</summary>
    public string Operation { get; }

    /// <summary>Expected classification.</summary>
    public string ExpectedClassification { get; }

    /// <summary>Observed classification.</summary>
    public string ActualClassification { get; }

    /// <summary>True when expected and observed classifications match.</summary>
    public bool Passed { get; }

    /// <summary>Source fixture feature labels.</summary>
    public IReadOnlyList<string> SourceFeatures { get; }

    /// <summary>Preservation report summary when the rewrite completed.</summary>
    public string? PreservationSummary { get; }

    /// <summary>Exception type name when the operation did not complete.</summary>
    public string? FailureType { get; }

    /// <summary>Exception message when the operation did not complete.</summary>
    public string? FailureMessage { get; }

    /// <summary>Preservation issues when the rewrite completed.</summary>
    public IReadOnlyList<PdfRewritePreservationIssueSummary>? Issues { get; }
}
