namespace OfficeIMO.Pdf;

/// <summary>
/// One-call health, capability, optimization, signature, repair, and optional compliance report.
/// </summary>
public sealed class PdfAnalysisReport {
    internal PdfAnalysisReport(
        PdfDocumentInfo info,
        PdfDocumentPreflight preflight,
        PdfDiagnosticReport diagnostics,
        PdfOptimizationReport optimization,
        PdfSignatureValidationReport signatures,
        PdfAppendOnlyMutationReport appendOnlyMutation,
        PdfRepairReport repair,
        PdfComplianceReadinessReport? compliance) {
        Info = info;
        Preflight = preflight;
        Diagnostics = diagnostics;
        Optimization = optimization;
        Signatures = signatures;
        AppendOnlyMutation = appendOnlyMutation;
        Repair = repair;
        Compliance = compliance;
    }

    /// <summary>Document metadata, pages, catalog features, fields, annotations, and security state.</summary>
    public PdfDocumentInfo Info { get; }

    /// <summary>Read and rewrite capabilities with structured blockers.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Object, stream, font, and parser findings.</summary>
    public PdfDiagnosticReport Diagnostics { get; }

    /// <summary>Lossless optimization opportunities and estimated savings.</summary>
    public PdfOptimizationReport Optimization { get; }

    /// <summary>Structural signature validation results.</summary>
    public PdfSignatureValidationReport Signatures { get; }

    /// <summary>Append-only mutation capabilities.</summary>
    public PdfAppendOnlyMutationReport AppendOnlyMutation { get; }

    /// <summary>Parser repairs applied while opening the document.</summary>
    public PdfRepairReport Repair { get; }

    /// <summary>Optional profile-specific compliance readiness.</summary>
    public PdfComplianceReadinessReport? Compliance { get; }

    /// <summary>True when the document can be read.</summary>
    public bool CanRead => Preflight.CanRead;

    /// <summary>True when a full rewrite has no known blocker.</summary>
    public bool CanRewrite => Preflight.CanRewrite;

    /// <summary>True when the analysis found an error-level diagnostic.</summary>
    public bool HasErrors => Diagnostics.Findings.Any(
        static finding => finding.Severity == PdfDiagnosticSeverity.Error);

    /// <summary>True when the analysis found a warning-level diagnostic.</summary>
    public bool HasWarnings => Diagnostics.Findings.Any(
        static finding => finding.Severity == PdfDiagnosticSeverity.Warning);

    /// <summary>
    /// True when the PDF is readable and has no error-level diagnostic.
    /// Rewrite or compliance restrictions remain available through their dedicated reports.
    /// </summary>
    public bool IsHealthy => CanRead && !HasErrors;
}
