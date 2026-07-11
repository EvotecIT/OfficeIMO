namespace OfficeIMO.Pdf;

/// <summary>Explicit, reproducible recoveries applied by lenient PDF parsing.</summary>
public sealed class PdfRepairReport {
    internal PdfRepairReport(IReadOnlyList<PdfRepairDiagnostic> diagnostics) {
        Diagnostics = diagnostics;
    }

    /// <summary>Structural recoveries in deterministic source order.</summary>
    public IReadOnlyList<PdfRepairDiagnostic> Diagnostics { get; }

    /// <summary>True when lenient parsing recovered at least one defect.</summary>
    public bool HasRepairs => Diagnostics.Count > 0;

    /// <summary>Number of structural recoveries.</summary>
    public int RepairCount => Diagnostics.Count;
}
