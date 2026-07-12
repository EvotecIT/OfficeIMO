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

    /// <summary>Number of deterministic structural recoveries.</summary>
    public int RepairCount => Diagnostics.Count(static diagnostic => diagnostic.WasRecovered);

    /// <summary>Number of detected issues intentionally left unchanged.</summary>
    public int DetectionOnlyCount => Diagnostics.Count(static diagnostic => !diagnostic.WasRecovered);

    internal PdfRepairReport Append(IEnumerable<PdfRepairDiagnostic> diagnostics) {
        PdfRepairDiagnostic[] appended = diagnostics.ToArray();
        if (appended.Length == 0) return this;
        return new PdfRepairReport(Diagnostics.Concat(appended).ToArray());
    }
}
