namespace PSWriteOffice.Models.Pdf;

/// <summary>The native PDF diagnostic report to return.</summary>
public enum OfficePdfDiagnosticType {
    /// <summary>Object structure, stream statistics and feature findings.</summary>
    Structure,
    /// <summary>Read and rewrite capabilities and blockers.</summary>
    Preflight,
    /// <summary>Incremental update and append-only mutation support.</summary>
    Incremental,
    /// <summary>Native optimization opportunities.</summary>
    Optimization
}
