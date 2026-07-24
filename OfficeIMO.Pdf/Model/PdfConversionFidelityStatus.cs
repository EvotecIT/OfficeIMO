namespace OfficeIMO.Pdf;

/// <summary>High-level fidelity outcome derived from structured conversion diagnostics.</summary>
public enum PdfConversionFidelityStatus {
    /// <summary>No lossy conversion diagnostics were reported.</summary>
    Faithful,
    /// <summary>The document was laid out successfully, but one or more font families were explicitly substituted.</summary>
    FaithfulWithSubstitutions,
    /// <summary>The converter reported an approximation, omission, unsupported layout behavior, or error beyond a font substitution.</summary>
    Degraded
}
