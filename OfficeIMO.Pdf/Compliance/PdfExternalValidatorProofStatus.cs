namespace OfficeIMO.Pdf;

/// <summary>
/// Claimability status for one required external validator family.
/// </summary>
public enum PdfExternalValidatorProofStatus {
    /// <summary>No matching external validation result was supplied.</summary>
    Missing,

    /// <summary>A matching validation result was supplied, but the validator was not run.</summary>
    NotRun,

    /// <summary>A matching required validator accepted the requested profile and no matching failure was supplied.</summary>
    Passed,

    /// <summary>A matching required validator rejected the requested profile.</summary>
    Failed,

    /// <summary>A matching required validator could not complete because of configuration, runtime, or process failure.</summary>
    Error
}
