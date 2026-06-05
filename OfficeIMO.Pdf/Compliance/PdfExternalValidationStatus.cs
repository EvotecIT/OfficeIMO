namespace OfficeIMO.Pdf;

/// <summary>
/// Outcome reported by an external compliance validator.
/// </summary>
public enum PdfExternalValidationStatus {
    /// <summary>The validator was expected but was not available or not executed.</summary>
    NotRun,

    /// <summary>The validator completed and accepted the PDF for the requested profile.</summary>
    Passed,

    /// <summary>The validator completed and rejected the PDF for the requested profile.</summary>
    Failed,

    /// <summary>The validator could not complete because of configuration, timeout, or tool failure.</summary>
    Error
}
