namespace OfficeIMO.Pdf;

/// <summary>
/// Classification used by rewrite-preservation matrix proof rows.
/// </summary>
public enum PdfRewritePreservationMatrixClassification {
    /// <summary>The rewrite completed and the preservation proof passed.</summary>
    RewriteSafe,

    /// <summary>The rewrite completed, but preservation proof found drift.</summary>
    PreservationFailed,

    /// <summary>The rewrite was intentionally blocked by OfficeIMO.Pdf safety checks.</summary>
    Blocked,

    /// <summary>The rewrite failed for an unexpected reason.</summary>
    OperationFailed
}
