namespace OfficeIMO.Pdf;

/// <summary>
/// Severity assigned to a converter warning produced while mapping source content into PDF output.
/// </summary>
public enum PdfConversionWarningSeverity {
    /// <summary>The converter recorded informational context that may be useful to callers.</summary>
    Information,

    /// <summary>The converter simplified, skipped, clipped, or otherwise degraded source content.</summary>
    Warning,

    /// <summary>The converter encountered a serious issue that callers may need to treat as a failed conversion.</summary>
    Error
}
