namespace OfficeIMO.Pdf;

/// <summary>
/// Classifies reusable PDF layout and visual fidelity diagnostics.
/// </summary>
public enum PdfLayoutDiagnosticKind {
    /// <summary>Content was clipped to fit the available page, frame, or shape bounds.</summary>
    ClippedContent,

    /// <summary>Content exceeded the available page, frame, or shape bounds.</summary>
    Overflow,

    /// <summary>Source geometry had to be adjusted before rendering.</summary>
    AdjustedGeometry,

    /// <summary>Source content was skipped because it could not be mapped safely.</summary>
    SkippedContent,

    /// <summary>Source content was rendered using a simplified approximation.</summary>
    SimplifiedContent
}
