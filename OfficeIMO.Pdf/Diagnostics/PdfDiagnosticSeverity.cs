namespace OfficeIMO.Pdf;

/// <summary>Severity for a PDF diagnostic or optimization finding.</summary>
public enum PdfDiagnosticSeverity {
    /// <summary>Informational finding.</summary>
    Info = 0,
    /// <summary>Potential issue or improvement opportunity.</summary>
    Warning = 1,
    /// <summary>Blocking or invalid state.</summary>
    Error = 2
}
