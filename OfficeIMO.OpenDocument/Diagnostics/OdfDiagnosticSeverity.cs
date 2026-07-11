namespace OfficeIMO.OpenDocument;

/// <summary>Severity assigned to an OpenDocument diagnostic.</summary>
public enum OdfDiagnosticSeverity {
    /// <summary>Informational feature or preservation evidence.</summary>
    Info,
    /// <summary>Non-fatal compatibility or preservation concern.</summary>
    Warning,
    /// <summary>Invalid or unsafe document condition.</summary>
    Error
}
