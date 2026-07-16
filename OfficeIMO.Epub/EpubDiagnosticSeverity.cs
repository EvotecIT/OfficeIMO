namespace OfficeIMO.Epub;

/// <summary>Severity of an EPUB diagnostic.</summary>
public enum EpubDiagnosticSeverity {
    /// <summary>Informational package characteristic.</summary>
    Info,

    /// <summary>Recoverable package or extraction problem.</summary>
    Warning,

    /// <summary>Fatal package or extraction problem.</summary>
    Error
}
