namespace OfficeIMO.Epub;

/// <summary>Describes a deterministic EPUB package or extraction diagnostic.</summary>
public sealed class EpubDiagnostic {
    /// <summary>Stable machine-readable diagnostic code.</summary>
    public string Code { get; internal set; } = string.Empty;

    /// <summary>Diagnostic severity.</summary>
    public EpubDiagnosticSeverity Severity { get; internal set; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; internal set; } = string.Empty;

    /// <summary>Normalized archive path associated with the diagnostic, when known.</summary>
    public string? Path { get; internal set; }
}
