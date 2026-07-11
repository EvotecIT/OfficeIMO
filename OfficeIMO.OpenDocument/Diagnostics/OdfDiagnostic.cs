namespace OfficeIMO.OpenDocument;

/// <summary>Describes a package, XML, compatibility, or preservation finding.</summary>
public sealed class OdfDiagnostic {
    /// <summary>Creates a diagnostic.</summary>
    public OdfDiagnostic(string id, OdfDiagnosticSeverity severity, string message, string? partPath = null, int? lineNumber = null, int? linePosition = null) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        PartPath = partPath;
        LineNumber = lineNumber;
        LinePosition = linePosition;
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Id { get; }
    /// <summary>Diagnostic severity.</summary>
    public OdfDiagnosticSeverity Severity { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Package part associated with the finding.</summary>
    public string? PartPath { get; }
    /// <summary>One-based XML line number when available.</summary>
    public int? LineNumber { get; }
    /// <summary>One-based XML line position when available.</summary>
    public int? LinePosition { get; }
}
