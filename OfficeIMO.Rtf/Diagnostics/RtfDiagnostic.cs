namespace OfficeIMO.Rtf.Diagnostics;

/// <summary>
/// Describes a parser, binder, or writer condition discovered while processing RTF.
/// </summary>
public sealed class RtfDiagnostic {
    /// <summary>
    /// Initializes a new diagnostic.
    /// </summary>
    public RtfDiagnostic(RtfDiagnosticSeverity severity, string code, string message, int position) {
        Severity = severity;
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Position = position;
    }

    /// <summary>Severity of the diagnostic.</summary>
    public RtfDiagnosticSeverity Severity { get; }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Zero-based input position associated with the diagnostic, when known.</summary>
    public int Position { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Severity} {Code} at {Position}: {Message}";
}
