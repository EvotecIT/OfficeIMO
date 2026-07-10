namespace OfficeIMO.AsciiDoc;

/// <summary>Severity of an AsciiDoc parser or writer diagnostic.</summary>
public enum AsciiDocDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Info = 0,
    /// <summary>Recoverable condition that may affect semantics.</summary>
    Warning = 1,
    /// <summary>Invalid or incomplete input recovered by preserving source.</summary>
    Error = 2
}
/// <summary>A source-located AsciiDoc parser or writer diagnostic.</summary>
public sealed class AsciiDocDiagnostic {
    internal AsciiDocDiagnostic(string code, AsciiDocDiagnosticSeverity severity, string message, AsciiDocSourceSpan span) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Span = span;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Diagnostic severity.</summary>
    public AsciiDocDiagnosticSeverity Severity { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>Relevant source range.</summary>
    public AsciiDocSourceSpan Span { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Code} {Severity}: {Message} ({Span})";
}
