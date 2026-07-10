namespace OfficeIMO.Latex;

/// <summary>LaTeX parser diagnostic severity.</summary>
public enum LatexDiagnosticSeverity {
    /// <summary>Informational.</summary>
    Info = 0,
    /// <summary>Recoverable warning.</summary>
    Warning,
    /// <summary>Malformed structure.</summary>
    Error
}

/// <summary>Source-located LaTeX parser or recovery diagnostic.</summary>
public sealed class LatexDiagnostic {
    internal LatexDiagnostic(string code, LatexDiagnosticSeverity severity, string message, LatexSourceSpan span) {
        Code = code;
        Severity = severity;
        Message = message;
        Span = span;
    }

    /// <summary>Stable code.</summary>
    public string Code { get; }
    /// <summary>Severity.</summary>
    public LatexDiagnosticSeverity Severity { get; }
    /// <summary>Message.</summary>
    public string Message { get; }
    /// <summary>Source span.</summary>
    public LatexSourceSpan Span { get; }
}
