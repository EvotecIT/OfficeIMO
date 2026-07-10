namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Severity of an AsciiDoc-to-Markdown conversion diagnostic.</summary>
public enum AsciiDocMarkdownDiagnosticSeverity {
    /// <summary>Informational conversion note.</summary>
    Info = 0,
    /// <summary>Potentially lossy conversion.</summary>
    Warning = 1
}
/// <summary>Outcome assigned to an AsciiDoc source feature during Markdown conversion.</summary>
public enum AsciiDocMarkdownConversionOutcome {
    /// <summary>Mapped to an equivalent Markdown semantic node.</summary>
    Converted = 0,
    /// <summary>Mapped to a simpler Markdown representation.</summary>
    Simplified,
    /// <summary>Original AsciiDoc source retained as visible fenced content.</summary>
    SourceFallback,
    /// <summary>Source feature intentionally omitted.</summary>
    Omitted
}

/// <summary>Source-located conversion diagnostic.</summary>
public sealed class AsciiDocMarkdownConversionDiagnostic {
    internal AsciiDocMarkdownConversionDiagnostic(
        string code,
        AsciiDocMarkdownDiagnosticSeverity severity,
        AsciiDocMarkdownConversionOutcome outcome,
        string feature,
        string message,
        AsciiDocSourceSpan sourceSpan) {
        Code = code;
        Severity = severity;
        Outcome = outcome;
        Feature = feature;
        Message = message;
        SourceSpan = sourceSpan;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Diagnostic severity.</summary>
    public AsciiDocMarkdownDiagnosticSeverity Severity { get; }

    /// <summary>Conversion outcome.</summary>
    public AsciiDocMarkdownConversionOutcome Outcome { get; }

    /// <summary>Source feature name.</summary>
    public string Feature { get; }

    /// <summary>Human-readable conversion explanation.</summary>
    public string Message { get; }

    /// <summary>Original AsciiDoc source range.</summary>
    public AsciiDocSourceSpan SourceSpan { get; }
}
