using System;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>
/// Severity for diagnostics emitted while translating between RTF and Markdown.
/// </summary>
public enum RtfMarkdownDiagnosticSeverity {
    /// <summary>Informational conversion note.</summary>
    Info,
    /// <summary>The conversion completed, but a feature was simplified or approximated.</summary>
    Warning,
    /// <summary>A feature could not be represented in the target format.</summary>
    Error
}

/// <summary>
/// Describes a conversion decision, simplification, or unsupported feature.
/// </summary>
public sealed class RtfMarkdownConversionDiagnostic {
    public RtfMarkdownConversionDiagnostic(string code, RtfMarkdownDiagnosticSeverity severity, string message, string? source = null) {
        Code = string.IsNullOrWhiteSpace(code) ? throw new ArgumentException("Diagnostic code is required.", nameof(code)) : code;
        Severity = severity;
        Message = string.IsNullOrWhiteSpace(message) ? throw new ArgumentException("Diagnostic message is required.", nameof(message)) : message;
        Source = source;
    }

    public string Code { get; }

    public RtfMarkdownDiagnosticSeverity Severity { get; }

    public string Message { get; }

    public string? Source { get; }

    public override string ToString() {
        return Source is null
            ? $"{Severity} {Code}: {Message}"
            : $"{Severity} {Code} ({Source}): {Message}";
    }
}
