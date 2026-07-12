namespace OfficeIMO.Email;

/// <summary>Structured diagnostic produced while reading or writing an email artifact.</summary>
public sealed class EmailDiagnostic {
    /// <summary>Creates a diagnostic.</summary>
    public EmailDiagnostic(string code, string message, EmailDiagnosticSeverity severity = EmailDiagnosticSeverity.Warning, string? location = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code is required.", nameof(code));
        Code = code;
        Message = message ?? string.Empty;
        Severity = severity;
        Location = location;
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public EmailDiagnosticSeverity Severity { get; }

    /// <summary>Logical source location.</summary>
    public string? Location { get; }
}
