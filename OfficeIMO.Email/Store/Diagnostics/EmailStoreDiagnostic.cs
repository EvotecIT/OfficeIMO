namespace OfficeIMO.Email.Store;

/// <summary>Severity assigned to an email-store diagnostic.</summary>
public enum EmailStoreDiagnosticSeverity {
    /// <summary>Informational observation.</summary>
    Information = 0,
    /// <summary>Recoverable compatibility or fidelity warning.</summary>
    Warning = 1,
    /// <summary>Content could not be interpreted completely.</summary>
    Error = 2
}

/// <summary>Structured diagnostic emitted while reading a mailbox store.</summary>
public sealed class EmailStoreDiagnostic {
    /// <summary>Creates a diagnostic.</summary>
    public EmailStoreDiagnostic(string code, string message,
        EmailStoreDiagnosticSeverity severity = EmailStoreDiagnosticSeverity.Warning,
        string? location = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code is required.", nameof(code));
        Code = code;
        Message = message ?? string.Empty;
        Severity = severity;
        Location = location;
    }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable description.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public EmailStoreDiagnosticSeverity Severity { get; }

    /// <summary>Logical source location.</summary>
    public string? Location { get; }
}
