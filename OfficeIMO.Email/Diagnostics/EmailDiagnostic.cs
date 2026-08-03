namespace OfficeIMO.Email;

/// <summary>Structured diagnostic produced while reading or writing an email artifact.</summary>
public sealed class EmailDiagnostic {
    /// <summary>Creates a non-sensitive actionable diagnostic for a bounded parser stop.</summary>
    public static EmailDiagnostic FromLimit(EmailLimitExceededException exception, string operation,
        string? location = null) {
        if (exception == null) throw new ArgumentNullException(nameof(exception));
        return new EmailDiagnostic("EMAIL_LIMIT_EXCEEDED", exception.Message, EmailDiagnosticSeverity.Error,
            location, operation, null, exception.LimitName, exception.ActualValue, exception.MaximumValue,
            EmailDiagnosticDisposition.Stopped, EmailDataLossRisk.None,
            "Raise the named limit only after validating the source and the caller's resource budget.");
    }

    /// <summary>Creates a diagnostic.</summary>
    public EmailDiagnostic(string code, string message, EmailDiagnosticSeverity severity = EmailDiagnosticSeverity.Warning, string? location = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code is required.", nameof(code));
        Code = code;
        Message = message ?? string.Empty;
        Severity = severity;
        Location = location;
        Disposition = EmailDiagnosticDisposition.Observed;
        DataLossRisk = EmailDataLossRisk.None;
    }

    /// <summary>Creates an actionable diagnostic with machine-readable operation and recovery context.</summary>
    public EmailDiagnostic(string code, string message, EmailDiagnosticSeverity severity,
        string? location, string? operation, long? byteOffset, string? limitName,
        long? actualValue, long? maximumValue, EmailDiagnosticDisposition disposition,
        EmailDataLossRisk dataLossRisk, string? suggestedAction, bool isRetryable = false) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code is required.", nameof(code));
        Code = code;
        Message = message ?? string.Empty;
        Severity = severity;
        Location = location;
        Operation = operation;
        ByteOffset = byteOffset;
        LimitName = limitName;
        ActualValue = actualValue;
        MaximumValue = maximumValue;
        Disposition = disposition;
        DataLossRisk = dataLossRisk;
        SuggestedAction = suggestedAction;
        IsRetryable = isRetryable;
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public EmailDiagnosticSeverity Severity { get; }

    /// <summary>Logical source location.</summary>
    public string? Location { get; }

    /// <summary>Parser, writer, validation, or maintenance operation that emitted the diagnostic.</summary>
    public string? Operation { get; }
    /// <summary>Source byte offset when the format exposes a stable position.</summary>
    public long? ByteOffset { get; }
    /// <summary>Configured resource limit associated with the diagnostic.</summary>
    public string? LimitName { get; }
    /// <summary>Observed resource value.</summary>
    public long? ActualValue { get; }
    /// <summary>Configured maximum resource value.</summary>
    public long? MaximumValue { get; }
    /// <summary>How processing continued after the condition.</summary>
    public EmailDiagnosticDisposition Disposition { get; }
    /// <summary>Whether the condition can omit or alter user data.</summary>
    public EmailDataLossRisk DataLossRisk { get; }
    /// <summary>Safe next action for an operator or calling application.</summary>
    public string? SuggestedAction { get; }
    /// <summary>Whether retrying the same operation without changing input or policy can be useful.</summary>
    public bool IsRetryable { get; }
}
