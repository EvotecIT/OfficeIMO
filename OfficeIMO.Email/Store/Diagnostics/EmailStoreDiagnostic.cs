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
    /// <summary>Creates a non-sensitive actionable diagnostic for a bounded Store stop or skip.</summary>
    public static EmailStoreDiagnostic FromLimit(EmailStoreLimitExceededException exception, string operation,
        string? location = null, EmailDiagnosticDisposition disposition = EmailDiagnosticDisposition.Stopped,
        string code = "EMAIL_STORE_LIMIT_EXCEEDED") {
        if (exception == null) throw new ArgumentNullException(nameof(exception));
        return new EmailStoreDiagnostic(code, exception.Message,
            EmailStoreDiagnosticSeverity.Warning, location, operation, null, exception.LimitName,
            exception.Actual, exception.Maximum, disposition, EmailDataLossRisk.Possible,
            "Review the source and explicitly increase the named Store limit before retrying.");
    }

    /// <summary>Creates a diagnostic.</summary>
    public EmailStoreDiagnostic(string code, string message,
        EmailStoreDiagnosticSeverity severity = EmailStoreDiagnosticSeverity.Warning,
        string? location = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code is required.", nameof(code));
        Code = code;
        Message = message ?? string.Empty;
        Severity = severity;
        Location = location;
        Disposition = EmailDiagnosticDisposition.Observed;
        DataLossRisk = EmailDataLossRisk.None;
    }

    /// <summary>Creates an actionable store diagnostic with machine-readable operation and recovery context.</summary>
    public EmailStoreDiagnostic(string code, string message, EmailStoreDiagnosticSeverity severity,
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

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable description.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public EmailStoreDiagnosticSeverity Severity { get; }

    /// <summary>Logical source location.</summary>
    public string? Location { get; }
    /// <summary>Store operation that emitted the diagnostic.</summary>
    public string? Operation { get; }
    /// <summary>Source byte offset when available.</summary>
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
    /// <summary>Whether retrying without changing source or policy can be useful.</summary>
    public bool IsRetryable { get; }
}
