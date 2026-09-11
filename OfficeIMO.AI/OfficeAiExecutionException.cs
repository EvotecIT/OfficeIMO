namespace OfficeIMO.AI;

/// <summary>Safe, provider-neutral failure categories. Unknown failures never expose raw provider messages.</summary>
public enum OfficeAiExecutionFailure {
    /// <summary>No trustworthy classification was available.</summary>
    Unknown,
    /// <summary>Authentication is missing, rejected, or no longer usable.</summary>
    AuthenticationRequired,
    /// <summary>The authenticated account is not permitted to perform the request.</summary>
    AccessDenied,
    /// <summary>The endpoint or selected model could not be found.</summary>
    NotFound,
    /// <summary>The provider refused the request due to a usage or rate limit.</summary>
    RateLimited,
    /// <summary>The service or network connection is unavailable.</summary>
    Unavailable,
    /// <summary>The operation exceeded its time allowance.</summary>
    TimedOut,
    /// <summary>The provider rejected the request format or options.</summary>
    InvalidRequest
}

/// <summary>Carries only an allowlisted failure category across the provider boundary, without credentials or response text.</summary>
public sealed class OfficeAiExecutionException : Exception {
    /// <summary>Creates a content-free provider failure. Raw exceptions are deliberately not retained.</summary>
    public OfficeAiExecutionException(OfficeAiExecutionFailure failure) : base("Document provider execution failed.") {
        Failure = Enum.IsDefined(failure) ? failure : OfficeAiExecutionFailure.Unknown;
    }
    /// <summary>Safe category supplied by the adapter.</summary>
    public OfficeAiExecutionFailure Failure { get; }
    /// <summary>Stable non-content diagnostic code suitable for reports and localized UI mapping.</summary>
    public string DiagnosticCode => Failure switch {
        OfficeAiExecutionFailure.AuthenticationRequired => "provider-authentication-required",
        OfficeAiExecutionFailure.AccessDenied => "provider-access-denied",
        OfficeAiExecutionFailure.NotFound => "provider-not-found",
        OfficeAiExecutionFailure.RateLimited => "provider-rate-limited",
        OfficeAiExecutionFailure.Unavailable => "provider-unavailable",
        OfficeAiExecutionFailure.TimedOut => "provider-timed-out",
        OfficeAiExecutionFailure.InvalidRequest => "provider-invalid-request",
        _ => "provider-execution-failed"
    };
}
