namespace OfficeIMO.Html;

/// <summary>
/// Thrown when RTF HTML conversion input exceeds a configured safety or resource limit.
/// </summary>
public sealed class RtfHtmlConversionLimitException : InvalidOperationException {
    /// <summary>
    /// Creates a conversion limit exception.
    /// </summary>
    /// <param name="code">Stable diagnostic code associated with the limit.</param>
    /// <param name="message">Human-readable message.</param>
    /// <param name="source">Configured limit or resource source that was exceeded.</param>
    /// <param name="actual">Observed value.</param>
    /// <param name="limit">Configured limit.</param>
    /// <param name="detail">Optional formatted detail.</param>
    public RtfHtmlConversionLimitException(string code, string message, string source, long actual, long limit, string? detail = null) : base(message) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        LimitSource = source ?? throw new ArgumentNullException(nameof(source));
        Actual = actual;
        Limit = limit;
        Detail = detail;
    }

    /// <summary>
    /// Stable diagnostic code associated with the limit.
    /// </summary>
    public string Code { get; }

    /// <summary>
    /// Configured limit or resource source that was exceeded.
    /// </summary>
    public string LimitSource { get; }

    /// <summary>
    /// Observed value.
    /// </summary>
    public long Actual { get; }

    /// <summary>
    /// Configured limit.
    /// </summary>
    public long Limit { get; }

    /// <summary>
    /// Optional formatted detail.
    /// </summary>
    public string? Detail { get; }
}
