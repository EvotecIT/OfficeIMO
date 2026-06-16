namespace OfficeIMO.Html;

/// <summary>
/// Thrown when shared HTML DOM traversal exceeds configured safety limits.
/// </summary>
public sealed class HtmlDomLimitException : InvalidOperationException {
    /// <summary>
    /// Creates an HTML DOM traversal limit exception.
    /// </summary>
    /// <param name="code">Stable diagnostic code associated with the limit.</param>
    /// <param name="message">Human-readable message.</param>
    /// <param name="source">Configured limit source that was exceeded.</param>
    /// <param name="actual">Observed value.</param>
    /// <param name="limit">Configured limit.</param>
    public HtmlDomLimitException(string code, string message, string source, long actual, long limit)
        : base(message) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        LimitSource = source ?? throw new ArgumentNullException(nameof(source));
        Actual = actual;
        Limit = limit;
        Detail = "Actual=" + actual + "; Limit=" + limit;
    }

    /// <summary>Stable diagnostic code associated with the limit.</summary>
    public string Code { get; }

    /// <summary>Configured limit source that was exceeded.</summary>
    public string LimitSource { get; }

    /// <summary>Observed value.</summary>
    public long Actual { get; }

    /// <summary>Configured limit.</summary>
    public long Limit { get; }

    /// <summary>Formatted detail containing the observed and configured values.</summary>
    public string Detail { get; }
}
