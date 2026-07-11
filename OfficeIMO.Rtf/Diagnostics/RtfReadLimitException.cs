namespace OfficeIMO.Rtf;

/// <summary>
/// Thrown when an RTF input exceeds a configured resource limit.
/// </summary>
public sealed class RtfReadLimitException : InvalidOperationException {
    /// <summary>Initializes a new resource-limit exception.</summary>
    public RtfReadLimitException(string code, string message, string limitSource, long actual, long limit, int position = -1)
        : base(message) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        LimitSource = limitSource ?? throw new ArgumentNullException(nameof(limitSource));
        Actual = actual;
        Limit = limit;
        Position = position;
    }

    /// <summary>Stable machine-readable limit code.</summary>
    public string Code { get; }

    /// <summary>Name of the <see cref="RtfReadOptions"/> property that supplied the limit.</summary>
    public string LimitSource { get; }

    /// <summary>Observed or declared value that exceeded the limit.</summary>
    public long Actual { get; }

    /// <summary>Configured limit.</summary>
    public long Limit { get; }

    /// <summary>Zero-based source position, or -1 when unavailable.</summary>
    public int Position { get; }
}
