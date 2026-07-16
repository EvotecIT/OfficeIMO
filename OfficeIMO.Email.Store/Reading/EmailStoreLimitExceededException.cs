namespace OfficeIMO.Email.Store;

/// <summary>Thrown when an email store exceeds an explicitly configured safety limit.</summary>
public sealed class EmailStoreLimitExceededException : IOException {
    internal EmailStoreLimitExceededException(string limitName, long actual, long maximum)
        : base(string.Concat(limitName, " limit exceeded: ", actual.ToString(CultureInfo.InvariantCulture),
            " > ", maximum.ToString(CultureInfo.InvariantCulture), ".")) {
        LimitName = limitName;
        Actual = actual;
        Maximum = maximum;
    }

    /// <summary>Name of the exceeded option.</summary>
    public string LimitName { get; }

    /// <summary>Observed value.</summary>
    public long Actual { get; }

    /// <summary>Configured maximum.</summary>
    public long Maximum { get; }
}
