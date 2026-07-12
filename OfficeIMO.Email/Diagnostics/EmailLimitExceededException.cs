namespace OfficeIMO.Email;

/// <summary>Thrown when an email artifact exceeds an explicitly configured resource limit.</summary>
public sealed class EmailLimitExceededException : IOException {
    /// <summary>Creates a limit exception.</summary>
    public EmailLimitExceededException(string limitName, long actualValue, long maximumValue)
        : base(string.Concat(limitName, " exceeded: ", actualValue.ToString(CultureInfo.InvariantCulture),
            " is greater than ", maximumValue.ToString(CultureInfo.InvariantCulture), ".")) {
        LimitName = limitName;
        ActualValue = actualValue;
        MaximumValue = maximumValue;
    }

    /// <summary>Name of the exceeded reader option.</summary>
    public string LimitName { get; }

    /// <summary>Observed value.</summary>
    public long ActualValue { get; }

    /// <summary>Configured maximum value.</summary>
    public long MaximumValue { get; }
}
