namespace OfficeIMO.Email.AddressBook;

/// <summary>Raised when an OAB structure exceeds a configured safety limit.</summary>
public sealed class OfflineAddressBookLimitExceededException : IOException {
    /// <summary>Creates a limit exception.</summary>
    public OfflineAddressBookLimitExceededException(string limitName, long actualValue, long limitValue,
        string? location = null)
        : base(string.Concat(
            "Offline address book limit '", limitName, "' was exceeded: ",
            actualValue.ToString(CultureInfo.InvariantCulture), " > ",
            limitValue.ToString(CultureInfo.InvariantCulture), ".")) {
        if (string.IsNullOrWhiteSpace(limitName)) throw new ArgumentException("Limit name is required.", nameof(limitName));
        LimitName = limitName;
        ActualValue = actualValue;
        LimitValue = limitValue;
        Location = location;
    }

    /// <summary>Name of the exceeded option.</summary>
    public string LimitName { get; }
    /// <summary>Observed value.</summary>
    public long ActualValue { get; }
    /// <summary>Configured maximum.</summary>
    public long LimitValue { get; }
    /// <summary>Logical source location when available.</summary>
    public string? Location { get; }
}
