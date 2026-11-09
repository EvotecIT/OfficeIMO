namespace OfficeIMO;

/// <summary>An owning inspection engine reached an explicit admission ceiling. This is a resource failure, not malformed input.</summary>
public sealed class OfficeInspectionResourceLimitException : IOException {
    /// <summary>Creates a limit failure retaining its canonical owner and optional underlying bounded-engine failure.</summary>
    public OfficeInspectionResourceLimitException(string owner, string limitName, long maximum, Exception? innerException = null)
        : base($"{owner} inspection exceeds {limitName} ({maximum}).", innerException) {
        ArgumentException.ThrowIfNullOrWhiteSpace(owner);
        ArgumentException.ThrowIfNullOrWhiteSpace(limitName);
        ArgumentOutOfRangeException.ThrowIfNegative(maximum);
        Owner = owner; LimitName = limitName; Maximum = maximum;
    }
    /// <summary>The canonical bounded inspection family, such as content-safety or text-integrity.</summary>
    public string Owner { get; }
    /// <summary>The exceeded native limit property or fixed admission ceiling.</summary>
    public string LimitName { get; }
    /// <summary>The maximum admitted value.</summary>
    public long Maximum { get; }
}
