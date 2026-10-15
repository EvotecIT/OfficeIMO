namespace OfficeIMO.Reader;

/// <summary>A native admission policy was exhausted. This failure remains terminal across aggregate reads.</summary>
/// <remarks>Fractional ceilings retain their exact numeric values. Null values mean the native engine did not expose that measurement.</remarks>
public sealed class ReaderAdmissionLimitException : IOException {
    /// <summary>Creates an admission failure retaining the native engine's structured cause.</summary>
    public ReaderAdmissionLimitException(string limitName, double? maximum, double? actual, Exception innerException)
        : base($"Reader admission exceeds {limitName}.", innerException) {
        if (string.IsNullOrWhiteSpace(limitName)) throw new ArgumentException("A limit name is required.", nameof(limitName));
        if (maximum.HasValue && (double.IsNaN(maximum.Value) || double.IsInfinity(maximum.Value) || maximum.Value < 0))
            throw new ArgumentOutOfRangeException(nameof(maximum));
        if (actual.HasValue && (double.IsNaN(actual.Value) || actual.Value < 0))
            throw new ArgumentOutOfRangeException(nameof(actual));
        if (innerException == null) throw new ArgumentNullException(nameof(innerException));
        LimitName = limitName;
        Maximum = maximum;
        Actual = actual;
    }

    /// <summary>The exhausted native policy or measurement.</summary>
    public string LimitName { get; }
    /// <summary>The native admission ceiling, or null when it is not exposed.</summary>
    public double? Maximum { get; }
    /// <summary>The observed native measurement, or null when it is not exposed.</summary>
    public double? Actual { get; }
}
