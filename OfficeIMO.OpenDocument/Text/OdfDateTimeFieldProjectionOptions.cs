namespace OfficeIMO.OpenDocument;

/// <summary>Selects saved text, saved values, or an explicitly supplied date/time for field projection.</summary>
public enum OdfDateTimeFieldProjectionMode {
    /// <summary>Uses saved display text without evaluating native values or data styles.</summary>
    CachedText,
    /// <summary>Formats native saved date/time values through their data styles.</summary>
    StoredValues,
    /// <summary>Formats dynamic fields from one supplied timestamp; fixed fields retain their saved values.</summary>
    RefreshDynamic
}

/// <summary>Deterministic date/time field settings. Projection never reads the system clock.</summary>
public sealed class OdfDateTimeFieldProjectionOptions {
    /// <summary>Evaluation mode. Default: saved display text.</summary>
    public OdfDateTimeFieldProjectionMode Mode { get; set; }
    /// <summary>Required for dynamic refresh. Its civil date/time and offset are retained without local-time conversion.</summary>
    public DateTimeOffset? RefreshTimestamp { get; set; }
    /// <summary>Explicit fallback locale when a data style omits one. Empty selects invariant culture.</summary>
    public string DefaultCultureName { get; set; } = string.Empty;
    /// <summary>Creates independent settings.</summary>
    public OdfDateTimeFieldProjectionOptions Clone() => new() {
        Mode = Mode, RefreshTimestamp = RefreshTimestamp, DefaultCultureName = DefaultCultureName
    };
    internal OdfDateTimeFieldProjectionOptions Snapshot() {
        var copy = Clone();
        if (copy.Mode < OdfDateTimeFieldProjectionMode.CachedText || copy.Mode > OdfDateTimeFieldProjectionMode.RefreshDynamic)
            throw new ArgumentOutOfRangeException(nameof(Mode));
        if ((copy.Mode == OdfDateTimeFieldProjectionMode.RefreshDynamic) != copy.RefreshTimestamp.HasValue)
            throw new ArgumentException("Supply a timestamp only when refreshing dynamic date/time fields.");
        if (copy.DefaultCultureName == null) throw new ArgumentNullException(nameof(DefaultCultureName));
        _ = CultureInfo.GetCultureInfo(copy.DefaultCultureName);
        return copy;
    }
}
