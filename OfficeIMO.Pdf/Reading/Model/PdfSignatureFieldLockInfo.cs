namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata from a signature field /Lock dictionary.
/// </summary>
public sealed class PdfSignatureFieldLockInfo {
    internal PdfSignatureFieldLockInfo(string? action, IReadOnlyList<string> fields) {
        Action = action;
        Fields = fields;
    }

    /// <summary>Field-lock /Action name such as All, Include, or Exclude, when readable.</summary>
    public string? Action { get; }

    /// <summary>Field names listed by the lock dictionary /Fields array.</summary>
    public IReadOnlyList<string> Fields { get; }

    /// <summary>True when the lock applies to every form field.</summary>
    public bool LocksAllFields => string.Equals(Action, "All", StringComparison.Ordinal);

    /// <summary>True when the lock applies only to the listed fields.</summary>
    public bool LocksIncludedFields => string.Equals(Action, "Include", StringComparison.Ordinal);

    /// <summary>True when the lock applies to fields except the listed fields.</summary>
    public bool LocksAllExceptListedFields => string.Equals(Action, "Exclude", StringComparison.Ordinal);
}
