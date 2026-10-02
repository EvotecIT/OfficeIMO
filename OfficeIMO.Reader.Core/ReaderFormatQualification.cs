namespace OfficeIMO.Reader;

/// <summary>The extraction contract declared for one registered format. Registration alone leaves it unqualified.</summary>
public enum ReaderFormatSupport {
    /// <summary>No explicit qualification contract has been supplied.</summary>
    Unqualified,
    /// <summary>Readable semantic projection with declared conversion and preservation boundaries.</summary>
    ReadConvert,
    /// <summary>Bounded recovery of selected content; complete semantic preservation is not claimed.</summary>
    SalvageRead,
    /// <summary>A deliberate unsupported boundary described by the qualification.</summary>
    IntentionalBoundary
}

/// <summary>An immutable per-extension extraction profile, supplied by the owning adapter or application.</summary>
/// <remarks>These are declarations with discoverable evidence, not a guarantee that every producer variant is supported.
/// Reader projection does not imply source-format writing, layout fidelity, macro execution or signature preservation.</remarks>
public sealed class ReaderFormatQualification {
    /// <summary>Creates a profile. Evidence references should identify the owning contract or reproducible fixture suite.</summary>
    public ReaderFormatQualification(string extension, string formatId, ReaderFormatSupport support = ReaderFormatSupport.Unqualified,
        string? profile = null, IEnumerable<string>? preservation = null, IEnumerable<string>? limitations = null,
        IEnumerable<string>? evidence = null) {
        if (string.IsNullOrWhiteSpace(extension)) throw new ArgumentException("An extension is required.", nameof(extension));
        if (string.IsNullOrWhiteSpace(formatId)) throw new ArgumentException("A format identifier is required.", nameof(formatId));
        if (!Enum.IsDefined(typeof(ReaderFormatSupport), support)) throw new ArgumentOutOfRangeException(nameof(support));
        Extension = extension.Trim().ToLowerInvariant();
        if (!Extension.StartsWith(".", StringComparison.Ordinal)) Extension = "." + Extension;
        FormatId = formatId.Trim(); Support = support; Profile = profile;
        Preservation = Snapshot(preservation); Limitations = Snapshot(limitations); Evidence = Snapshot(evidence);
    }
    /// <summary>Registered extension including its leading period.</summary>
    public string Extension { get; }
    /// <summary>Owning format identifier, reused from its catalog where available.</summary>
    public string FormatId { get; }
    /// <summary>Declared extraction maturity.</summary>
    public ReaderFormatSupport Support { get; }
    /// <summary>Qualified version or extraction profile.</summary>
    public string? Profile { get; }
    /// <summary>Semantic content retained in the projection.</summary>
    public IReadOnlyList<string> Preservation { get; }
    /// <summary>Known preservation, security, rendering or producer boundaries.</summary>
    public IReadOnlyList<string> Limitations { get; }
    /// <summary>Owning contracts or reproducible evidence references.</summary>
    public IReadOnlyList<string> Evidence { get; }
    private static IReadOnlyList<string> Snapshot(IEnumerable<string>? values) => Array.AsReadOnly(
        (values ?? Array.Empty<string>()).Select(value => value ?? throw new ArgumentException("Profile values cannot be null.")).ToArray());
}
