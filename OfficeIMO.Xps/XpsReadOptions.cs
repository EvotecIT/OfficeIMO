namespace OfficeIMO.Xps;

/// <summary>Resource limits applied before materializing package parts and page XML.</summary>
public sealed class XpsReadOptions {
    /// <summary>Maximum compressed input bytes (default 128 MiB).</summary>
    public int MaximumInputBytes { get; set; } = 128 * 1024 * 1024;
    /// <summary>Maximum expanded bytes across all parts (default 256 MiB).</summary>
    public long MaximumExpandedBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum expanded bytes for one part (default 32 MiB).</summary>
    public int MaximumPartBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum ZIP entries.</summary>
    public int MaximumParts { get; set; } = 10000;
    /// <summary>Maximum page references across the document sequence.</summary>
    public int MaximumPages { get; set; } = 5000;
    /// <summary>Maximum XML element nesting depth.</summary>
    public int MaximumXmlDepth { get; set; } = 64;
    /// <summary>Returns an independent, validated copy of these package and XML limits.</summary>
    public XpsReadOptions Clone() => Snapshot();
    internal XpsReadOptions Snapshot() {
        if (MaximumInputBytes <= 0 || MaximumExpandedBytes <= 0 || MaximumPartBytes <= 0 ||
            MaximumParts <= 0 || MaximumPages <= 0 || MaximumXmlDepth <= 0)
            throw new ArgumentOutOfRangeException(nameof(XpsReadOptions), "All XPS limits must be positive.");
        return (XpsReadOptions)MemberwiseClone();
    }
}

/// <summary>The native fixed-page markup dialect, independent of the filename extension.</summary>
public enum XpsFormat {
    /// <summary>Microsoft XML Paper Specification 1.0.</summary>
    Xps,
    /// <summary>ECMA-388 OpenXPS.</summary>
    OpenXps
}
