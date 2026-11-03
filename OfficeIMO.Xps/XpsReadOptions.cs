namespace OfficeIMO.Xps;

/// <summary>Options applied when loading XPS and OpenXPS packages. The defaults are bounded for untrusted input.</summary>
public sealed class XpsReadOptions {
    /// <summary>Resource budgets applied before materializing package parts and page XML.</summary>
    public XpsReadLimits Limits { get; set; } = new XpsReadLimits();
    /// <summary>Returns an independent, validated copy of these options.</summary>
    public XpsReadOptions Clone() => new XpsReadOptions { Limits = SnapshotLimits() };
    internal XpsReadLimits SnapshotLimits() => (Limits ?? throw new ArgumentNullException(nameof(Limits))).Snapshot();
}

/// <summary>Resource budgets applied while reading an XPS package. Every limit must be positive.</summary>
public sealed class XpsReadLimits {
    /// <summary>Maximum compressed input bytes (default 128 MiB).</summary>
    public int MaxInputBytes { get; set; } = 128 * 1024 * 1024;
    /// <summary>Maximum expanded bytes across all parts (default 256 MiB).</summary>
    public long MaxExpandedBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum expanded bytes for one part (default 32 MiB).</summary>
    public int MaxPartBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum ZIP entries (default 10,000).</summary>
    public int MaxParts { get; set; } = 10000;
    /// <summary>Maximum page references across the document sequence (default 5,000).</summary>
    public int MaxPages { get; set; } = 5000;
    /// <summary>Maximum XML element nesting depth (default 64).</summary>
    public int MaxDepth { get; set; } = 64;
    internal XpsReadLimits Snapshot() {
        if (MaxInputBytes <= 0 || MaxExpandedBytes <= 0 || MaxPartBytes <= 0 ||
            MaxParts <= 0 || MaxPages <= 0 || MaxDepth <= 0)
            throw new ArgumentOutOfRangeException(nameof(XpsReadLimits), "All XPS limits must be positive.");
        return (XpsReadLimits)MemberwiseClone();
    }
}

/// <summary>The native fixed-page markup dialect, independent of the filename extension.</summary>
public enum XpsFormat {
    /// <summary>Microsoft XML Paper Specification 1.0.</summary>
    Xps,
    /// <summary>ECMA-388 OpenXPS.</summary>
    OpenXps
}
