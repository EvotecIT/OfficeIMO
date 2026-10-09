namespace OfficeIMO.Visio;

/// <summary>Controls bounded, read-only reconstruction of legacy binary Visio files.</summary>
public sealed class VisioLegacyBinaryImportOptions {
    /// <summary>Source bytes, compound streams, records, projected items and text budgets.</summary>
    public OfficeLegacyImportLimits Limits { get; set; } = new();

    /// <summary>Maximum cumulative bytes decoded from distinct native streams. Default: 128 MiB.</summary>
    public int MaxDecompressedBytes { get; set; } = 128 * 1024 * 1024;

    /// <summary>Maximum pointer and reconstructed shape nesting depth. Default: 64.</summary>
    public int MaxDepth { get; set; } = 64;

    internal VisioLegacyBinaryImportOptions Snapshot() {
        if (Limits == null) throw new ArgumentNullException(nameof(Limits));
        Limits.Validate();
        if (MaxDecompressedBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxDecompressedBytes));
        if (MaxDepth is < 1 or > 100) throw new ArgumentOutOfRangeException(nameof(MaxDepth), "Depth must be between 1 and 100.");
        return new VisioLegacyBinaryImportOptions {
            Limits = Limits.Clone(), MaxDecompressedBytes = MaxDecompressedBytes, MaxDepth = MaxDepth
        };
    }
}
