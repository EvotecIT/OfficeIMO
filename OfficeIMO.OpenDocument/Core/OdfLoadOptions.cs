namespace OfficeIMO.OpenDocument;

/// <summary>Controls bounded package and XML loading.</summary>
public sealed class OdfLoadOptions {
    /// <summary>Password used to decrypt an encrypted package.</summary>
    /// <remarks>
    /// The password is used only while loading and is not retained by the document. A password must be supplied
    /// again when an encrypted document is saved with encryption.
    /// </remarks>
    public string? Password { get; set; }

    /// <summary>Maximum source package size in bytes.</summary>
    public long MaxPackageBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Maximum number of ZIP entries.</summary>
    public int MaxEntries { get; set; } = 10000;

    /// <summary>Maximum uncompressed size of one entry.</summary>
    public long MaxEntryUncompressedBytes { get; set; } = 128L * 1024L * 1024L;

    /// <summary>Maximum aggregate uncompressed package size.</summary>
    public long MaxTotalUncompressedBytes { get; set; } = 512L * 1024L * 1024L;

    /// <summary>Maximum aggregate PBKDF2 iterations across all encrypted entries.</summary>
    /// <remarks>The default is 10,000,000 iterations. The complete manifest is preflighted before any key derivation.</remarks>
    public long MaxTotalKdfIterations { get; set; } = 10_000_000L;

    /// <summary>Maximum declared expansion ratio for a compressed entry.</summary>
    public double MaxCompressionRatio { get; set; } = 300d;

    /// <summary>Maximum archive path depth.</summary>
    public int MaxDepth { get; set; } = 32;

    /// <summary>Maximum characters allowed in one parsed XML part.</summary>
    public long MaxXmlCharacters { get; set; } = 64L * 1024L * 1024L;

    /// <summary>Maximum element nesting depth allowed in one parsed XML part.</summary>
    public int MaxXmlDepth { get; set; } = 256;

    internal OdfLoadOptions Normalize() {
        return new OdfLoadOptions {
            Password = Password,
            MaxPackageBytes = Math.Max(1L, MaxPackageBytes),
            MaxEntries = Math.Max(1, MaxEntries),
            MaxEntryUncompressedBytes = Math.Max(1L, MaxEntryUncompressedBytes),
            MaxTotalUncompressedBytes = Math.Max(1L, MaxTotalUncompressedBytes),
            MaxTotalKdfIterations = Math.Max(1L, MaxTotalKdfIterations),
            MaxCompressionRatio = MaxCompressionRatio <= 0d ? 1d : MaxCompressionRatio,
            MaxDepth = Math.Max(1, MaxDepth),
            MaxXmlCharacters = Math.Max(1L, MaxXmlCharacters),
            MaxXmlDepth = Math.Max(1, MaxXmlDepth)
        };
    }
}
