namespace OfficeIMO.OneNote;

/// <summary>
/// Safety and compatibility limits applied while reading OneNote artifacts.
/// </summary>
public sealed class OneNoteReaderOptions {
    /// <summary>Default maximum input size: 512 MiB.</summary>
    public const long DefaultMaxInputBytes = 512L * 1024L * 1024L;

    /// <summary>Maximum number of file-node-list fragments followed from one file.</summary>
    public const int DefaultMaxFileNodeListFragments = 100_000;

    /// <summary>Maximum number of file nodes retained from one revision store.</summary>
    public const int DefaultMaxFileNodes = 2_000_000;

    /// <summary>Maximum number of transaction-log fragments followed from one revision store.</summary>
    public const int DefaultMaxTransactionLogFragments = 100_000;

    /// <summary>Maximum number of transaction entries inspected from one revision store.</summary>
    public const int DefaultMaxTransactionEntries = 4_000_000;

    /// <summary>Maximum number of object declarations retained from one file.</summary>
    public const int DefaultMaxObjects = 1_000_000;

    /// <summary>Maximum number of properties retained for one object.</summary>
    public const int DefaultMaxPropertiesPerObject = 65_536;

    /// <summary>Maximum nesting depth for child property sets.</summary>
    public const int DefaultMaxPropertySetDepth = 128;

    /// <summary>Maximum number of distinct page object-space references traversed from one section.</summary>
    public const int DefaultMaxPageGraphNodes = 100_000;

    /// <summary>Maximum nesting depth across conflict and version-history page relationships.</summary>
    public const int DefaultMaxPageRelationshipDepth = 128;

    /// <summary>Maximum bytes retained for one embedded asset.</summary>
    public const long DefaultMaxAssetBytes = 64L * 1024L * 1024L;

    /// <summary>Maximum bytes retained across embedded assets.</summary>
    public const long DefaultMaxTotalAssetBytes = 256L * 1024L * 1024L;

    /// <summary>Maximum MS-FSSHTTPB stream objects decoded from one package.</summary>
    public const int DefaultMaxStreamObjects = 2_000_000;

    /// <summary>Maximum nesting depth of compound MS-FSSHTTPB stream objects.</summary>
    public const int DefaultMaxStreamObjectDepth = 256;

    /// <summary>
    /// Maximum input size when the source length is available. Set to <see langword="null"/>
    /// only when the caller deliberately accepts unbounded source sizes.
    /// </summary>
    public long? MaxInputBytes { get; set; } = DefaultMaxInputBytes;

    /// <summary>Maximum file-node-list fragments followed while resolving a revision store.</summary>
    public int MaxFileNodeListFragments { get; set; } = DefaultMaxFileNodeListFragments;

    /// <summary>Maximum number of file nodes retained while following file-node lists.</summary>
    public int MaxFileNodes { get; set; } = DefaultMaxFileNodes;

    /// <summary>Maximum transaction-log fragments followed while resolving committed file-node counts.</summary>
    public int MaxTransactionLogFragments { get; set; } = DefaultMaxTransactionLogFragments;

    /// <summary>Maximum transaction entries inspected while resolving committed file-node counts.</summary>
    public int MaxTransactionEntries { get; set; } = DefaultMaxTransactionEntries;

    /// <summary>Maximum object declarations retained while resolving a revision store.</summary>
    public int MaxObjects { get; set; } = DefaultMaxObjects;

    /// <summary>Maximum properties retained for one object.</summary>
    public int MaxPropertiesPerObject { get; set; } = DefaultMaxPropertiesPerObject;

    /// <summary>Maximum recursive nesting depth for child property sets.</summary>
    public int MaxPropertySetDepth { get; set; } = DefaultMaxPropertySetDepth;

    /// <summary>Maximum distinct current, conflict, and version-history object-space references traversed from one section.</summary>
    public int MaxPageGraphNodes { get; set; } = DefaultMaxPageGraphNodes;

    /// <summary>Maximum nesting depth across conflict and version-history page relationships.</summary>
    public int MaxPageRelationshipDepth { get; set; } = DefaultMaxPageRelationshipDepth;

    /// <summary>Maximum bytes materialized for one image or embedded file.</summary>
    public long MaxAssetBytes { get; set; } = DefaultMaxAssetBytes;

    /// <summary>Maximum bytes materialized across all assets in one read operation.</summary>
    public long MaxTotalAssetBytes { get; set; } = DefaultMaxTotalAssetBytes;

    /// <summary>Maximum number of stream objects decoded from an alternative package-store file.</summary>
    public int MaxStreamObjects { get; set; } = DefaultMaxStreamObjects;

    /// <summary>Maximum nesting depth for compound stream objects in an alternative package-store file.</summary>
    public int MaxStreamObjectDepth { get; set; } = DefaultMaxStreamObjectDepth;

    /// <summary>
    /// When true, violations of required header values fail immediately. When false,
    /// recoverable mismatches are returned as diagnostics where safe continuation is possible.
    /// </summary>
    public bool StrictHeaderValidation { get; set; } = true;

    /// <summary>
    /// When true, transaction sentinel checksums are validated before their file-node counts are accepted.
    /// Set to false only for compatibility with a known producer that emits non-conforming checksums.
    /// </summary>
    public bool ValidateTransactionChecksums { get; set; } = true;

    /// <summary>
    /// When true, unsupported objects and properties retain their encoded bytes for later round-trip writing.
    /// </summary>
    public bool PreserveUnknownData { get; set; } = true;

    internal void Validate() {
        if (MaxInputBytes.HasValue && MaxInputBytes.Value < 1) {
            throw new ArgumentOutOfRangeException(nameof(MaxInputBytes), "MaxInputBytes must be greater than zero when specified.");
        }
        if (MaxFileNodeListFragments < 1) throw new ArgumentOutOfRangeException(nameof(MaxFileNodeListFragments));
        if (MaxFileNodes < 1) throw new ArgumentOutOfRangeException(nameof(MaxFileNodes));
        if (MaxTransactionLogFragments < 1) throw new ArgumentOutOfRangeException(nameof(MaxTransactionLogFragments));
        if (MaxTransactionEntries < 1) throw new ArgumentOutOfRangeException(nameof(MaxTransactionEntries));
        if (MaxObjects < 1) throw new ArgumentOutOfRangeException(nameof(MaxObjects));
        if (MaxPropertiesPerObject < 1) throw new ArgumentOutOfRangeException(nameof(MaxPropertiesPerObject));
        if (MaxPropertySetDepth < 1) throw new ArgumentOutOfRangeException(nameof(MaxPropertySetDepth));
        if (MaxPageGraphNodes < 1) throw new ArgumentOutOfRangeException(nameof(MaxPageGraphNodes));
        if (MaxPageRelationshipDepth < 1) throw new ArgumentOutOfRangeException(nameof(MaxPageRelationshipDepth));
        if (MaxAssetBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxAssetBytes));
        if (MaxTotalAssetBytes < MaxAssetBytes) {
            throw new ArgumentOutOfRangeException(nameof(MaxTotalAssetBytes), "MaxTotalAssetBytes must be at least MaxAssetBytes.");
        }
        if (MaxStreamObjects < 1) throw new ArgumentOutOfRangeException(nameof(MaxStreamObjects));
        if (MaxStreamObjectDepth < 1) throw new ArgumentOutOfRangeException(nameof(MaxStreamObjectDepth));
    }
}
