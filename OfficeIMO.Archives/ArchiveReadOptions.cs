namespace OfficeIMO.Archives;

/// <summary>Finite limits for verified in-memory archive member reads.</summary>
public sealed class ArchiveReadOptions {
    internal const int DefaultMaxEntryBytes = 16 * 1024 * 1024;
    internal const long DefaultMaxWindowBytes = 64L * 1024 * 1024;
    /// <summary>Optional cumulative budget shared with related and nested archive operations.</summary>
    public ArchiveOperationBudget? OperationBudget { get; set; }
    /// <summary>Explicit code page for unflagged non-ASCII ZIP names. Null retains strict ASCII/UTF-8 admission;
    /// the UTF-8 flag always takes precedence. No host-locale inference occurs.</summary>
    public int? ZipLegacyCodePage { get; set; }

    /// <summary>Maximum compressed input bytes, including unselected members. Defaults to 64 MiB. Stream snapshots are also bounded by Int32.MaxValue.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum physical members, including directories and unselected files.</summary>
    public int MaxEntryCount { get; set; } = 10000;
    /// <summary>Maximum declared size of any member, selected or not. Defaults to 16 MiB. Selected in-memory members are also bounded by Int32.MaxValue.</summary>
    public long MaxEntryBytes { get; set; } = DefaultMaxEntryBytes;
    /// <summary>Maximum expanded data, including siblings in a selected solid folder. Defaults to 64 MiB.</summary>
    public long MaxExpandedBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum decompression work for 7z. Defaults to 2 GiB.</summary>
    public long MaxDecompressionWorkBytes { get; set; } = 2L * 1024 * 1024 * 1024;
    /// <summary>Maximum history window or dictionary a compressed stream may declare. Defaults to 64 MiB, which covers the
    /// largest standard presets. Decoders allocate history only as output is produced. Single-segment frames keep their whole content as the window and are bounded by the
    /// output limit instead.</summary>
    public long MaxWindowBytes { get; set; } = DefaultMaxWindowBytes;
    /// <summary>Maximum member path segments.</summary>
    public int MaxPathDepth { get; set; } = 64;
    /// <summary>Maximum characters in an authored relative member path.</summary>
    public int MaxPathLength { get; set; } = 1024;
    /// <summary>Optional selector receiving each normalized relative file path once.
    /// All members still undergo metadata, path and collision validation.</summary>
    public Func<string, bool>? IncludeFile { get; set; }

    internal ArchiveLimits Limits() => new(new ArchiveExtractionOptions {
        OperationBudget = OperationBudget, ZipLegacyCodePage = ZipLegacyCodePage, MaxInputBytes = MaxInputBytes, MaxEntryCount = MaxEntryCount, MaxEntryBytes = MaxEntryBytes,
        MaxExpandedBytes = MaxExpandedBytes, MaxDecompressionWorkBytes = MaxDecompressionWorkBytes, MaxPathDepth = MaxPathDepth, MaxPathLength = MaxPathLength,
        IncludeFile = IncludeFile
    });
}
