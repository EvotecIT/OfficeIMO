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

    /// <summary>Maximum compressed input bytes, including unselected members. Defaults to 64 MiB. Nonseekable input is snapshotted first, within this limit; depending on the reader the snapshot stays in memory, bounded by Int32.MaxValue, or moves to a temporary file beyond 8 MiB.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum physical members, including directories and unselected files.</summary>
    public int MaxEntryCount { get; set; } = 10000;
    /// <summary>Maximum declared size of any member, selected or not. Defaults to 16 MiB. Selected in-memory members are also bounded by Int32.MaxValue.</summary>
    public long MaxEntryBytes { get; set; } = DefaultMaxEntryBytes;
    /// <summary>Maximum expanded data, including siblings in a selected solid folder. Defaults to 64 MiB.</summary>
    public long MaxExpandedBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum decompression work for 7z and Zstandard. Zstandard charges each block before decoding it: its input
    /// plus its largest possible output (the block maximum for a compressed block), at least 4 KiB per block. It also
    /// charges 4 KiB per skippable frame, the dictionary copied into every frame that uses one, and the history copied
    /// when it grows or when a long window slides. Defaults to 2 GiB.</summary>
    public long MaxDecompressionWorkBytes { get; set; } = 2L * 1024 * 1024 * 1024;
    /// <summary>Maximum Zstandard window. Defaults to 64 MiB, enough for zstd's default parameters at levels up to 19;
    /// levels 20 to 22 with <c>--ultra</c>, and <c>--long</c> or an explicit <c>windowLog</c>, can declare 128 MiB or more and
    /// need a higher limit. It applies to the window descriptor, to the declared
    /// content size of a single-segment frame, which is that frame's window, and to the content of a supplied dictionary,
    /// all before any history is allocated. Each frame starts with a buffer for its dictionary and two blocks,
    /// and the history grows towards the window only as output is produced.</summary>
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
