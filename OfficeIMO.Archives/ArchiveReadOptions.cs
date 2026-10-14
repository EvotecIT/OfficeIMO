using System;

namespace OfficeIMO.Archives;

/// <summary>Finite limits for verified, in-memory ZIP member reads.</summary>
public sealed class ArchiveReadOptions {
    /// <summary>Maximum compressed input bytes, including unselected members. Defaults to 64 MiB.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum physical members, including directories and unselected files.</summary>
    public int MaxEntryCount { get; set; } = 10000;
    /// <summary>Maximum declared size of any member, selected or not. Defaults to 16 MiB.</summary>
    public long MaxEntryBytes { get; set; } = 16L * 1024 * 1024;
    /// <summary>Maximum aggregate selected content retained in the result. Defaults to 64 MiB.</summary>
    public long MaxExpandedBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum member path segments.</summary>
    public int MaxPathDepth { get; set; } = 64;
    /// <summary>Maximum characters in an authored relative member path.</summary>
    public int MaxPathLength { get; set; } = 1024;
    /// <summary>Optional selector receiving each normalized relative file path once.
    /// All members still undergo metadata, path and collision validation.</summary>
    public Func<string, bool>? IncludeFile { get; set; }

    internal ArchiveLimits Limits() => new(new ArchiveExtractionOptions {
        MaxInputBytes = MaxInputBytes, MaxEntryCount = MaxEntryCount, MaxEntryBytes = MaxEntryBytes,
        MaxExpandedBytes = MaxExpandedBytes, MaxPathDepth = MaxPathDepth, MaxPathLength = MaxPathLength,
        IncludeFile = IncludeFile
    });
}
