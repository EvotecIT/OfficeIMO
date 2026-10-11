using System;

namespace OfficeIMO.Archives;

/// <summary>Finite resource and selection limits for one archive extraction.</summary>
public sealed class ArchiveExtractionOptions {
    /// <summary>Maximum package size in bytes. Defaults to 512 MiB.</summary>
    public long MaxInputBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum number of declared archive members or MSI File rows.</summary>
    public int MaxEntryCount { get; set; } = 200000;
    /// <summary>Maximum expanded size of any member. Defaults to 256 MiB.</summary>
    public long MaxEntryBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum aggregate expanded bytes. CAB folder decoding also observes this bound.</summary>
    public long MaxExpandedBytes { get; set; } = 1024L * 1024 * 1024;
    /// <summary>Maximum aggregate MSI metadata stream bytes. Defaults to 64 MiB.</summary>
    public long MaxMetadataBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum rows per MSI metadata table.</summary>
    public int MaxTableRows { get; set; } = 200000;
    /// <summary>Maximum output-path segments and structural MSI directory depth, including the root and '.' directories.</summary>
    public int MaxPathDepth { get; set; } = 64;
    /// <summary>Maximum characters in an authored relative output path.</summary>
    public int MaxPathLength { get; set; } = 1024;
    /// <summary>Optional file selection predicate receiving a normalized authored relative path.
    /// All member names remain subject to safety and collision checks.</summary>
    public Func<string, bool>? IncludeFile { get; set; }
}
