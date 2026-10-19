using System;
using System.Collections.Generic;

namespace OfficeIMO.Archives;

/// <summary>Finite resource and selection limits for one archive extraction.</summary>
public sealed class ArchiveExtractionOptions {
    /// <summary>Optional cumulative budget shared with related and nested archive operations.</summary>
    public ArchiveOperationBudget? OperationBudget { get; set; }
    /// <summary>Explicit code page for unflagged non-ASCII ZIP names. Null retains strict ASCII/UTF-8 admission;
    /// the UTF-8 flag always takes precedence. No host-locale inference occurs.</summary>
    public int? ZipLegacyCodePage { get; set; }

    /// <summary>Allowed formats detected from package bytes. Null permits every supported format.
    /// A disallowed format raises NotSupportedException before staging or payload parsing.</summary>
    public IReadOnlyCollection<ArchiveFormat>? AllowedFormats { get; set; }
    /// <summary>Maximum package size in bytes. Defaults to 512 MiB.</summary>
    public long MaxInputBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum number of declared archive members or MSI File rows.</summary>
    public int MaxEntryCount { get; set; } = 200000;
    /// <summary>Maximum distinct embedded MSI cabinets. Defaults to 256, including cabinets with empty or unselected files.</summary>
    public int MaxCabinetCount { get; set; } = 256;
    /// <summary>Maximum expanded size of any member. Defaults to 256 MiB.</summary>
    public long MaxEntryBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum aggregate expanded bytes. CAB folder decoding also observes this bound.</summary>
    public long MaxExpandedBytes { get; set; } = 1024L * 1024 * 1024;
    /// <summary>Maximum aggregate CAB decompression work bytes across all folders and embedded MSI cabinets.
    /// Defaults to 2 GiB, independently of expanded content. Each frame charges the greatest of 32 KiB,
    /// its complete CFDATA record size, and its expanded size plus any MSZIP dictionary replay.
    /// This bounds processing work rather than measuring exact CPU time.</summary>
    public long MaxDecompressionWorkBytes { get; set; } = 2L * 1024 * 1024 * 1024;
    /// <summary>Maximum aggregate MSI metadata stream bytes. Defaults to 64 MiB.</summary>
    public long MaxMetadataBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum rows per MSI metadata table.</summary>
    public int MaxTableRows { get; set; } = 200000;
    /// <summary>Maximum MSI string-pool entries, excluding the null reference. Defaults to 200,000, independently of table rows.</summary>
    public int MaxStringPoolEntries { get; set; } = 200000;
    /// <summary>Maximum output-path segments and structural MSI directory depth, including the root and '.' directories.</summary>
    public int MaxPathDepth { get; set; } = 64;
    /// <summary>Maximum characters in an authored relative output path.</summary>
    public int MaxPathLength { get; set; } = 1024;
    /// <summary>Optional file selection predicate receiving a normalized authored relative path.
    /// MSI aliases invoke it once with their deterministic logical target.
    /// All member names remain subject to safety, collision and payload-consistency checks.</summary>
    public Func<string, bool>? IncludeFile { get; set; }
}
