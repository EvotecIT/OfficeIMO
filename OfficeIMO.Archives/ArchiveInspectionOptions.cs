namespace OfficeIMO.Archives;

/// <summary>Finite limits for one passive archive inspection and all member reads from its snapshot.</summary>
public sealed class ArchiveInspectionOptions {
    /// <summary>Optional cumulative budget shared with related and nested archive operations.</summary>
    public ArchiveOperationBudget? OperationBudget { get; set; }
    /// <summary>Explicit code page for unflagged non-ASCII ZIP names. Null retains strict ASCII/UTF-8 admission;
    /// the UTF-8 flag always takes precedence. No host-locale inference occurs.</summary>
    public int? ZipLegacyCodePage { get; set; }

    /// <summary>Maximum input snapshot size. Defaults to 512 MiB. ZIP snapshots spill above 64 MiB; CAB/MSI arrays also require an Int32 length.</summary>
    public long MaxInputBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum ZIP/CAB file records or MSI File rows, including unsupported and empty members.</summary>
    public int MaxEntryCount { get; set; } = 200000;
    /// <summary>Maximum distinct embedded cabinets, including cabinets without mapped files.</summary>
    public int MaxCabinetCount { get; set; } = 256;
    /// <summary>Maximum declared size of a file member or bytes returned by one storage-stream read.</summary>
    public long MaxEntryBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum decoded bytes staged privately by one CopyTo operation, across ZIP, CAB and MSI.
    /// Defaults to 256 MiB; zero permits only empty members. Checked before creating a spool or charging reads.
    /// Excludes input snapshots and filesystem extraction. This is a disk bound, not a process-memory bound.</summary>
    public long MaxTemporaryBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum declared aggregate CAB folder/file expansion at inventory time, and cumulative
    /// decoded folder or raw storage-stream bytes across subsequent reads. A grouped read charges each selected folder once;
    /// reading it in another call consumes this budget again.</summary>
    public long MaxExpandedBytes { get; set; } = 1024L * 1024 * 1024;
    /// <summary>Maximum cumulative CAB frame-processing work. Defaults to 2 GiB and uses the same frame charges as extraction.</summary>
    public long MaxDecompressionWorkBytes { get; set; } = 2L * 1024 * 1024 * 1024;
    /// <summary>Maximum MSI metadata bytes read while building the inventory.</summary>
    public long MaxMetadataBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum rows per MSI metadata table.</summary>
    public int MaxTableRows { get; set; } = 200000;
    /// <summary>Maximum entries in the MSI string pool, excluding the null reference.</summary>
    public int MaxStringPoolEntries { get; set; } = 200000;
    /// <summary>Maximum CAB folder, file and CFDATA records visited while building the inventory, summed across all cabinets.
    /// Also bounds the MSI compound directory slots. Defaults to 1,000,000.</summary>
    public int MaxDirectoryRecords { get; set; } = 1000000;
    /// <summary>Maximum structural depth of authored MSI directories.</summary>
    public int MaxPathDepth { get; set; } = 64;
    /// <summary>Maximum characters in an assembled authored MSI target path.</summary>
    public int MaxPathLength { get; set; } = 1024;
}
