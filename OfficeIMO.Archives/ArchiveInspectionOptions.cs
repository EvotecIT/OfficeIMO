namespace OfficeIMO.Archives;

/// <summary>Finite limits for one passive archive inspection and all member reads from its snapshot.</summary>
public sealed class ArchiveInspectionOptions {
    /// <summary>Maximum input size buffered in memory. Defaults to 512 MiB; managed arrays also require an Int32 length.</summary>
    public long MaxInputBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum CAB file records or MSI File rows, including unsupported and empty members.</summary>
    public int MaxEntryCount { get; set; } = 200000;
    /// <summary>Maximum distinct embedded cabinets, including cabinets without mapped files.</summary>
    public int MaxCabinetCount { get; set; } = 256;
    /// <summary>Maximum declared size of a file member or bytes returned by one storage-stream read.</summary>
    public long MaxEntryBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum declared aggregate CAB folder/file expansion at inventory time, and cumulative
    /// decoded folder or raw storage-stream bytes across subsequent reads. Repeated reads consume this budget again.</summary>
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
