namespace OfficeIMO.Email.Store;

/// <summary>Aggregate outcome for one merge source.</summary>
public sealed class EmailStoreMergeSourceReport {
    internal EmailStoreMergeSourceReport(string sourcePath, EmailStoreFormat format, int folderCount,
        int inspectedItems, int writtenItems, int duplicateItems, int skippedItems, int retryCount,
        bool completed) {
        SourcePath = sourcePath;
        Format = format;
        FolderCount = folderCount;
        InspectedItems = inspectedItems;
        WrittenItems = writtenItems;
        DuplicateItems = duplicateItems;
        SkippedItems = skippedItems;
        RetryCount = retryCount;
        Completed = completed;
    }

    /// <summary>Absolute source path supplied by the caller. It is not written into the PST.</summary>
    public string SourcePath { get; }
    /// <summary>Detected source format.</summary>
    public EmailStoreFormat Format { get; }
    /// <summary>Source folders considered.</summary>
    public int FolderCount { get; }
    /// <summary>Items inspected.</summary>
    public int InspectedItems { get; }
    /// <summary>Items written.</summary>
    public int WrittenItems { get; }
    /// <summary>Semantic duplicates omitted.</summary>
    public int DuplicateItems { get; }
    /// <summary>Items skipped.</summary>
    public int SkippedItems { get; }
    /// <summary>Transient I/O retries consumed.</summary>
    public int RetryCount { get; }
    /// <summary>Whether enumeration reached the end of this source.</summary>
    public bool Completed { get; }
}
