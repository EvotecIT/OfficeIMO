using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

/// <summary>Controls bounded, dependency-free merging of multiple stores into one Unicode PST.</summary>
public sealed class EmailStorePstMergeOptions {
    /// <summary>Creates merge options.</summary>
    public EmailStorePstMergeOptions(
        bool overwriteExisting = false,
        string? displayName = null,
        EmailStoreMergeFolderMode folderMode = EmailStoreMergeFolderMode.SeparateSourceRoots,
        bool deduplicate = true,
        EmailSemanticComparisonOptions? deduplicationOptions = null,
        bool continueOnSourceError = true,
        bool continueOnItemError = true,
        bool includeAssociatedItems = true,
        bool includeOrphanedItems = true,
        bool includeSearchFolders = false,
        int maxItems = int.MaxValue,
        int maxFolderCount = 100_000,
        int maxNestedMessageDepth = 32,
        int maxRetries = 2,
        TimeSpan? retryDelay = null,
        int maxIndexRecordsInMemory = 65_536,
        int maxDiagnostics = 10_000,
        IProgress<EmailStoreMergeProgress>? progress = null) {
        if (!Enum.IsDefined(typeof(EmailStoreMergeFolderMode), folderMode)) {
            throw new ArgumentOutOfRangeException(nameof(folderMode));
        }
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxFolderCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxFolderCount));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        if (maxRetries < 0) throw new ArgumentOutOfRangeException(nameof(maxRetries));
        TimeSpan delay = retryDelay ?? TimeSpan.FromMilliseconds(100);
        if (delay < TimeSpan.Zero || delay > TimeSpan.FromMinutes(1)) {
            throw new ArgumentOutOfRangeException(nameof(retryDelay));
        }
        if (maxIndexRecordsInMemory <= 0) throw new ArgumentOutOfRangeException(nameof(maxIndexRecordsInMemory));
        if (maxDiagnostics <= 0) throw new ArgumentOutOfRangeException(nameof(maxDiagnostics));

        OverwriteExisting = overwriteExisting;
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? "OfficeIMO Merged Stores" : displayName!.Trim();
        FolderMode = folderMode;
        Deduplicate = deduplicate;
        DeduplicationOptions = deduplicationOptions ?? CreatePrivateDeduplicationOptions();
        ContinueOnSourceError = continueOnSourceError;
        ContinueOnItemError = continueOnItemError;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        IncludeSearchFolders = includeSearchFolders;
        MaxItems = maxItems;
        MaxFolderCount = maxFolderCount;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        MaxRetries = maxRetries;
        RetryDelay = delay;
        MaxIndexRecordsInMemory = maxIndexRecordsInMemory;
        MaxDiagnostics = maxDiagnostics;
        Progress = progress;
    }

    /// <summary>Whether an existing destination may be atomically replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Destination message-store display name.</summary>
    public string DisplayName { get; }
    /// <summary>Folder mapping policy.</summary>
    public EmailStoreMergeFolderMode FolderMode { get; }
    /// <summary>Whether semantically equivalent items are written once.</summary>
    public bool Deduplicate { get; }
    /// <summary>Semantic policy used by the disk-backed deduplication index.</summary>
    public EmailSemanticComparisonOptions DeduplicationOptions { get; }
    /// <summary>Whether an unreadable source is diagnosed and the remaining sources continue.</summary>
    public bool ContinueOnSourceError { get; }
    /// <summary>Whether a permanently unreadable item is diagnosed and skipped.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Whether folder-associated information items are included.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether orphaned PST/OST items recovered from source indexes are included.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>Whether search-folder results are copied as static folders and items.</summary>
    public bool IncludeSearchFolders { get; }
    /// <summary>Maximum source items inspected across the complete merge.</summary>
    public int MaxItems { get; }
    /// <summary>Maximum destination folders.</summary>
    public int MaxFolderCount { get; }
    /// <summary>Maximum embedded-message nesting depth written.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Maximum retries after transient source-open or item-read I/O failures.</summary>
    public int MaxRetries { get; }
    /// <summary>Delay between transient I/O attempts.</summary>
    public TimeSpan RetryDelay { get; }
    /// <summary>Maximum PST index records sorted in managed memory at once.</summary>
    public int MaxIndexRecordsInMemory { get; }
    /// <summary>Maximum detailed diagnostics retained in the report.</summary>
    public int MaxDiagnostics { get; }
    /// <summary>Optional privacy-safe progress sink.</summary>
    public IProgress<EmailStoreMergeProgress>? Progress { get; }

    private static EmailSemanticComparisonOptions CreatePrivateDeduplicationOptions() {
        var key = new byte[32];
        using (RandomNumberGenerator random = RandomNumberGenerator.Create()) random.GetBytes(key);
        try {
            return new EmailSemanticComparisonOptions(
                EmailSemanticComparisonProfile.Deduplication, digestKey: key);
        } finally {
            Array.Clear(key, 0, key.Length);
        }
    }
}
