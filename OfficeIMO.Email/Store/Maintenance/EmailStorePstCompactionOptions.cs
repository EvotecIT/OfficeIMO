using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Selection, estimation, fidelity, and verification policy for a safe PST rewrite compaction.</summary>
public sealed class EmailStorePstCompactionOptions {
    /// <summary>Creates compaction options.</summary>
    public EmailStorePstCompactionOptions(
        bool overwriteExisting = false,
        bool failOnDataLoss = true,
        bool continueOnItemError = false,
        bool includeAssociatedItems = true,
        bool includeOrphanedItems = true,
        bool includeSearchFolders = false,
        int maxItems = 1_000_000,
        int maxNestedMessageDepth = 32,
        long unknownItemEstimateBytes = 1L * 1024L * 1024L,
        long perItemOverheadBytes = 16L * 1024L,
        long fixedPstOverheadBytes = 1L * 1024L * 1024L,
        string? displayName = null,
        EmailSemanticComparisonOptions? verificationOptions = null,
        int maxVerificationIssues = 1000) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        if (unknownItemEstimateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(unknownItemEstimateBytes));
        if (perItemOverheadBytes < 0) throw new ArgumentOutOfRangeException(nameof(perItemOverheadBytes));
        if (fixedPstOverheadBytes < 0) throw new ArgumentOutOfRangeException(nameof(fixedPstOverheadBytes));
        if (maxVerificationIssues <= 0) throw new ArgumentOutOfRangeException(nameof(maxVerificationIssues));
        OverwriteExisting = overwriteExisting;
        FailOnDataLoss = failOnDataLoss;
        ContinueOnItemError = continueOnItemError;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        IncludeSearchFolders = includeSearchFolders;
        MaxItems = maxItems;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        UnknownItemEstimateBytes = unknownItemEstimateBytes;
        PerItemOverheadBytes = perItemOverheadBytes;
        FixedPstOverheadBytes = fixedPstOverheadBytes;
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName;
        VerificationOptions = verificationOptions;
        MaxVerificationIssues = maxVerificationIssues;
    }

    /// <summary>Whether an existing distinct destination may be atomically replaced after verification.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Whether any fidelity warning/error blocks destination commit.</summary>
    public bool FailOnDataLoss { get; }
    /// <summary>Whether an unreadable item is reported and skipped.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Whether folder-associated information items are retained.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether readable source-index orphans are recovered into the compacted output.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>Whether search folders are retained as static folders.</summary>
    public bool IncludeSearchFolders { get; }
    /// <summary>Maximum source references scanned and written.</summary>
    public int MaxItems { get; }
    /// <summary>Maximum embedded-item depth read and written.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Planning estimate for items without a declared source size.</summary>
    public long UnknownItemEstimateBytes { get; }
    /// <summary>Planning estimate for per-item PST rows, nodes, and indexes.</summary>
    public long PerItemOverheadBytes { get; }
    /// <summary>Planning estimate for fixed PST structures and allocation maps.</summary>
    public long FixedPstOverheadBytes { get; }
    /// <summary>Optional destination display name.</summary>
    public string? DisplayName { get; }
    /// <summary>Optional semantic verification policy.</summary>
    public EmailSemanticComparisonOptions? VerificationOptions { get; }
    /// <summary>Maximum mismatch/failure details retained.</summary>
    public int MaxVerificationIssues { get; }
}
