using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Query, size-estimation, verification, and commit policy for PST splitting.</summary>
public sealed class EmailStorePstSplitOptions {
    /// <summary>Creates split options.</summary>
    public EmailStorePstSplitOptions(
        EmailStoreTableQuery? query = null,
        long maxEstimatedBytesPerPart = 3_500_000_000L,
        long unknownItemEstimateBytes = 1L * 1024L * 1024L,
        long perItemOverheadBytes = 16L * 1024L,
        int maxParts = 1000,
        bool overwriteExisting = false,
        bool failOnDataLoss = true,
        bool continueOnItemError = false,
        bool includeSearchFolders = false,
        int maxNestedMessageDepth = 32,
        EmailSemanticComparisonOptions? verificationOptions = null,
        int maxVerificationIssues = 1000) {
        if (maxEstimatedBytesPerPart <= 0) throw new ArgumentOutOfRangeException(nameof(maxEstimatedBytesPerPart));
        if (unknownItemEstimateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(unknownItemEstimateBytes));
        if (perItemOverheadBytes < 0) throw new ArgumentOutOfRangeException(nameof(perItemOverheadBytes));
        if (maxParts <= 0) throw new ArgumentOutOfRangeException(nameof(maxParts));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        if (maxVerificationIssues <= 0) throw new ArgumentOutOfRangeException(nameof(maxVerificationIssues));
        if (query?.ContinuationToken != null) throw new ArgumentException(
            "A split query must start at the beginning; continuation tokens are not accepted.", nameof(query));
        Query = query;
        MaxEstimatedBytesPerPart = maxEstimatedBytesPerPart;
        UnknownItemEstimateBytes = unknownItemEstimateBytes;
        PerItemOverheadBytes = perItemOverheadBytes;
        MaxParts = maxParts;
        OverwriteExisting = overwriteExisting;
        FailOnDataLoss = failOnDataLoss;
        ContinueOnItemError = continueOnItemError;
        IncludeSearchFolders = includeSearchFolders;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        VerificationOptions = verificationOptions;
        MaxVerificationIssues = maxVerificationIssues;
    }

    /// <summary>
    /// Optional typed table query. Null selects all normal, associated, and source-index orphan items.
    /// Query sorting defines deterministic part order.
    /// </summary>
    public EmailStoreTableQuery? Query { get; }
    /// <summary>
    /// Estimated partition target. Final PST index/allocation overhead is reported after write and can differ.
    /// </summary>
    public long MaxEstimatedBytesPerPart { get; }
    /// <summary>Estimate used when the source summary does not declare an item size.</summary>
    public long UnknownItemEstimateBytes { get; }
    /// <summary>Additional planning allowance per item for PST rows, nodes, and indexes.</summary>
    public long PerItemOverheadBytes { get; }
    /// <summary>Maximum planned output parts.</summary>
    public int MaxParts { get; }
    /// <summary>Whether an existing complete output set may be transactionally replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Whether any preservation warning/error or verification mismatch aborts before commit.</summary>
    public bool FailOnDataLoss { get; }
    /// <summary>Whether an unreadable selected item is reported and skipped.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Whether search folders are retained as static folders.</summary>
    public bool IncludeSearchFolders { get; }
    /// <summary>Maximum embedded-item depth read and written.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Optional semantic verification policy.</summary>
    public EmailSemanticComparisonOptions? VerificationOptions { get; }
    /// <summary>Maximum mismatch/failure details retained per part.</summary>
    public int MaxVerificationIssues { get; }
}
