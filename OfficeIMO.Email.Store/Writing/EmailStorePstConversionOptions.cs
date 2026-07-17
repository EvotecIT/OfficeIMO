namespace OfficeIMO.Email.Store;

/// <summary>Controls safe conversion of a supported mailbox store into a newly created Unicode PST.</summary>
public sealed class EmailStorePstConversionOptions {
    /// <summary>Creates conversion options.</summary>
    public EmailStorePstConversionOptions(
        bool overwriteExisting = false,
        bool failOnDataLoss = false,
        bool continueOnItemError = true,
        bool includeAssociatedItems = true,
        bool includeOrphanedItems = true,
        bool includeSearchFolders = true,
        int maxItems = int.MaxValue,
        int maxNestedMessageDepth = 32,
        string? displayName = null) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        OverwriteExisting = overwriteExisting;
        FailOnDataLoss = failOnDataLoss;
        ContinueOnItemError = continueOnItemError;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        IncludeSearchFolders = includeSearchFolders;
        MaxItems = maxItems;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName;
    }

    /// <summary>Whether an existing destination may be atomically replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Whether any fidelity warning or error blocks completion.</summary>
    public bool FailOnDataLoss { get; }
    /// <summary>Whether unreadable individual items are reported and skipped instead of aborting immediately.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Whether folder-associated information items are copied.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether items recovered from source indexes but absent from contents tables are copied.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>Whether search-folder results are copied as static folders and items.</summary>
    public bool IncludeSearchFolders { get; }
    /// <summary>Maximum source items inspected by one conversion.</summary>
    public int MaxItems { get; }
    /// <summary>Maximum embedded-message nesting depth written.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Optional destination display name. The source display name is used when null.</summary>
    public string? DisplayName { get; }
}
