namespace OfficeIMO.Email.Store;

/// <summary>One deterministic output partition in a dry-run PST split plan.</summary>
public sealed class EmailStorePstSplitPlanPart {
    internal EmailStorePstSplitPlanPart(int number, string destinationPath,
        IReadOnlyList<EmailStoreItemReference> items, long estimatedBytes,
        long estimatedTargetBytes, bool containsOversizedItem) {
        Number = number;
        DestinationPath = destinationPath;
        Items = items;
        EstimatedBytes = estimatedBytes;
        EstimatedTargetBytes = estimatedTargetBytes;
        ContainsOversizedItem = containsOversizedItem;
    }

    /// <summary>One-based part number.</summary>
    public int Number { get; }
    /// <summary>Absolute final destination path.</summary>
    public string DestinationPath { get; }
    /// <summary>Ordered source references selected for this part.</summary>
    public IReadOnlyList<EmailStoreItemReference> Items { get; }
    /// <summary>Declared/unknown item estimates plus configured per-item overhead.</summary>
    public long EstimatedBytes { get; }
    /// <summary>Configured estimated partition target used for this part.</summary>
    public long EstimatedTargetBytes { get; }
    /// <summary>Whether one item alone exceeded the configured estimated part target.</summary>
    public bool ContainsOversizedItem { get; }
}

/// <summary>Value-free-write dry run for a query- and estimated-size-based PST split.</summary>
public sealed class EmailStorePstSplitPlan {
    private readonly EmailStoreSession _owner;
    internal EmailStorePstSplitPlan(EmailStoreSession owner, string outputBasePath,
        EmailStorePstSplitOptions options, EmailStoreQueryPlan queryPlan,
        int itemsScanned, int matchedItems, bool scanLimitReached, int unknownSizeItems,
        IReadOnlyList<EmailStorePstSplitPlanPart> parts,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        _owner = owner;
        OutputBasePath = outputBasePath;
        Options = options;
        QueryPlan = queryPlan;
        ItemsScanned = itemsScanned;
        MatchedItems = matchedItems;
        ScanLimitReached = scanLimitReached;
        UnknownSizeItems = unknownSizeItems;
        Parts = parts;
        Diagnostics = diagnostics;
    }

    /// <summary>Absolute naming base used to produce <c>.partNNN.pst</c> outputs.</summary>
    public string OutputBasePath { get; }
    /// <summary>Immutable split policy.</summary>
    public EmailStorePstSplitOptions Options { get; }
    /// <summary>Typed selection and ordering plan.</summary>
    public EmailStoreQueryPlan QueryPlan { get; }
    /// <summary>Lightweight references scanned by the query.</summary>
    public int ItemsScanned { get; }
    /// <summary>Total query matches after applying the search-folder policy and before the part-count bound.</summary>
    public int MatchedItems { get; }
    /// <summary>Whether the query scope exceeded its scan bound.</summary>
    public bool ScanLimitReached { get; }
    /// <summary>Selected items whose source summary lacked a declared size.</summary>
    public int UnknownSizeItems { get; }
    /// <summary>Deterministic output partitions.</summary>
    public IReadOnlyList<EmailStorePstSplitPlanPart> Parts { get; }
    /// <summary>Planning, size, conflict, and limit diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Total selected items.</summary>
    public int SelectedItems => Parts.Sum(part => part.Items.Count);
    /// <summary>Total estimated bytes across parts.</summary>
    public long EstimatedBytes => Parts.Sum(part => part.EstimatedBytes);
    /// <summary>Whether execution can start without accepting an incomplete query or current path conflict.</summary>
    public bool IsExecutable => !ScanLimitReached && Parts.Count > 0 &&
        SelectedItems == MatchedItems &&
        !Diagnostics.Any(diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);

    internal void ValidateOwner(EmailStoreSession session) {
        if (!ReferenceEquals(_owner, session)) throw new ArgumentException(
            "The split plan belongs to another store session snapshot.", nameof(session));
    }
}
