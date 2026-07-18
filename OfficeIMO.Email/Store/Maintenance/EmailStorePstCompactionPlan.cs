namespace OfficeIMO.Email.Store;

/// <summary>Read-only capacity and selection plan for verified PST rewrite compaction.</summary>
public sealed class EmailStorePstCompactionPlan {
    internal EmailStorePstCompactionPlan(string destinationPath,
        EmailStorePstCompactionOptions options, long sourceBytes, int itemsScanned,
        int selectedItems, int associatedItems, int orphanedItems,
        int excludedSearchFolderItems, int unknownSizeItems,
        long estimatedOutputBytes, bool itemLimitReached,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationPath = destinationPath;
        Options = options;
        SourceBytes = sourceBytes;
        ItemsScanned = itemsScanned;
        SelectedItems = selectedItems;
        AssociatedItems = associatedItems;
        OrphanedItems = orphanedItems;
        ExcludedSearchFolderItems = excludedSearchFolderItems;
        UnknownSizeItems = unknownSizeItems;
        EstimatedOutputBytes = estimatedOutputBytes;
        ItemLimitReached = itemLimitReached;
        Diagnostics = diagnostics;
    }

    /// <summary>Absolute distinct output path.</summary>
    public string DestinationPath { get; }
    /// <summary>Immutable compaction policy.</summary>
    public EmailStorePstCompactionOptions Options { get; }
    /// <summary>Validated source file length.</summary>
    public long SourceBytes { get; }
    /// <summary>Source references examined.</summary>
    public int ItemsScanned { get; }
    /// <summary>References selected for the verified rewrite.</summary>
    public int SelectedItems { get; }
    /// <summary>Selected associated items.</summary>
    public int AssociatedItems { get; }
    /// <summary>Selected source-index orphans.</summary>
    public int OrphanedItems { get; }
    /// <summary>Items intentionally excluded with search folders.</summary>
    public int ExcludedSearchFolderItems { get; }
    /// <summary>Selected items whose source size was not declared.</summary>
    public int UnknownSizeItems { get; }
    /// <summary>Planning estimate only; final PST index and allocation overhead is measured after write.</summary>
    public long EstimatedOutputBytes { get; }
    /// <summary>Source bytes minus estimated output bytes. Negative means estimated growth.</summary>
    public long EstimatedReductionBytes => SourceBytes - EstimatedOutputBytes;
    /// <summary>Whether at least one source reference existed beyond the configured bound.</summary>
    public bool ItemLimitReached { get; }
    /// <summary>Format, protection, bounds, destination, source, and selection diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Whether the verified rewrite may start.</summary>
    public bool IsExecutable => SelectedItems > 0 && !ItemLimitReached &&
        !Diagnostics.Any(diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error) &&
        (!Options.FailOnDataLoss || !Diagnostics.Any(diagnostic =>
            diagnostic.Severity != EmailStoreDiagnosticSeverity.Information));
}
