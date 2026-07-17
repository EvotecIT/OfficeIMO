namespace OfficeIMO.Email.Store;

/// <summary>Merge stages reported without message content or identifiers.</summary>
public enum EmailStoreMergeStage {
    /// <summary>Opening a source.</summary>
    OpeningSource = 0,
    /// <summary>Mapping source folders.</summary>
    MappingFolders = 1,
    /// <summary>Reading, deduplicating, and writing items.</summary>
    WritingItems = 2,
    /// <summary>Finalizing the destination PST.</summary>
    Finalizing = 3,
    /// <summary>The destination has been committed.</summary>
    Completed = 4
}

/// <summary>Privacy-safe aggregate progress for a multi-store PST merge.</summary>
public sealed class EmailStoreMergeProgress {
    internal EmailStoreMergeProgress(EmailStoreMergeStage stage, int sourceIndex, int sourceCount,
        int inspectedItems, int writtenItems, int duplicateItems, int skippedItems) {
        Stage = stage;
        SourceIndex = sourceIndex;
        SourceCount = sourceCount;
        InspectedItems = inspectedItems;
        WrittenItems = writtenItems;
        DuplicateItems = duplicateItems;
        SkippedItems = skippedItems;
    }

    /// <summary>Current operation stage.</summary>
    public EmailStoreMergeStage Stage { get; }
    /// <summary>Zero-based current source index.</summary>
    public int SourceIndex { get; }
    /// <summary>Total configured sources.</summary>
    public int SourceCount { get; }
    /// <summary>Total source items inspected.</summary>
    public int InspectedItems { get; }
    /// <summary>Total items accepted by the PST writer.</summary>
    public int WrittenItems { get; }
    /// <summary>Total semantic duplicates omitted.</summary>
    public int DuplicateItems { get; }
    /// <summary>Total source items skipped after a reported failure or mapping decision.</summary>
    public int SkippedItems { get; }
}
