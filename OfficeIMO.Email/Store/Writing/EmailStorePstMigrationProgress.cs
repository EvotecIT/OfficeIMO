namespace OfficeIMO.Email.Store;

/// <summary>Privacy-safe progress for a resumable Store-to-PST migration.</summary>
public sealed class EmailStorePstMigrationProgress {
    internal EmailStorePstMigrationProgress(int inspectedItems, int convertedItems,
        int skippedItems, string? checkpointPath) {
        InspectedItems = inspectedItems;
        ConvertedItems = convertedItems;
        SkippedItems = skippedItems;
        CheckpointPath = checkpointPath;
    }

    /// <summary>Selected source references inspected so far.</summary>
    public int InspectedItems { get; }
    /// <summary>Items durably accepted by the destination writer so far.</summary>
    public int ConvertedItems { get; }
    /// <summary>Items skipped under the configured partial-result policy.</summary>
    public int SkippedItems { get; }
    /// <summary>Configured durable checkpoint path, or null when resume is disabled.</summary>
    public string? CheckpointPath { get; }
}
