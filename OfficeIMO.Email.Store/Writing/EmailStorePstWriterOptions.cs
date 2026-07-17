namespace OfficeIMO.Email.Store;

/// <summary>Controls creation of a new managed Unicode PST file.</summary>
public sealed class EmailStorePstWriterOptions {
    /// <summary>Creates PST writer options.</summary>
    public EmailStorePstWriterOptions(
        string? displayName = null,
        bool overwriteExisting = false,
        bool failOnDataLoss = false,
        int maxFolderCount = 100_000,
        int maxItemCount = int.MaxValue,
        int maxNestedMessageDepth = 32,
        string? checkpointPath = null,
        int checkpointIntervalItems = 1_000,
        int maxIndexRecordsInMemory = 65_536,
        bool retainCheckpointOnDispose = true,
        int maxDiagnostics = 10_000,
        IProgress<EmailStorePstWriteProgress>? progress = null) {
        if (maxFolderCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxFolderCount));
        if (maxItemCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemCount));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        if (checkpointIntervalItems <= 0) throw new ArgumentOutOfRangeException(nameof(checkpointIntervalItems));
        if (maxIndexRecordsInMemory <= 0) throw new ArgumentOutOfRangeException(nameof(maxIndexRecordsInMemory));
        if (maxDiagnostics <= 0) throw new ArgumentOutOfRangeException(nameof(maxDiagnostics));
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? "OfficeIMO Personal Folders" : displayName!;
        OverwriteExisting = overwriteExisting;
        FailOnDataLoss = failOnDataLoss;
        MaxFolderCount = maxFolderCount;
        MaxItemCount = maxItemCount;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        CheckpointPath = string.IsNullOrWhiteSpace(checkpointPath) ? null : Path.GetFullPath(checkpointPath);
        CheckpointIntervalItems = checkpointIntervalItems;
        MaxIndexRecordsInMemory = maxIndexRecordsInMemory;
        RetainCheckpointOnDispose = retainCheckpointOnDispose;
        MaxDiagnostics = maxDiagnostics;
        Progress = progress;
    }

    /// <summary>Display name stored in the message-store object.</summary>
    public string DisplayName { get; }

    /// <summary>Whether an existing destination may be atomically replaced.</summary>
    public bool OverwriteExisting { get; }

    /// <summary>Whether a fidelity warning prevents completion.</summary>
    public bool FailOnDataLoss { get; }

    /// <summary>Maximum folders accepted by one writer.</summary>
    public int MaxFolderCount { get; }

    /// <summary>Maximum top-level items accepted by one writer.</summary>
    public int MaxItemCount { get; }

    /// <summary>Maximum embedded-message nesting depth written.</summary>
    public int MaxNestedMessageDepth { get; }

    /// <summary>Optional durable state file used by <see cref="EmailStorePstWriter.Resume"/>.</summary>
    public string? CheckpointPath { get; }

    /// <summary>Number of newly accepted items between automatic durable checkpoints.</summary>
    public int CheckpointIntervalItems { get; }

    /// <summary>Maximum fixed index records sorted in managed memory at one time.</summary>
    public int MaxIndexRecordsInMemory { get; }

    /// <summary>Whether an incomplete writer preserves or creates its checkpoint when disposed.</summary>
    public bool RetainCheckpointOnDispose { get; }

    /// <summary>Maximum detailed diagnostics retained in managed memory and checkpoints.</summary>
    public int MaxDiagnostics { get; }

    /// <summary>Optional progress sink. Reports contain counts and paths, never message content.</summary>
    public IProgress<EmailStorePstWriteProgress>? Progress { get; }
}
