namespace OfficeIMO.Email.Store;

/// <summary>Current phase of incremental PST creation.</summary>
public enum EmailStorePstWriteStage {
    /// <summary>Working files and mandatory folders are being initialized.</summary>
    Initializing = 0,
    /// <summary>A folder was accepted.</summary>
    WritingFolders = 1,
    /// <summary>An item was accepted and spooled into the working PST.</summary>
    WritingItems = 2,
    /// <summary>A resumable checkpoint is being committed.</summary>
    Checkpointing = 3,
    /// <summary>Folder tables and PST indexes are being finalized.</summary>
    Finalizing = 4,
    /// <summary>The destination PST was atomically committed.</summary>
    Completed = 5
}

/// <summary>Bounded progress snapshot for a PST write or conversion.</summary>
public sealed class EmailStorePstWriteProgress {
    internal EmailStorePstWriteProgress(EmailStorePstWriteStage stage,
        int folderCount, int itemCount, long workingBytes, string? checkpointPath) {
        Stage = stage;
        FolderCount = folderCount;
        ItemCount = itemCount;
        WorkingBytes = workingBytes;
        CheckpointPath = checkpointPath;
    }
    /// <summary>Current writer stage.</summary>
    public EmailStorePstWriteStage Stage { get; }
    /// <summary>Number of user folders accepted.</summary>
    public int FolderCount { get; }
    /// <summary>Number of top-level items accepted.</summary>
    public int ItemCount { get; }
    /// <summary>Current working PST length, excluding disk-backed indexes.</summary>
    public long WorkingBytes { get; }
    /// <summary>Configured durable checkpoint path, when enabled.</summary>
    public string? CheckpointPath { get; }
}
