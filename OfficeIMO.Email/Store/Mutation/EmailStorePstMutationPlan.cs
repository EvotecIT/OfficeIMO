namespace OfficeIMO.Email.Store;

/// <summary>Kind of effective operation in a PST mutation dry run.</summary>
public enum EmailStorePstMutationOperationKind {
    /// <summary>Create a folder.</summary>
    CreateFolder = 0,
    /// <summary>Rename a folder.</summary>
    RenameFolder = 1,
    /// <summary>Move a folder.</summary>
    MoveFolder = 2,
    /// <summary>Delete a folder.</summary>
    DeleteFolder = 3,
    /// <summary>Add a new item.</summary>
    AddItem = 4,
    /// <summary>Copy a source item.</summary>
    CopyItem = 5,
    /// <summary>Replace a complete item.</summary>
    ReplaceItem = 6,
    /// <summary>Patch properties or attachments.</summary>
    PatchItem = 7,
    /// <summary>Move an item or change its associated-content placement.</summary>
    MoveItem = 8,
    /// <summary>Delete an item.</summary>
    DeleteItem = 9
}

/// <summary>One value-free effective operation reported by a PST mutation dry run.</summary>
public sealed class EmailStorePstMutationPlanOperation {
    internal EmailStorePstMutationPlanOperation(EmailStorePstMutationOperationKind kind,
        string entityId, string? destinationId, int changeCount = 0) {
        Kind = kind;
        EntityId = entityId;
        DestinationId = destinationId;
        ChangeCount = changeCount;
    }
    /// <summary>Operation kind.</summary>
    public EmailStorePstMutationOperationKind Kind { get; }
    /// <summary>Source or transaction-local entity identifier.</summary>
    public string EntityId { get; }
    /// <summary>Destination folder or copied-source identifier, when relevant.</summary>
    public string? DestinationId { get; }
    /// <summary>Number of property and attachment changes for a patch operation.</summary>
    public int ChangeCount { get; }
}

/// <summary>
/// Immutable, value-free preview of the effective operations that a PST transaction would commit.
/// Producing this report never writes or replaces an artifact.
/// </summary>
public sealed class EmailStorePstMutationPlan {
    internal EmailStorePstMutationPlan(string sourcePath,
        IReadOnlyList<EmailStorePstMutationPlanOperation> operations,
        int resultingFolderCount, int resultingItemCount, long estimatedRewriteBytes,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        SourcePath = sourcePath;
        Operations = operations;
        ResultingFolderCount = resultingFolderCount;
        ResultingItemCount = resultingItemCount;
        EstimatedRewriteBytes = estimatedRewriteBytes;
        Diagnostics = diagnostics;
    }
    /// <summary>Full source path that remains unchanged by the dry run.</summary>
    public string SourcePath { get; }
    /// <summary>Ordered value-free effective operations.</summary>
    public IReadOnlyList<EmailStorePstMutationPlanOperation> Operations { get; }
    /// <summary>Expected active folder count after commit.</summary>
    public int ResultingFolderCount { get; }
    /// <summary>Expected active item count after commit.</summary>
    public int ResultingItemCount { get; }
    /// <summary>
    /// Conservative number of existing bytes that a full verified rewrite must process. The final output can be
    /// larger or smaller after serialization and is deliberately not presented as an exact size.
    /// </summary>
    public long EstimatedRewriteBytes { get; }
    /// <summary>Source and preflight diagnostics known without writing the staged artifact.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Whether at least one effective operation is staged.</summary>
    public bool HasChanges => Operations.Count > 0;
}

/// <summary>Successful result of one effective operation in a committed PST mutation.</summary>
public sealed class EmailStorePstMutationOperationResult {
    internal EmailStorePstMutationOperationResult(EmailStorePstMutationPlanOperation operation,
        string? destinationEntityId) {
        Operation = operation;
        DestinationEntityId = destinationEntityId;
    }
    /// <summary>Dry-run operation that was applied.</summary>
    public EmailStorePstMutationPlanOperation Operation { get; }
    /// <summary>Identifier in the rewritten PST, or null for deletion.</summary>
    public string? DestinationEntityId { get; }
    /// <summary>True because failed commits throw and never replace the original artifact.</summary>
    public bool IsSuccessful => true;
}

/// <summary>Deterministic handling of an existing same-name folder during a copy.</summary>
public enum EmailStorePstCopyConflictPolicy {
    /// <summary>Reject the copy before staging any copy operation.</summary>
    Fail = 0,
    /// <summary>Create a distinct folder with the same display name.</summary>
    AllowDuplicate = 1
}

/// <summary>Transaction-local mappings produced by a folder or subtree copy.</summary>
public sealed class EmailStorePstFolderCopyResult {
    internal EmailStorePstFolderCopyResult(string rootFolderId,
        IReadOnlyDictionary<string, string> folderIdMap,
        IReadOnlyDictionary<string, string> itemIdMap) {
        RootFolderId = rootFolderId;
        FolderIdMap = folderIdMap;
        ItemIdMap = itemIdMap;
    }
    /// <summary>Transaction-local identifier of the copied root folder.</summary>
    public string RootFolderId { get; }
    /// <summary>Maps copied source folder IDs to transaction-local destination IDs.</summary>
    public IReadOnlyDictionary<string, string> FolderIdMap { get; }
    /// <summary>Maps copied source item IDs to transaction-local destination IDs.</summary>
    public IReadOnlyDictionary<string, string> ItemIdMap { get; }
}
